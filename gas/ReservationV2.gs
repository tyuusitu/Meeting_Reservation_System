/**
 * 予約登録 v2 の静的設定キャッシュキーです。
 */
const RESERVATION_V2_STATIC_CACHE_KEY = 'reservation-v2-static-context';
const RESERVATION_V2_STATIC_CACHE_SECONDS = 300;
const RESERVATION_V2_LOCK_WAIT_MS = 5000;
const RESERVATION_V2_SLOT_PROPERTY_PREFIX = 'RESERVATION_V2_SLOT_';
const RESERVATION_V2_REQUEST_PROPERTY_PREFIX = 'RESERVATION_V2_REQUEST_';
const RESERVATION_V2_MODE_PROPERTY_KEY = 'RESERVATION_V2_MODE';
const RESERVATION_V2_ENABLED_PROPERTY_KEY = 'RESERVATION_V2_ENABLED';
const RESERVATION_V2_INDEX_READY_PROPERTY_KEY = 'RESERVATION_V2_INDEX_READY';
const RESERVATION_V2_INDEX_BUILT_AT_PROPERTY_KEY = 'RESERVATION_V2_INDEX_BUILT_AT';
const RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY = 'RESERVATION_V2_INDEX_STALE_REASON';
const RESERVATION_V2_LAST_PRUNED_AT_PROPERTY_KEY = 'RESERVATION_V2_LAST_PRUNED_AT';
const RESERVATION_V2_PRUNE_INTERVAL_MS = 60 * 60 * 1000;
const RESERVATION_V2_REQUEST_RETENTION_MS = 72 * 60 * 60 * 1000;
const RESERVATION_V2_MODES = {
  legacy: 'LEGACY',
  preparing: 'PREPARING',
  v2: 'V2',
};

/** 現在の予約 API 運用モードを返します。未設定時は必ず旧方式です。 */
function getReservationV2Mode_() {
  const properties = PropertiesService.getScriptProperties();
  const mode = normalizeString_(properties.getProperty(RESERVATION_V2_MODE_PROPERTY_KEY)).toUpperCase();
  if (mode === RESERVATION_V2_MODES.preparing || mode === RESERVATION_V2_MODES.v2) return mode;
  if (properties.getProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY) === 'TRUE') return RESERVATION_V2_MODES.v2;
  return RESERVATION_V2_MODES.legacy;
}

/** 新予約方式が有効かどうかを返します。 */
function isReservationV2Enabled_() {
  return getReservationV2Mode_() === RESERVATION_V2_MODES.v2;
}

/** 競合確認用索引が利用可能かどうかを返します。 */
function isReservationV2IndexReady_() {
  return PropertiesService.getScriptProperties().getProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY) === 'TRUE';
}

/** 切替中または索引未準備時の安全な応答を返します。 */
function createReservationV2MaintenanceResult_(requestId) {
  return {
    ok: false,
    state: 'maintenance',
    code: 'RESERVATION_V2_NOT_READY',
    request_id: normalizeString_(requestId),
    retry_after_ms: 3000,
    error: '予約機能を安全に切り替えています。少し待ってから同じ内容で再度お試しください。',
  };
}

/**
 * 短時間ロック、冪等キー、処理中状態を使って予約を登録します。
 *
 * <p>ロック中は最新予約の取得、競合確認、処理中行の確保だけを行います。
 * Google カレンダー操作はロック解除後に実行します。</p>
 *
 * @param {Object} payload 予約登録リクエスト。
 * @return {Object} 登録結果と段階別処理時間。
 */
function insertReservationsV2_(payload) {
  const totalStartedAt = Date.now();
  const metrics = {};
  const input = payload || {};
  const requestId = normalizeString_(input.request_id || input.requestId) || Utilities.getUuid();

  if (!isReservationV2IndexReady_()) {
    return createReservationV2MaintenanceResult_(requestId);
  }

  const staticStartedAt = Date.now();
  const context = getReservationStaticContextV2_();
  metrics.static_context_ms = Date.now() - staticStartedAt;

  const validation = validateReservationRequestV2_(input, context);
  if (!validation.ok) {
    metrics.total_ms = Date.now() - totalStartedAt;
    return {
      ok: false,
      state: 'invalid',
      request_id: requestId,
      errors: validation.errors,
      error: validation.errors.map((item) => item.message).join('\n'),
      metrics,
    };
  }

  const requestHash = createReservationRequestHashV2_(validation.common, validation.reservations);
  const lock = LockService.getScriptLock();
  const lockWaitStartedAt = Date.now();
  if (!lock.tryLock(RESERVATION_V2_LOCK_WAIT_MS)) {
    metrics.lock_wait_ms = Date.now() - lockWaitStartedAt;
    metrics.total_ms = Date.now() - totalStartedAt;
    return {
      ok: false,
      state: 'busy',
      code: 'RESERVATION_BUSY',
      request_id: requestId,
      retry_after_ms: 500,
      error: '予約が混み合っています。同じリクエストIDで再確認します。',
      metrics,
    };
  }

  metrics.lock_wait_ms = Date.now() - lockWaitStartedAt;
  let sheet;
  let startRow = 0;
  let pendingRows = [];
  let registeredReservations = [];
  try {
    if (getReservationV2Mode_() !== RESERVATION_V2_MODES.v2 || !isReservationV2IndexReady_()) {
      return createReservationV2MaintenanceResult_(requestId);
    }
    const properties = PropertiesService.getScriptProperties();
    pruneReservationV2Properties_(properties);
    const readStartedAt = Date.now();
    const requestPropertyKey = getReservationRequestPropertyKeyV2_(requestId);
    const hasRequestProperty = Boolean(properties.getProperty(requestPropertyKey));
    const previousRecord = getReservationRequestRecordV2_(properties, requestId);
    metrics.reservation_read_ms = Date.now() - readStartedAt;

    if (hasRequestProperty && !previousRecord) {
      markReservationV2IndexStale_('request_id 状態の破損を検出しました。');
      return createReservationV2MaintenanceResult_(requestId);
    }

    if (previousRecord) {
      metrics.lock_held_ms = Date.now() - lockWaitStartedAt - metrics.lock_wait_ms;
      metrics.total_ms = Date.now() - totalStartedAt;
      return buildPreviousReservationResultV2_(requestId, requestHash, previousRecord.rows, metrics);
    }

    const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    const pending = createPendingReservationRowsV2_(
      requestId,
      requestHash,
      validation.common,
      validation.reservations,
      context.roomMap,
      now
    );
    pendingRows = pending.rows;
    registeredReservations = pending.reservations;
    const claimPlan = createReservationClaimPlanV2_(
      properties,
      requestId,
      validation.reservations,
      registeredReservations,
      context.roomMap
    );
    if (claimPlan.errors.length > 0) {
      metrics.lock_held_ms = Date.now() - lockWaitStartedAt - metrics.lock_wait_ms;
      metrics.total_ms = Date.now() - totalStartedAt;
      return {
        ok: false,
        state: 'conflict',
        code: 'RESERVATION_CONFLICT',
        request_id: requestId,
        errors: claimPlan.errors,
        error: claimPlan.errors.map((item) => item.message).join('\n'),
        metrics,
      };
    }

    const requestRows = pendingRows.map(reservationRowArrayToObjectV2_);
    const propertyUpdates = Object.assign({}, claimPlan.propertyUpdates);
    propertyUpdates[getReservationRequestPropertyKeyV2_(requestId)] = JSON.stringify(
      createReservationRequestRecordV2_(requestId, requestHash, requestRows, now)
    );
    const writeStartedAt = Date.now();
    try {
      properties.setProperties(propertyUpdates, false);
      sheet = getSheet_(SHEET_NAMES.reservations);
      startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, pendingRows.length, pendingRows[0].length).setValues(pendingRows);
      SpreadsheetApp.flush();
    } catch (error) {
      restoreReservationPropertiesV2_(properties, claimPlan.originalValues);
      properties.deleteProperty(getReservationRequestPropertyKeyV2_(requestId));
      metrics.slot_write_ms = Date.now() - writeStartedAt;
      metrics.lock_held_ms = Date.now() - lockWaitStartedAt - metrics.lock_wait_ms;
      metrics.total_ms = Date.now() - totalStartedAt;
      writeOperationLogFastV2_(
        '予約登録',
        validation.common.user_name || '利用者',
        requestId,
        '予約枠の確保に失敗しました。',
        '失敗',
        `${error.message} metrics=${JSON.stringify(metrics)}`
      );
      return {
        ok: false,
        state: 'failed',
        code: 'RESERVATION_PREPARE_FAILED',
        request_id: requestId,
        error: '予約枠を保存できませんでした。時間をおいて再度お試しください。',
        metrics,
      };
    }
    metrics.slot_write_ms = Date.now() - writeStartedAt;
    metrics.lock_held_ms = Date.now() - lockWaitStartedAt - metrics.lock_wait_ms;
  } finally {
    lock.releaseLock();
  }

  const createdEvents = [];
  try {
    const calendarStartedAt = Date.now();
    registeredReservations.forEach((reservation, index) => {
      const room = context.roomMap[reservation.room_id];
      const startDate = createDateTime_(reservation.usage_date, reservation.start_time);
      const endDate = createDateTime_(reservation.usage_date, reservation.end_time);
      const eventTitle = validation.common.organization_name
        ? `【${validation.common.organization_name}】${reservation.meeting_name}`
        : reservation.meeting_name;
      const eventDescription = [
        validation.common.organization_name ? `団体名：${validation.common.organization_name}` : null,
        `会議名：${reservation.meeting_name}`,
        validation.common.user_name ? `お名前：${validation.common.user_name}` : null,
        `会議室：${room.room_name}`,
        `予約ID：${reservation.reservation_id}`,
        `リクエストID：${requestId}`,
      ].filter(Boolean).join('\n');
      const calendarEvents = createCalendarEventsForReservation_(
        room,
        context.roomMap,
        eventTitle,
        startDate,
        endDate,
        eventDescription
      );
      calendarEvents.forEach((calendarEvent) => createdEvents.push(calendarEvent.event));
      const eventRefs = serializeCalendarEventRefs_(calendarEvents);
      registeredReservations[index].calendar_event_id = eventRefs;
      setReservationRowValueV2_(pendingRows[index], 'calendar_event_id', eventRefs);
    });
    metrics.calendar_ms = Date.now() - calendarStartedAt;

    const finalizedAt = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    pendingRows.forEach((row) => {
      setReservationRowValueV2_(row, 'status', RESERVATION_STATUS.active);
      setReservationRowValueV2_(row, 'updated_at', finalizedAt);
      setReservationRowValueV2_(row, 'error_message', '');
    });
    const finalizeStartedAt = Date.now();
    sheet.getRange(startRow, 1, pendingRows.length, pendingRows[0].length).setValues(pendingRows);
    metrics.finalize_ms = Date.now() - finalizeStartedAt;
    setReservationRequestRecordV2_(requestId, requestHash, pendingRows, finalizedAt);
    metrics.total_ms = Date.now() - totalStartedAt;

    writeOperationLogFastV2_(
      '予約登録',
      validation.common.user_name || '利用者',
      requestId,
      `${registeredReservations.length}件の予約を登録しました。 metrics=${JSON.stringify(metrics)}`,
      '成功',
      ''
    );

    return {
      ok: true,
      state: 'active',
      request_id: requestId,
      reservations: registeredReservations.map(toReservationResponseV2_),
      metrics,
    };
  } catch (error) {
    createdEvents.forEach((calendarEvent) => {
      try {
        calendarEvent.deleteEvent();
      } catch (_) {
      }
    });
    const failedAt = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    pendingRows.forEach((row) => {
      setReservationRowValueV2_(row, 'calendar_event_id', '');
      setReservationRowValueV2_(row, 'status', RESERVATION_STATUS.failed);
      setReservationRowValueV2_(row, 'updated_at', failedAt);
      setReservationRowValueV2_(row, 'error_message', error.message);
    });
    try {
      sheet.getRange(startRow, 1, pendingRows.length, pendingRows[0].length).setValues(pendingRows);
    } catch (_) {
    }
    const claimsReleased = releaseReservationClaimsV2_(requestId, validation.reservations, context.roomMap);
    if (!claimsReleased) {
      markReservationV2IndexStale_('予約登録失敗後の枠解放でロックを取得できませんでした。');
    }
    setReservationRequestRecordV2_(requestId, requestHash, pendingRows, failedAt);
    metrics.total_ms = Date.now() - totalStartedAt;
    writeOperationLogFastV2_(
      '予約登録',
      validation.common.user_name || '利用者',
      requestId,
      '予約登録に失敗しました。',
      '失敗',
      `${error.message} metrics=${JSON.stringify(metrics)}`
    );
    return {
      ok: false,
      state: 'failed',
      code: 'CALENDAR_OR_FINALIZE_FAILED',
      request_id: requestId,
      error: error.message,
      metrics,
    };
  }
}

/**
 * request_id に紐づく現在状態を返します。
 *
 * @param {Object} params GET パラメータ。
 * @return {Object} 現在状態。
 */
function handleGetReservationStatusV2_(params) {
  const requestId = normalizeString_(params && (params.request_id || params.requestId));
  if (!requestId) {
    return { ok: false, state: 'invalid', error: 'リクエストIDが必要です。' };
  }
  const record = getReservationRequestRecordV2_(PropertiesService.getScriptProperties(), requestId);
  if (record) {
    return buildPreviousReservationResultV2_(requestId, record.request_hash, record.rows, {});
  }
  if (!normalizeBoolean_(params && params.deep_check)) {
    return { ok: true, state: 'not_found', request_id: requestId, reservations: [] };
  }
  const rows = selectSheetObjects_(SHEET_NAMES.reservations)
    .map(normalizeReservationRow_)
    .filter((reservation) => reservation.request_id === requestId);
  if (rows.length === 0) {
    return { ok: true, state: 'not_found', request_id: requestId, reservations: [] };
  }
  return buildPreviousReservationResultV2_(requestId, rows[0].request_hash, rows, {});
}

/** request_id に対応する Script Properties のキーを返します。 */
function getReservationRequestPropertyKeyV2_(requestId) {
  return `${RESERVATION_V2_REQUEST_PROPERTY_PREFIX}${createReservationPropertyDigestV2_(requestId).slice(0, 40)}`;
}

/** 利用日と実会議室に対応する予約枠索引キーを返します。 */
function getReservationSlotPropertyKeyV2_(usageDate, effectiveRoomId) {
  const digest = createReservationPropertyDigestV2_(`${usageDate}|${effectiveRoomId}`).slice(0, 40);
  return `${RESERVATION_V2_SLOT_PROPERTY_PREFIX}${digest}`;
}

/** Script Properties のキー生成に使う SHA-256 文字列を返します。 */
function createReservationPropertyDigestV2_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalizeString_(value),
    Utilities.Charset.UTF_8
  ).map((byte) => (`0${(byte + 256).toString(16)}`).slice(-2)).join('');
}

/** request_id の処理状態を Script Properties から取得します。 */
function getReservationRequestRecordV2_(properties, requestId) {
  const text = properties.getProperty(getReservationRequestPropertyKeyV2_(requestId));
  if (!text) return null;
  try {
    const record = JSON.parse(text);
    if (record.request_id !== requestId || !Array.isArray(record.rows)) return null;
    record.rows = record.rows.map(normalizeReservationRow_);
    return record;
  } catch (_) {
    return null;
  }
}

/** 予約シート用配列を内部オブジェクトへ変換します。 */
function reservationRowArrayToObjectV2_(row) {
  const object = {};
  SHEET_COLUMN_KEYS.reservations.forEach((key, index) => {
    object[key] = row[index];
  });
  return normalizeReservationRow_(object);
}

/** request_id の処理状態を Script Properties へ保存します。 */
function setReservationRequestRecordV2_(requestId, requestHash, rows, updatedAt) {
  PropertiesService.getScriptProperties().setProperty(
    getReservationRequestPropertyKeyV2_(requestId),
    JSON.stringify(createReservationRequestRecordV2_(requestId, requestHash, rows, updatedAt))
  );
}

/** Script Properties の容量を抑えた request_id 状態を作ります。 */
function createReservationRequestRecordV2_(requestId, requestHash, rows, updatedAt) {
  const normalizedRows = rows.map((row) => Array.isArray(row)
    ? reservationRowArrayToObjectV2_(row)
    : normalizeReservationRow_(row));
  return {
    request_id: requestId,
    request_hash: requestHash,
    rows: normalizedRows.map((row) => ({
      reservation_id: row.reservation_id,
      meeting_name: row.meeting_name,
      room_id: row.room_id,
      room_name: row.room_name,
      usage_date: row.usage_date,
      start_time: row.start_time,
      end_time: row.end_time,
      calendar_event_id: row.calendar_event_id,
      status: row.status,
      request_id: requestId,
      request_hash: requestHash,
      updated_at: row.updated_at || updatedAt,
      error_message: row.error_message,
    })),
    updated_at: updatedAt,
  };
}

/**
 * スプレッドシートを読まず、日付・実会議室別の軽量索引で競合確認と枠確保を準備します。
 */
function createReservationClaimPlanV2_(properties, requestId, reservations, registeredReservations, roomMap) {
  const originalValues = {};
  const claimsByKey = {};
  const errors = [];

  reservations.forEach((reservation, index) => {
    const effectiveRoomIds = getEffectiveRoomIds_(reservation.room_id, roomMap);
    let conflict = null;
    effectiveRoomIds.some((effectiveRoomId) => {
      const key = getReservationSlotPropertyKeyV2_(reservation.usage_date, effectiveRoomId);
      if (!Object.prototype.hasOwnProperty.call(claimsByKey, key)) {
        const original = properties.getProperty(key);
        originalValues[key] = original;
        try {
          claimsByKey[key] = original ? JSON.parse(original) : [];
          if (!Array.isArray(claimsByKey[key])) throw new Error('invalid reservation claims');
        } catch (_) {
          markReservationV2IndexStale_('予約枠索引の破損を検出しました。');
          throw new Error('予約枠索引が破損しています。rebuildReservationV2Indexes() を実行してください。');
        }
      }
      conflict = claimsByKey[key].find((claim) =>
        claim.request_id !== requestId
        && hasTimeOverlap_(claim.start_time, claim.end_time, reservation.start_time, reservation.end_time)
      );
      return Boolean(conflict);
    });

    if (conflict) {
      errors.push({
        index,
        field: 'room_id',
        message: `${conflict.room_name || reservation.room_id}は${reservation.start_time}-${reservation.end_time}に既存予約があります。`,
      });
      return;
    }

    effectiveRoomIds.forEach((effectiveRoomId) => {
      const key = getReservationSlotPropertyKeyV2_(reservation.usage_date, effectiveRoomId);
      claimsByKey[key].push({
        request_id: requestId,
        reservation_id: registeredReservations[index].reservation_id,
        room_id: reservation.room_id,
        room_name: roomMap[reservation.room_id].room_name,
        usage_date: reservation.usage_date,
        start_time: reservation.start_time,
        end_time: reservation.end_time,
      });
    });
  });

  const propertyUpdates = {};
  Object.keys(claimsByKey).forEach((key) => {
    propertyUpdates[key] = JSON.stringify(claimsByKey[key]);
  });
  return { errors, propertyUpdates, originalValues };
}

/** 枠確保前の Script Properties 値へ戻します。 */
function restoreReservationPropertiesV2_(properties, originalValues) {
  Object.keys(originalValues).forEach((key) => {
    if (originalValues[key] === null) {
      properties.deleteProperty(key);
    } else {
      properties.setProperty(key, originalValues[key]);
    }
  });
}

/** 失敗した request_id、または指定予約IDの枠確保を索引から解放します。 */
function releaseReservationClaimsV2_(requestId, reservations, roomMap, reservationIds) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return false;
  try {
    const properties = PropertiesService.getScriptProperties();
    const targetIds = Array.isArray(reservationIds) ? reservationIds.filter(Boolean) : [];
    const keys = [];
    let corrupted = false;
    reservations.forEach((reservation) => {
      getEffectiveRoomIds_(reservation.room_id, roomMap).forEach((effectiveRoomId) => {
        keys.push(getReservationSlotPropertyKeyV2_(reservation.usage_date, effectiveRoomId));
      });
    });
    Array.from(new Set(keys)).forEach((key) => {
      const text = properties.getProperty(key);
      if (!text) return;
      let claims = [];
      try {
        claims = JSON.parse(text).filter((claim) => targetIds.length > 0
          ? targetIds.indexOf(claim.reservation_id) === -1
          : claim.request_id !== requestId);
      } catch (_) {
        markReservationV2IndexStale_('予約枠索引の破損を検出しました。');
        corrupted = true;
        return;
      }
      if (claims.length > 0) properties.setProperty(key, JSON.stringify(claims));
      else properties.deleteProperty(key);
    });
    return !corrupted;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 完了済み request_id と過去日の枠索引を定期削除し、Script Properties の上限超過を防ぎます。
 */
function pruneReservationV2Properties_(properties) {
  const now = Date.now();
  const lastPrunedAt = Number(properties.getProperty(RESERVATION_V2_LAST_PRUNED_AT_PROPERTY_KEY) || 0);
  if (now - lastPrunedAt < RESERVATION_V2_PRUNE_INTERVAL_MS) return;

  const today = formatDate_(new Date(), 'yyyy-MM-dd');
  const allProperties = properties.getProperties();
  Object.keys(allProperties).forEach((key) => {
    if (key.indexOf(RESERVATION_V2_REQUEST_PROPERTY_PREFIX) === 0) {
      try {
        const record = JSON.parse(allProperties[key]);
        const rows = Array.isArray(record.rows) ? record.rows.map(normalizeReservationRow_) : [];
        const hasProcessing = rows.some((row) => row.status === RESERVATION_STATUS.processing);
        const updatedAt = parseDateTime_(record.updated_at);
        if (!hasProcessing && updatedAt && now - updatedAt.getTime() > RESERVATION_V2_REQUEST_RETENTION_MS) {
          properties.deleteProperty(key);
        }
      } catch (_) {
        markReservationV2IndexStale_('request_id 状態の破損を検出しました。');
        throw new Error('request_id 状態が破損しています。rebuildReservationV2Indexes() を実行してください。');
      }
      return;
    }
    if (key.indexOf(RESERVATION_V2_SLOT_PROPERTY_PREFIX) === 0) {
      try {
        const claims = JSON.parse(allProperties[key]);
        if (!Array.isArray(claims)) throw new Error('invalid reservation claims');
        const retained = claims.filter((claim) => normalizeString_(claim.usage_date) >= today);
        if (retained.length === 0) properties.deleteProperty(key);
        else if (retained.length !== claims.length) properties.setProperty(key, JSON.stringify(retained));
      } catch (_) {
        markReservationV2IndexStale_('予約枠索引の破損を検出しました。');
        throw new Error('予約枠索引が破損しています。rebuildReservationV2Indexes() を実行してください。');
      }
    }
  });
  properties.setProperty(RESERVATION_V2_LAST_PRUNED_AT_PROPERTY_KEY, String(now));
}

/**
 * 管理者が取消を確定した予約を競合索引と request_id 状態へ同期します。
 * 同期できなかった場合は索引を利用不可にし、二重予約を防ぐ側へ倒します。
 */
function synchronizeCancelledReservationsV2_(reservations) {
  if (!isReservationV2Enabled_() || !isReservationV2IndexReady_()) return true;
  const normalizedRows = (reservations || []).map(normalizeReservationRow_).filter((row) => row.reservation_id);
  if (normalizedRows.length === 0) return true;
  const roomMap = createRoomMap_(selectActiveRooms_());
  const released = releaseReservationClaimsV2_(
    '',
    normalizedRows,
    roomMap,
    normalizedRows.map((row) => row.reservation_id)
  );
  if (!released) {
    markReservationV2IndexStale_('予約取消後の索引同期でロックを取得できませんでした。');
    return false;
  }

  const properties = PropertiesService.getScriptProperties();
  const cancelledIds = normalizedRows.reduce((map, row) => {
    map[row.reservation_id] = true;
    return map;
  }, {});
  const requestIds = Array.from(new Set(normalizedRows.map((row) => row.request_id).filter(Boolean)));
  const updatedAt = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  requestIds.forEach((requestId) => {
    const record = getReservationRequestRecordV2_(properties, requestId);
    if (!record) return;
    record.rows = record.rows.map((row) => {
      if (!cancelledIds[row.reservation_id]) return row;
      row.status = RESERVATION_STATUS.cancelled;
      row.updated_at = updatedAt;
      return row;
    });
    setReservationRequestRecordV2_(requestId, record.request_hash, record.rows, updatedAt);
  });
  return true;
}

/** 索引を無効化し、新方式を安全停止状態へ移します。 */
function markReservationV2IndexStale_(reason) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
  properties.setProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY, normalizeString_(reason));
}

/**
 * 既存予約から軽量競合索引と冪等状態を再構築します。
 * 初回導入時、および予約シートを手動修正した後に実行します。
 */
function rebuildReservationV2Indexes() {
  const properties = PropertiesService.getScriptProperties();
  const previousMode = getReservationV2Mode_();
  properties.setProperty(RESERVATION_V2_MODE_PROPERTY_KEY, RESERVATION_V2_MODES.preparing);
  properties.setProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY, 'FALSE');
  properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    restoreReservationV2Mode_(previousMode);
    throw new Error('予約処理が実行中のため索引を再構築できません。少し待ってから再実行してください。');
  }
  try {
    ensureReservationV2SheetSchema_();
    const processing = selectSheetObjects_(SHEET_NAMES.reservations)
      .map(normalizeReservationRow_)
      .filter((row) => row.status === RESERVATION_STATUS.processing);
    if (processing.length > 0) {
      throw new Error(`処理中の予約が${processing.length}件あるため索引を再構築できません。完了後に再実行してください。`);
    }
    const result = rebuildReservationV2IndexesUnlocked_();
    markReservationV2IndexReady_();
    restoreReservationV2Mode_(previousMode);
    return Object.assign(result, getReservationV2ReleaseStatus());
  } catch (error) {
    properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
    properties.setProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY, error.message);
    restoreReservationV2Mode_(previousMode);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/** ロック取得済みの呼び出し元から予約索引を再構築します。 */
function rebuildReservationV2IndexesUnlocked_() {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  Object.keys(allProperties).forEach((key) => {
    if (key.indexOf(RESERVATION_V2_SLOT_PROPERTY_PREFIX) === 0
      || key.indexOf(RESERVATION_V2_REQUEST_PROPERTY_PREFIX) === 0) {
      properties.deleteProperty(key);
    }
  });

  const roomMap = createRoomMap_(selectActiveRooms_());
  const reservations = selectSheetObjects_(SHEET_NAMES.reservations).map(normalizeReservationRow_);
  const propertyUpdates = {};
  const claimsByKey = {};
  const rowsByRequestId = {};
  const today = formatDate_(new Date(), 'yyyy-MM-dd');
  const requestRetentionStartedAt = Date.now() - RESERVATION_V2_REQUEST_RETENTION_MS;

  reservations.forEach((reservation) => {
    const requestUpdatedAt = parseDateTime_(reservation.updated_at || reservation.created_at);
    if (reservation.request_id && (
      reservation.status === RESERVATION_STATUS.processing
      || (requestUpdatedAt && requestUpdatedAt.getTime() >= requestRetentionStartedAt)
    )) {
      if (!rowsByRequestId[reservation.request_id]) rowsByRequestId[reservation.request_id] = [];
      rowsByRequestId[reservation.request_id].push(reservation);
    }
    if (reservation.status !== RESERVATION_STATUS.active
      && reservation.status !== RESERVATION_STATUS.processing
      && reservation.status !== RESERVATION_STATUS.cancelRequested) return;
    if (reservation.usage_date < today) return;
    const claimRequestId = reservation.request_id || `legacy:${reservation.reservation_id}`;
    getEffectiveRoomIds_(reservation.room_id, roomMap).forEach((effectiveRoomId) => {
      const key = getReservationSlotPropertyKeyV2_(reservation.usage_date, effectiveRoomId);
      if (!claimsByKey[key]) claimsByKey[key] = [];
      claimsByKey[key].push({
        request_id: claimRequestId,
        reservation_id: reservation.reservation_id,
        room_id: reservation.room_id,
        room_name: reservation.room_name,
        usage_date: reservation.usage_date,
        start_time: reservation.start_time,
        end_time: reservation.end_time,
      });
    });
  });

  Object.keys(claimsByKey).forEach((key) => {
    propertyUpdates[key] = JSON.stringify(claimsByKey[key]);
  });
  Object.keys(rowsByRequestId).forEach((requestId) => {
    const rows = rowsByRequestId[requestId];
    propertyUpdates[getReservationRequestPropertyKeyV2_(requestId)] = JSON.stringify(
      createReservationRequestRecordV2_(
        requestId,
        rows[0].request_hash,
        rows,
        rows[0].updated_at || rows[0].created_at
      )
    );
  });
  if (Object.keys(propertyUpdates).length > 0) properties.setProperties(propertyUpdates, false);
  return {
    ok: true,
    reservation_count: reservations.length,
    slot_property_count: Object.keys(claimsByKey).length,
    request_property_count: Object.keys(rowsByRequestId).length,
  };
}

/** 予約シートへ v2 用の不足列を追加します。予約データは変更しません。 */
function ensureReservationV2SheetSchema_() {
  ensureSheetAndHeader_(getSpreadsheet_(), 'reservations');
  applyFormattingForSheet_(SHEET_NAMES.reservations, ['usage_date'], ['created_at', 'processing_started_at', 'updated_at']);
}

/** 索引準備済みの時刻と状態を保存します。 */
function markReservationV2IndexReady_() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'TRUE');
  properties.setProperty(
    RESERVATION_V2_INDEX_BUILT_AT_PROPERTY_KEY,
    formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss")
  );
  properties.deleteProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY);
}

/** 指定モードを復元します。 */
function restoreReservationV2Mode_(mode) {
  const properties = PropertiesService.getScriptProperties();
  const normalizedMode = mode === RESERVATION_V2_MODES.v2
    ? RESERVATION_V2_MODES.v2
    : RESERVATION_V2_MODES.legacy;
  properties.setProperty(RESERVATION_V2_MODE_PROPERTY_KEY, normalizedMode);
  properties.setProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY, normalizedMode === RESERVATION_V2_MODES.v2 ? 'TRUE' : 'FALSE');
}

/**
 * 本番切替前のスキーマ準備だけを行います。予約方式は旧方式のままです。
 * Apps Script エディタから手動実行します。
 */
function prepareReservationV2Release() {
  ensureReservationV2SheetSchema_();
  invalidateReservationStaticCacheV2_();
  return getReservationV2ReleaseStatus();
}

/**
 * 受付を一時停止し、最新予約から索引を作ってから v2 を有効化します。
 * 索引構築とモード切替は同じ Script Lock 内で行います。
 */
function enableReservationV2Release() {
  const properties = PropertiesService.getScriptProperties();
  if (isReservationV2Enabled_() && isReservationV2IndexReady_()) return getReservationV2ReleaseStatus();
  properties.setProperty(RESERVATION_V2_MODE_PROPERTY_KEY, RESERVATION_V2_MODES.preparing);
  properties.setProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY, 'FALSE');
  properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    restoreReservationV2Mode_(RESERVATION_V2_MODES.legacy);
    throw new Error('予約処理が実行中のため切り替えできません。少し待ってから再実行してください。');
  }
  try {
    ensureReservationV2SheetSchema_();
    const processing = selectSheetObjects_(SHEET_NAMES.reservations)
      .map(normalizeReservationRow_)
      .filter((row) => row.status === RESERVATION_STATUS.processing);
    if (processing.length > 0) {
      throw new Error(`処理中の予約が${processing.length}件あります。状態を確認してから再実行してください。`);
    }
    const result = rebuildReservationV2IndexesUnlocked_();
    markReservationV2IndexReady_();
    restoreReservationV2Mode_(RESERVATION_V2_MODES.v2);
    return Object.assign(result, getReservationV2ReleaseStatus());
  } catch (error) {
    properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
    properties.setProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY, error.message);
    restoreReservationV2Mode_(RESERVATION_V2_MODES.legacy);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * v2 を旧方式へ戻します。処理中行がある間は安全のため切替を拒否します。
 */
function disableReservationV2Release() {
  const properties = PropertiesService.getScriptProperties();
  const previousMode = getReservationV2Mode_();
  properties.setProperty(RESERVATION_V2_MODE_PROPERTY_KEY, RESERVATION_V2_MODES.preparing);
  properties.setProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY, 'FALSE');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    restoreReservationV2Mode_(previousMode);
    throw new Error('予約処理が実行中のためロールバックできません。少し待ってから再実行してください。');
  }
  try {
    const processing = selectSheetObjects_(SHEET_NAMES.reservations)
      .map(normalizeReservationRow_)
      .filter((row) => row.status === RESERVATION_STATUS.processing);
    if (processing.length > 0) {
      throw new Error(`処理中の予約が${processing.length}件あります。完了後に再実行してください。`);
    }
    properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
    properties.setProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY, '旧予約方式へ切り替えました。');
    restoreReservationV2Mode_(RESERVATION_V2_MODES.legacy);
    return getReservationV2ReleaseStatus();
  } catch (error) {
    restoreReservationV2Mode_(previousMode);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/** Apps Script エディタから確認できる安全な切替状態を返します。 */
function getReservationV2ReleaseStatus() {
  const properties = PropertiesService.getScriptProperties();
  const mode = getReservationV2Mode_();
  return {
    ok: true,
    mode,
    enabled: mode === RESERVATION_V2_MODES.v2,
    index_ready: isReservationV2IndexReady_(),
    index_built_at: properties.getProperty(RESERVATION_V2_INDEX_BUILT_AT_PROPERTY_KEY) || '',
    stale_reason: properties.getProperty(RESERVATION_V2_INDEX_STALE_REASON_PROPERTY_KEY) || '',
  };
}

/**
 * 設定と会議室を短時間キャッシュします。
 *
 * @return {Object} 設定、会議室一覧、会議室マップ。
 */
function getReservationStaticContextV2_() {
  const cache = CacheService.getScriptCache();
  const cachedText = cache.get(RESERVATION_V2_STATIC_CACHE_KEY);
  if (cachedText) {
    try {
      const cached = JSON.parse(cachedText);
      cached.roomMap = createRoomMap_(cached.rooms || []);
      return cached;
    } catch (_) {
    }
  }
  const context = {
    settings: getSettings_(),
    rooms: selectActiveRooms_(),
  };
  context.roomMap = createRoomMap_(context.rooms);
  try {
    cache.put(
      RESERVATION_V2_STATIC_CACHE_KEY,
      JSON.stringify({ settings: context.settings, rooms: context.rooms }),
      RESERVATION_V2_STATIC_CACHE_SECONDS
    );
  } catch (_) {
  }
  return context;
}

/** キャッシュ済みの設定・会議室情報を破棄します。 */
function invalidateReservationStaticCacheV2_() {
  try {
    CacheService.getScriptCache().remove(RESERVATION_V2_STATIC_CACHE_KEY);
  } catch (_) {
  }
}

/**
 * シートアクセスを伴わず入力値を検証します。
 *
 * @param {Object} payload 入力値。
 * @param {Object} context 設定・会議室情報。
 * @return {Object} 検証結果。
 */
function validateReservationRequestV2_(payload, context) {
  const common = normalizeCommonInput_(payload.common || {});
  const reservations = Array.isArray(payload.reservations)
    ? payload.reservations.map((reservation) => normalizeReservationInput_(reservation))
    : [];
  const settings = context.settings;
  const errors = [];
  const maxReservationCount = Number(settings.MAX_RESERVATIONS_PER_SUBMIT || DEFAULT_SETTINGS.MAX_RESERVATIONS_PER_SUBMIT);
  const fieldConfig = {
    organizationName: normalizeBoolean_(settings.FIELD_ORGANIZATION || DEFAULT_SETTINGS.FIELD_ORGANIZATION),
    userName: normalizeBoolean_(settings.FIELD_USER_NAME || DEFAULT_SETTINGS.FIELD_USER_NAME),
  };
  validateCommonInput_(common, fieldConfig).forEach((error) => errors.push(error));
  if (reservations.length === 0) {
    errors.push({ index: null, field: 'reservations', message: '予約を1件以上入力してください。' });
  }
  if (reservations.length > maxReservationCount) {
    errors.push({ index: null, field: 'reservations', message: `一度に登録できる予約は${maxReservationCount}件までです。` });
  }
  reservations.forEach((reservation, index) => {
    validateSingleReservationInput_(reservation, index, context.roomMap, settings).forEach((error) => errors.push(error));
  });
  reservations.forEach((reservation, index) => {
    reservations.forEach((otherReservation, otherIndex) => {
      if (index >= otherIndex) return;
      if (
        hasReservationRoomConflict_(reservation.room_id, otherReservation.room_id, context.roomMap)
        && reservation.usage_date === otherReservation.usage_date
        && hasTimeOverlap_(reservation.start_time, reservation.end_time, otherReservation.start_time, otherReservation.end_time)
      ) {
        const message = `${index + 1}件目と${otherIndex + 1}件目の予約時間が会議室構成上重複しています。`;
        errors.push({ index, field: 'room_id', message });
        errors.push({ index: otherIndex, field: 'room_id', message });
      }
    });
  });
  return { ok: errors.length === 0, errors, common, reservations };
}

/**
 * 有効または処理中の既存予約との競合を検出します。
 */
function findReservationConflictsV2_(reservations, existingReservations, roomMap) {
  const blocking = existingReservations.filter((reservation) =>
    reservation.status === RESERVATION_STATUS.active || reservation.status === RESERVATION_STATUS.processing
  );
  const errors = [];
  reservations.forEach((reservation, index) => {
    const conflict = blocking.find((existing) =>
      existing.usage_date === reservation.usage_date
      && hasReservationRoomConflict_(existing.room_id, reservation.room_id, roomMap)
      && hasTimeOverlap_(existing.start_time, existing.end_time, reservation.start_time, reservation.end_time)
    );
    if (conflict) {
      const suffix = conflict.status === RESERVATION_STATUS.processing ? '（現在処理中）' : '';
      errors.push({
        index,
        field: 'room_id',
        message: `${conflict.room_name}は${reservation.start_time}-${reservation.end_time}に既存予約があります${suffix}。`,
      });
    }
  });
  return errors;
}

/**
 * 処理中として一括保存する行とレスポンス用オブジェクトを作ります。
 */
function createPendingReservationRowsV2_(requestId, requestHash, common, reservations, roomMap, now) {
  const rows = [];
  const registeredReservations = [];
  reservations.forEach((reservation) => {
    const room = roomMap[reservation.room_id];
    const reservationId = `R-${formatDate_(new Date(), 'yyyyMMdd')}-${Utilities.getUuid().slice(0, 12)}`;
    rows.push([
      reservationId,
      common.line_user_id,
      '',
      common.organization_name,
      reservation.meeting_name,
      common.user_name,
      room.room_id,
      room.room_name,
      reservation.usage_date,
      reservation.start_time,
      reservation.end_time,
      '',
      RESERVATION_STATUS.processing,
      now,
      requestId,
      requestHash,
      now,
      now,
      '',
    ]);
    registeredReservations.push({
      reservation_id: reservationId,
      meeting_name: reservation.meeting_name,
      room_id: room.room_id,
      room_name: room.room_name,
      usage_date: reservation.usage_date,
      start_time: reservation.start_time,
      end_time: reservation.end_time,
      calendar_event_id: '',
      calendar_url: room.calendar_url,
    });
  });
  return { rows, reservations: registeredReservations };
}

/** 予約行内の列キーに対応する値を更新します。 */
function setReservationRowValueV2_(row, columnKey, value) {
  const index = SHEET_COLUMN_KEYS.reservations.indexOf(columnKey);
  if (index !== -1) row[index] = value;
}

/** 同一リクエストがすでに存在するときの結果を構築します。 */
function buildPreviousReservationResultV2_(requestId, requestHash, rows, metrics) {
  const hashMismatch = rows.some((row) => row.request_hash && row.request_hash !== requestHash);
  if (hashMismatch) {
    return {
      ok: false,
      state: 'invalid',
      code: 'REQUEST_ID_MISMATCH',
      request_id: requestId,
      error: '同じリクエストIDで異なる予約内容は送信できません。',
      metrics,
    };
  }
  if (rows.some((row) => row.status === RESERVATION_STATUS.failed)) {
    return {
      ok: false,
      state: 'failed',
      request_id: requestId,
      error: rows.map((row) => row.error_message).filter(Boolean).join('\n') || '予約登録に失敗しました。',
      reservations: rows.map(toReservationResponseV2_),
      metrics,
    };
  }
  const hasProcessing = rows.some((row) => row.status === RESERVATION_STATUS.processing);
  const allCancelled = rows.every((row) => row.status === RESERVATION_STATUS.cancelled);
  const state = hasProcessing ? 'processing' : (allCancelled ? 'cancelled' : 'active');
  return {
    ok: true,
    state,
    duplicate: true,
    request_id: requestId,
    reservations: rows.map(toReservationResponseV2_),
    metrics,
  };
}

/** 予約行を公開レスポンス用に整形します。 */
function toReservationResponseV2_(reservation) {
  return {
    reservation_id: reservation.reservation_id,
    meeting_name: reservation.meeting_name,
    room_id: reservation.room_id,
    room_name: reservation.room_name,
    usage_date: reservation.usage_date,
    start_time: reservation.start_time,
    end_time: reservation.end_time,
    calendar_event_id: reservation.calendar_event_id || '',
  };
}

/** 入力値を正規化して SHA-256 ハッシュを作ります。 */
function createReservationRequestHashV2_(common, reservations) {
  const text = JSON.stringify({ common, reservations });
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map((byte) => (`0${(byte + 256).toString(16)}`).slice(-2))
    .join('');
}

/** 連番走査を行わず、操作ログを1行追加します。 */
function writeOperationLogFastV2_(actionType, actor, targetId, detail, result, errorMessage) {
  try {
    getSheet_(SHEET_NAMES.logs).appendRow([
      `LOG-${Utilities.getUuid()}`,
      normalizeString_(actionType),
      normalizeString_(actor),
      normalizeString_(targetId),
      normalizeString_(detail),
      formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"),
      normalizeString_(result) || '成功',
      normalizeString_(errorMessage),
    ]);
  } catch (_) {
  }
}
