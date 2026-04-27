/**
 * GAS Web アプリの GET 入口です。
 *
 * @param {Object} e Apps Script から渡されるリクエストイベント。
 * @return {TextOutput} JSON レスポンス。
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = normalizeString_(params.action);

  try {
    if (action === 'get_config') {
      return jsonResponse_(handleGetConfig_());
    }
    if (action === 'get_public_settings') {
      return jsonResponse_(handleGetPublicSettings_());
    }
    return jsonResponse_({ ok: false, error: 'Unknown action' });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

/**
 * GAS Web アプリの POST 入口です。
 *
 * <p>予約登録、予約取消、出欠登録、会議保存、設定更新、集計更新を
 * action 名で振り分けます。</p>
 *
 * @param {Object} e Apps Script から渡されるリクエストイベント。
 * @return {TextOutput} JSON レスポンス。
 */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonResponse_({ ok: false, error: 'Invalid JSON' });
  }

  const action = normalizeString_(body.action);
  const payload = body.payload || {};

  try {
    if (action === 'insert_reservations') {
      return jsonResponse_(handleInsertReservations_(payload));
    }
    if (action === 'cancel_request') {
      return jsonResponse_(handleCancelRequest_(payload));
    }
    if (action === 'admin_delete') {
      return jsonResponse_(handleAdminDelete_(payload));
    }
    if (action === 'verifyAdminPassword') {
      return jsonResponse_(handleVerifyAdminPassword_(payload));
    }
    if (action === 'submitAttendance') {
      return jsonResponse_(handleSubmitAttendance_(payload));
    }
    if (action === 'saveMeeting') {
      return jsonResponse_(handleSaveMeeting_(payload));
    }
    if (action === 'updateSetting') {
      return jsonResponse_(handleUpdateSetting_(payload));
    }
    if (action === 'refreshAggregations') {
      return jsonResponse_(handleRefreshAggregations_(payload));
    }
    return jsonResponse_({ ok: false, error: 'Unknown action' });
  } catch (error) {
    writeOperationLog_('APIエラー', 'システム', action, error.message, '失敗', error.stack || '');
    return jsonResponse_({ ok: false, error: error.message });
  }
}

/**
 * 予約画面・予約確認画面で使う公開設定を返します。
 *
 * @return {Object} 画面設定、会議室一覧、今日の日付。
 */
function handleGetConfig_() {
  const settings = getSettings_();
  return {
    ok: true,
    today: formatDate_(new Date(), 'yyyy-MM-dd'),
    systemName: settings.SYSTEM_NAME || DEFAULT_SETTINGS.SYSTEM_NAME,
    settings: {
      systemName: settings.SYSTEM_NAME || DEFAULT_SETTINGS.SYSTEM_NAME,
      businessStartTime: settings.BUSINESS_START_TIME || DEFAULT_SETTINGS.BUSINESS_START_TIME,
      businessEndTime: settings.BUSINESS_END_TIME || DEFAULT_SETTINGS.BUSINESS_END_TIME,
      timeSlotMinutes: Number(settings.TIME_SLOT_MINUTES || DEFAULT_SETTINGS.TIME_SLOT_MINUTES),
      maxReservationsPerSubmit: Number(settings.MAX_RESERVATIONS_PER_SUBMIT || DEFAULT_SETTINGS.MAX_RESERVATIONS_PER_SUBMIT),
      labels: {
        organizationName: settings.LABEL_ORGANIZATION || DEFAULT_SETTINGS.LABEL_ORGANIZATION,
        userName: settings.LABEL_USER_NAME || DEFAULT_SETTINGS.LABEL_USER_NAME,
        meetingName: settings.LABEL_MEETING_NAME || DEFAULT_SETTINGS.LABEL_MEETING_NAME,
        room: settings.LABEL_ROOM || DEFAULT_SETTINGS.LABEL_ROOM,
      },
      fieldConfig: {
        organizationName: normalizeBoolean_(settings.FIELD_ORGANIZATION || DEFAULT_SETTINGS.FIELD_ORGANIZATION),
        userName: normalizeBoolean_(settings.FIELD_USER_NAME || DEFAULT_SETTINGS.FIELD_USER_NAME),
      },
      attendanceConfig: {
        absenceReasons: ABSENCE_REASONS.slice(),
        otherReasonDetailRequired: normalizeBoolean_(settings.OTHER_REASON_DETAIL_REQUIRED || DEFAULT_SETTINGS.OTHER_REASON_DETAIL_REQUIRED),
        absenceDetailAlwaysRequired: normalizeBoolean_(settings.ABSENCE_DETAIL_ALWAYS_REQUIRED || DEFAULT_SETTINGS.ABSENCE_DETAIL_ALWAYS_REQUIRED),
        alertStreakCount: Number(settings.ALERT_STREAK_COUNT || DEFAULT_SETTINGS.ALERT_STREAK_COUNT),
        criticalStreakCount: Number(settings.CRITICAL_STREAK_COUNT || DEFAULT_SETTINGS.CRITICAL_STREAK_COUNT),
      },
    },
    rooms: selectActiveRooms_(),
  };
}

/**
 * 出欠登録画面や管理画面で必要な公開設定だけを返します。
 *
 * @return {Object} システム名、欠席理由必須設定、連続欠席しきい値。
 */
function handleGetPublicSettings_() {
  const settings = getSettings_();
  return {
    ok: true,
    systemName: settings.SYSTEM_NAME || DEFAULT_SETTINGS.SYSTEM_NAME,
    otherReasonDetailRequired: normalizeBoolean_(settings.OTHER_REASON_DETAIL_REQUIRED || DEFAULT_SETTINGS.OTHER_REASON_DETAIL_REQUIRED),
    absenceDetailAlwaysRequired: normalizeBoolean_(settings.ABSENCE_DETAIL_ALWAYS_REQUIRED || DEFAULT_SETTINGS.ABSENCE_DETAIL_ALWAYS_REQUIRED),
    alertStreakCount: Number(settings.ALERT_STREAK_COUNT || DEFAULT_SETTINGS.ALERT_STREAK_COUNT),
    criticalStreakCount: Number(settings.CRITICAL_STREAK_COUNT || DEFAULT_SETTINGS.CRITICAL_STREAK_COUNT),
  };
}

/**
 * 予約登録 API の処理本体へ委譲します。
 *
 * @param {Object} payload 予約登録リクエスト。
 * @return {Object} 登録結果。
 */
function handleInsertReservations_(payload) {
  return insertReservations(payload);
}

/**
 * 利用者からのキャンセル依頼を予約シートに反映します。
 *
 * @param {Object} payload 予約IDと任意理由。
 * @return {Object} 更新結果。
 */
function handleCancelRequest_(payload) {
  const reservationId = normalizeString_(payload.reservation_id);
  const reason = normalizeString_(payload.reason);
  if (!reservationId) {
    return { ok: false, error: '予約IDが必要です。' };
  }

  const sheet = getSheet_(SHEET_NAMES.reservations);
  const rows = selectSheetObjects_(SHEET_NAMES.reservations);
  const targetIndex = rows.findIndex((row) => normalizeString_(row.reservation_id) === reservationId);
  if (targetIndex === -1) {
    return { ok: false, error: '予約が見つかりません。' };
  }
  const rowNumber = targetIndex + 2;
  const statusColumn = getColumnIndex_(SHEET_NAMES.reservations, 'status');
  const currentStatus = normalizeReservationStatus_(rows[targetIndex].status);
  if (currentStatus !== RESERVATION_STATUS.active) {
    return { ok: false, error: 'この予約はすでにキャンセル済みまたは依頼済みです。' };
  }
  sheet.getRange(rowNumber, statusColumn).setValue(RESERVATION_STATUS.cancelRequested);
  writeOperationLog_('キャンセル依頼', normalizeString_(rows[targetIndex].user_name) || '利用者', reservationId, reason || 'キャンセル依頼を登録しました。', '成功', '');
  return { ok: true, reservation_id: reservationId };
}

/**
 * 管理者操作として予約を取消状態に更新します。
 *
 * @param {Object} payload 管理者パスワードと予約ID配列。
 * @return {Object} 削除できたIDと見つからなかったID。
 */
function handleAdminDelete_(payload) {
  validateAdminPassword_(payload.adminPassword || payload.admin_password);
  const reservationIds = Array.isArray(payload.reservation_ids) ? payload.reservation_ids.map((id) => normalizeString_(id)).filter(Boolean) : [];
  if (reservationIds.length === 0) {
    return { ok: false, error: '削除する予約IDが必要です。' };
  }

  const sheet = getSheet_(SHEET_NAMES.reservations);
  const rows = selectSheetObjects_(SHEET_NAMES.reservations);
  const statusColumn = getColumnIndex_(SHEET_NAMES.reservations, 'status');
  const calendarEventIdColumn = getColumnIndex_(SHEET_NAMES.reservations, 'calendar_event_id');
  const roomIdColumn = getColumnIndex_(SHEET_NAMES.reservations, 'room_id');
  const deleted = [];
  const notFound = [];

  reservationIds.forEach((reservationId) => {
    const index = rows.findIndex((row) => normalizeString_(row.reservation_id) === reservationId);
    if (index === -1) {
      notFound.push(reservationId);
      return;
    }
    const row = rows[index];
    if (calendarEventIdColumn && roomIdColumn) {
      deleteCalendarEvent_(normalizeString_(row.room_id), normalizeString_(row.calendar_event_id));
    }
    sheet.getRange(index + 2, statusColumn).setValue(RESERVATION_STATUS.cancelled);
    deleted.push(reservationId);
    writeOperationLog_('予約削除', '管理者', reservationId, '予約を取消に更新しました。', '成功', '');
  });

  return { ok: true, deleted, notFound };
}

/**
 * 管理者パスワードだけを検証します。
 *
 * @param {Object} payload パスワード入力。
 * @return {Object} 検証成功結果。
 */
function handleVerifyAdminPassword_(payload) {
  validateAdminPassword_(payload.adminPassword || payload.admin_password);
  return { ok: true };
}

/**
 * メンバーの出欠回答を新規登録または更新します。
 *
 * <p>同じ会議IDとメンバーIDの組み合わせが既にある場合は、行追加せず
 * 既存行を上書きします。保存後に集計シートも再計算します。</p>
 *
 * @param {Object} payload 出欠登録リクエスト。
 * @return {Object} 保存した回答IDと更新有無。
 */
function handleSubmitAttendance_(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の出欠処理が実行中です。少し待ってから再度お試しください。');
  }

  try {
    const normalized = normalizeAttendancePayload_(payload);
    validateAttendancePayload_(normalized);

    const member = getActiveMemberById_(normalized.memberId);
    const meeting = getMeetingById_(normalized.meetingId);
    if (!member) {
      throw new Error('メンバーが見つかりません。');
    }
    if (!meeting || normalizeStatus_(meeting.status) !== '有効') {
      throw new Error('会議が見つからないか、無効です。');
    }
    if (!isWithinAnswerDeadline_(meeting.answer_deadline_at)) {
      throw new Error('回答締切日時を過ぎているため登録できません。');
    }

    const sheet = getSheet_(SHEET_NAMES.attendanceResponses);
    const rows = selectSheetObjects_(SHEET_NAMES.attendanceResponses);
    const rowIndex = rows.findIndex((row) => normalizeString_(row.meeting_id) === normalized.meetingId && normalizeString_(row.member_id) === normalized.memberId);
    const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    const record = [
      rowIndex === -1 ? createSequentialId_('ANS', now) : normalizeString_(rows[rowIndex].answer_id),
      normalized.meetingId,
      normalized.memberId,
      member.name,
      member.department,
      member.grade,
      member.role,
      normalized.attendance,
      normalized.attendance === ATTENDANCE_OPTIONS.absence ? normalized.absenceReason : '',
      normalized.attendance === ATTENDANCE_OPTIONS.absence ? normalized.absenceDetail : '',
      rowIndex === -1 ? now : normalizeString_(rows[rowIndex].answered_at) || now,
      now,
      rowIndex === -1 ? 1 : Number(rows[rowIndex].update_count || 0) + 1,
    ];

    if (rowIndex === -1) {
      appendSheetRows_(SHEET_NAMES.attendanceResponses, [record]);
      writeOperationLog_('出欠登録', member.name, normalized.meetingId, `${meeting.meeting_name} に ${normalized.attendance} で登録しました。`, '成功', '');
    } else {
      sheet.getRange(rowIndex + 2, 1, 1, record.length).setValues([record]);
      writeOperationLog_('出欠更新', member.name, normalized.meetingId, `${meeting.meeting_name} の回答を更新しました。`, '成功', '');
    }

    refreshAttendanceAggregations_();
    return {
      ok: true,
      answerId: record[0],
      updated: rowIndex !== -1,
      meetingId: normalized.meetingId,
      memberId: normalized.memberId,
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理者画面から会議を新規登録または更新します。
 *
 * @param {Object} payload 会議情報と管理者パスワード。
 * @return {Object} 保存した会議ID。
 */
function handleSaveMeeting_(payload) {
  validateAdminPassword_(payload.adminPassword || payload.admin_password);
  const normalized = normalizeMeetingPayload_(payload);
  validateMeetingPayload_(normalized);

  const sheet = getSheet_(SHEET_NAMES.meetings);
  const rows = selectSheetObjects_(SHEET_NAMES.meetings);
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  const meetingId = normalized.meetingId || createMeetingId_(normalized.meetingDate);
  const row = [
    meetingId,
    normalized.meetingName,
    normalized.meetingDate,
    normalized.startTime,
    normalized.endTime,
    normalized.recruitStartAt,
    normalized.recruitEndAt,
    normalized.answerDeadlineAt,
    normalized.note,
    normalized.status,
    now,
    now,
  ];

  const index = rows.findIndex((existing) => normalizeString_(existing.meeting_id) === meetingId);
  if (index === -1) {
    appendSheetRows_(SHEET_NAMES.meetings, [row]);
    writeOperationLog_('会議登録', '管理者', meetingId, `${normalized.meetingName} を登録しました。`, '成功', '');
  } else {
    row[10] = normalizeString_(rows[index].created_at) || now;
    sheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
    writeOperationLog_('会議更新', '管理者', meetingId, `${normalized.meetingName} を更新しました。`, '成功', '');
  }

  refreshAttendanceAggregations_();
  return { ok: true, meetingId };
}

/**
 * 管理者画面から設定シートの値を更新します。
 *
 * @param {Object} payload 設定キー、設定値、管理者パスワード。
 * @return {Object} 更新した設定キーと値。
 */
function handleUpdateSetting_(payload) {
  validateAdminPassword_(payload.adminPassword || payload.admin_password);
  const settingKey = getSettingInternalKey_(payload.settingKey || payload.setting_key);
  const settingValue = normalizeString_(payload.settingValue || payload.setting_value);
  if (!settingKey) {
    throw new Error('設定キーが不正です。');
  }

  const sheet = getSheet_(SHEET_NAMES.settings);
  const rows = selectSheetObjects_(SHEET_NAMES.settings);
  const index = rows.findIndex((row) => getSettingInternalKey_(row.setting_key) === settingKey);
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  const row = [
    SETTING_KEYS[settingKey] || settingKey,
    settingValue,
    DEFAULT_SETTING_DESCRIPTIONS[settingKey] || '',
    now,
  ];

  if (index === -1) {
    appendSheetRows_(SHEET_NAMES.settings, [row]);
  } else {
    sheet.getRange(index + 2, 1, 1, row.length).setValues([row]);
  }

  refreshAttendanceAggregations_();
  writeOperationLog_('設定更新', '管理者', settingKey, `${SETTING_KEYS[settingKey] || settingKey} を更新しました。`, '成功', '');
  return { ok: true, settingKey, settingValue };
}

/**
 * 出欠系の集計シートを手動で再計算します。
 *
 * @param {Object} payload 管理者パスワード。
 * @return {Object} 再計算件数。
 */
function handleRefreshAggregations_(payload) {
  validateAdminPassword_(payload.adminPassword || payload.admin_password);
  const result = refreshAttendanceAggregations_();
  writeOperationLog_('集計更新', '管理者', '', '出欠集計を再計算しました。', '成功', '');
  return { ok: true, refreshed: result };
}

/**
 * 複数の会議室予約を一括登録します。
 *
 * <p>Google カレンダー作成とスプレッドシート行追加をまとめて行い、
 * 途中失敗時は作成済みカレンダー予定を削除して巻き戻します。</p>
 *
 * @param {Object} payload 共通予約者情報と予約配列。
 * @return {Object} 登録された予約一覧。
 */
function insertReservations(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の予約処理が実行中です。少し待ってからもう一度お試しください。');
  }

  const createdEvents = [];
  try {
    const validation = validateReservationsInternal_(payload);
    if (!validation.ok) {
      return validation;
    }

    const common = validation.common;
    const roomMap = createRoomMap_(selectActiveRooms_());
    const createdAt = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    const nextReservationId = createSequentialIdGenerator_('R', createdAt, SHEET_NAMES.reservations, 'reservation_id');
    const reservationRows = [];
    const registeredReservations = [];

    validation.reservations.forEach((reservation) => {
      const room = roomMap[reservation.room_id];
      if (!room || !room.calendar_id) {
        throw new Error(`会議室「${room ? room.room_name : reservation.room_id}」のカレンダーIDが未設定です。`);
      }

      const reservationId = nextReservationId();
      const startDate = createDateTime_(reservation.usage_date, reservation.start_time);
      const endDate = createDateTime_(reservation.usage_date, reservation.end_time);
      const calendar = CalendarApp.getCalendarById(room.calendar_id);
      if (!calendar) {
        throw new Error(`会議室「${room.room_name}」の Google カレンダーが見つかりません。`);
      }

      const eventTitle = common.organization_name ? `【${common.organization_name}】${reservation.meeting_name}` : reservation.meeting_name;
      const eventDescription = [
        common.organization_name ? `団体名：${common.organization_name}` : null,
        `会議名：${reservation.meeting_name}`,
        common.user_name ? `お名前：${common.user_name}` : null,
        `会議室：${room.room_name}`,
        `予約ID：${reservationId}`,
      ].filter(Boolean).join('\n');
      const calendarEvent = calendar.createEvent(eventTitle, startDate, endDate, { description: eventDescription });
      createdEvents.push(calendarEvent);

      reservationRows.push([
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
        calendarEvent.getId(),
        RESERVATION_STATUS.active,
        createdAt,
      ]);

      registeredReservations.push({
        reservation_id: reservationId,
        meeting_name: reservation.meeting_name,
        room_id: room.room_id,
        room_name: room.room_name,
        usage_date: reservation.usage_date,
        start_time: reservation.start_time,
        end_time: reservation.end_time,
        calendar_url: room.calendar_url,
      });
    });

    appendSheetRows_(SHEET_NAMES.reservations, reservationRows);
    writeOperationLog_('予約登録', common.user_name || '利用者', '', `${registeredReservations.length}件の予約を登録しました。`, '成功', '');
    return { ok: true, errors: [], reservations: registeredReservations };
  } catch (error) {
    createdEvents.forEach((calendarEvent) => {
      try {
        calendarEvent.deleteEvent();
      } catch (_) {
      }
    });
    writeOperationLog_('予約登録', '利用者', '', '予約登録に失敗しました。', '失敗', error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 予約登録前の入力値と重複予約を検証します。
 *
 * @param {Object} payload フロントから送信された予約情報。
 * @return {Object} 検証結果、正規化済み入力、エラー一覧。
 */
function validateReservationsInternal_(payload) {
  const requestPayload = payload || {};
  const common = normalizeCommonInput_(requestPayload.common || {});
  const reservations = Array.isArray(requestPayload.reservations)
    ? requestPayload.reservations.map((reservation) => normalizeReservationInput_(reservation))
    : [];
  const errors = [];
  const settings = getSettings_();
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

  const roomMap = createRoomMap_(selectActiveRooms_());
  const existingReservations = selectActiveReservations_();
  reservations.forEach((reservation, index) => {
    validateSingleReservationInput_(reservation, index, roomMap, settings).forEach((error) => errors.push(error));
  });

  reservations.forEach((reservation, index) => {
    if (!reservation.usage_date || !reservation.start_time || !reservation.end_time || !reservation.room_id) {
      return;
    }
    const duplicateExisting = existingReservations.find((existing) => existing.room_id === reservation.room_id && existing.usage_date === reservation.usage_date && existing.status === RESERVATION_STATUS.active && hasTimeOverlap_(existing.start_time, existing.end_time, reservation.start_time, reservation.end_time));
    if (duplicateExisting) {
      errors.push({ index, field: 'room_id', message: `${duplicateExisting.room_name}は${reservation.start_time}-${reservation.end_time}に既存予約があります。` });
    }
    reservations.forEach((otherReservation, otherIndex) => {
      if (index >= otherIndex) {
        return;
      }
      if (reservation.room_id === otherReservation.room_id && reservation.usage_date === otherReservation.usage_date && hasTimeOverlap_(reservation.start_time, reservation.end_time, otherReservation.start_time, otherReservation.end_time)) {
        const message = `${index + 1}件目と${otherIndex + 1}件目の予約時間が同じ会議室で重複しています。`;
        errors.push({ index, field: 'room_id', message });
        errors.push({ index: otherIndex, field: 'room_id', message });
      }
    });
  });

  if (errors.length > 0) {
    return { ok: false, errors, common, reservations };
  }
  return { ok: true, errors: [], common, reservations };
}

/**
 * 予約者共通情報の必須チェックを行います。
 *
 * @param {Object} common 正規化済み共通情報。
 * @param {Object} fieldConfig 表示中フィールドの設定。
 * @return {Array<Object>} エラー一覧。
 */
function validateCommonInput_(common, fieldConfig) {
  const errors = [];
  if (fieldConfig.organizationName && !common.organization_name) {
    errors.push({ index: null, field: 'organization_name', message: '団体名を入力してください。' });
  }
  if (fieldConfig.userName && !common.user_name) {
    errors.push({ index: null, field: 'user_name', message: 'お名前を入力してください。' });
  }
  return errors;
}

/**
 * 予約1件分の必須項目、時間、会議室利用可否を検証します。
 *
 * @param {Object} reservation 正規化済み予約情報。
 * @param {number} index 予約配列内の位置。
 * @param {Object} roomMap 会議室IDをキーにした会議室情報。
 * @param {Object} settings 設定シート由来の設定値。
 * @return {Array<Object>} エラー一覧。
 */
function validateSingleReservationInput_(reservation, index, roomMap, settings) {
  const errors = [];
  if (!reservation.meeting_name) {
    errors.push({ index, field: 'meeting_name', message: '会議名を入力してください。' });
  }
  validateDateTimeInput_(reservation.usage_date, reservation.start_time, reservation.end_time, settings).forEach((message) => {
    errors.push({ index, field: 'date_time', message });
  });
  if (!reservation.room_id) {
    errors.push({ index, field: 'room_id', message: '会議室を選択してください。' });
  } else if (!roomMap[reservation.room_id]) {
    errors.push({ index, field: 'room_id', message: '選択された会議室は利用できません。' });
  } else if (!roomMap[reservation.room_id].calendar_id) {
    errors.push({ index, field: 'room_id', message: `会議室「${roomMap[reservation.room_id].room_name}」のカレンダーIDが未設定です。` });
  }
  return errors;
}

/**
 * 利用日・開始時刻・終了時刻の形式と営業時間を検証します。
 *
 * @param {string} usageDate yyyy-MM-dd 形式の利用日。
 * @param {string} startTime HH:mm 形式の開始時刻。
 * @param {string} endTime HH:mm 形式の終了時刻。
 * @param {Object} settings 営業時間や時間刻み設定。
 * @return {Array<string>} エラーメッセージ一覧。
 */
function validateDateTimeInput_(usageDate, startTime, endTime, settings) {
  const errors = [];
  if (!usageDate || !/^\d{4}-\d{2}-\d{2}$/.test(usageDate)) {
    errors.push('利用日の形式を確認してください。');
  }
  if (!startTime || !/^\d{2}:\d{2}$/.test(startTime)) {
    errors.push('開始時間の形式を確認してください。');
  }
  if (!endTime || !/^\d{2}:\d{2}$/.test(endTime)) {
    errors.push('終了時間の形式を確認してください。');
  }
  if (startTime && endTime && toMinutes_(startTime) >= toMinutes_(endTime)) {
    errors.push('終了時間は開始時間より後にしてください。');
  }
  if (settings && startTime && endTime) {
    const slotMinutes = Number(settings.TIME_SLOT_MINUTES || DEFAULT_SETTINGS.TIME_SLOT_MINUTES || 5);
    const businessStartMin = toMinutes_(settings.BUSINESS_START_TIME || DEFAULT_SETTINGS.BUSINESS_START_TIME || '09:00');
    const businessEndMin = toMinutes_(settings.BUSINESS_END_TIME || DEFAULT_SETTINGS.BUSINESS_END_TIME || '22:00');
    const startMin = toMinutes_(startTime);
    const endMin = toMinutes_(endTime);
    if (startMin < businessStartMin || startMin >= businessEndMin) {
      errors.push(`開始時間は営業時間内（${settings.BUSINESS_START_TIME || DEFAULT_SETTINGS.BUSINESS_START_TIME}〜${settings.BUSINESS_END_TIME || DEFAULT_SETTINGS.BUSINESS_END_TIME}）で入力してください。`);
    }
    if (endMin <= businessStartMin || endMin > businessEndMin) {
      errors.push(`終了時間は営業時間内（${settings.BUSINESS_START_TIME || DEFAULT_SETTINGS.BUSINESS_START_TIME}〜${settings.BUSINESS_END_TIME || DEFAULT_SETTINGS.BUSINESS_END_TIME}）で入力してください。`);
    }
    if (slotMinutes > 1 && startMin % slotMinutes !== 0) {
      errors.push(`開始時間は${slotMinutes}分刻みで入力してください。`);
    }
    if (slotMinutes > 1 && endMin % slotMinutes !== 0) {
      errors.push(`終了時間は${slotMinutes}分刻みで入力してください。`);
    }
  }
  return errors;
}

/**
 * 予約者共通情報を保存用キーへ正規化します。
 *
 * @param {Object} commonInput フロント入力。
 * @return {Object} 正規化済み共通情報。
 */
function normalizeCommonInput_(commonInput) {
  return {
    line_user_id: normalizeString_(commonInput.line_user_id || commonInput.lineUserId),
    organization_name: normalizeString_(commonInput.organization_name || commonInput.organizationName),
    user_name: normalizeString_(commonInput.user_name || commonInput.userName),
  };
}

/**
 * 予約1件分の入力を保存用キーへ正規化します。
 *
 * @param {Object} reservationInput フロント入力。
 * @return {Object} 正規化済み予約情報。
 */
function normalizeReservationInput_(reservationInput) {
  const input = reservationInput || {};
  return {
    meeting_name: normalizeString_(input.meeting_name || input.meetingName),
    usage_date: normalizeDateString_(input.usage_date || input.usageDate),
    start_time: normalizeTimeString_(input.start_time || input.startTime),
    end_time: normalizeTimeString_(input.end_time || input.endTime),
    room_id: normalizeString_(input.room_id || input.roomId),
  };
}

/**
 * 出欠登録リクエストを内部キーへ正規化します。
 *
 * @param {Object} payload フロント入力。
 * @return {Object} 正規化済み出欠情報。
 */
function normalizeAttendancePayload_(payload) {
  const input = payload || {};
  return {
    meetingId: normalizeString_(input.meetingId || input.meeting_id),
    memberId: normalizeString_(input.memberId || input.member_id),
    attendance: normalizeString_(input.attendance),
    absenceReason: normalizeString_(input.absenceReason || input.absence_reason),
    absenceDetail: normalizeString_(input.absenceDetail || input.absence_detail),
  };
}

/**
 * 出欠登録の入力値を検証します。
 *
 * @param {Object} payload 正規化済み出欠情報。
 */
function validateAttendancePayload_(payload) {
  const settings = getSettings_();
  const otherReasonDetailRequired = normalizeBoolean_(settings.OTHER_REASON_DETAIL_REQUIRED || DEFAULT_SETTINGS.OTHER_REASON_DETAIL_REQUIRED);
  const absenceDetailAlwaysRequired = normalizeBoolean_(settings.ABSENCE_DETAIL_ALWAYS_REQUIRED || DEFAULT_SETTINGS.ABSENCE_DETAIL_ALWAYS_REQUIRED);
  if (!payload.meetingId) {
    throw new Error('会議が選択されていません。');
  }
  if (!payload.memberId) {
    throw new Error('名前を選択してください。');
  }
  if (payload.attendance !== ATTENDANCE_OPTIONS.attend && payload.attendance !== ATTENDANCE_OPTIONS.absence) {
    throw new Error('出欠は出席または欠席を選択してください。');
  }
  if (payload.attendance === ATTENDANCE_OPTIONS.absence) {
    if (!payload.absenceReason) {
      throw new Error('欠席理由を選択してください。');
    }
    if (ABSENCE_REASONS.indexOf(payload.absenceReason) === -1) {
      throw new Error('欠席理由が不正です。');
    }
    if (payload.absenceReason === 'その他' && otherReasonDetailRequired && !payload.absenceDetail) {
      throw new Error('その他を選んだ場合は詳細理由を入力してください。');
    }
    if (absenceDetailAlwaysRequired && !payload.absenceDetail) {
      throw new Error('欠席理由の詳細を入力してください。');
    }
  }
}

/**
 * 会議登録リクエストを内部キーへ正規化します。
 *
 * @param {Object} payload フロント入力。
 * @return {Object} 正規化済み会議情報。
 */
function normalizeMeetingPayload_(payload) {
  const input = payload || {};
  return {
    meetingId: normalizeString_(input.meetingId || input.meeting_id),
    meetingName: normalizeString_(input.meetingName || input.meeting_name),
    meetingDate: normalizeDateString_(input.meetingDate || input.meeting_date),
    startTime: normalizeTimeString_(input.startTime || input.start_time),
    endTime: normalizeTimeString_(input.endTime || input.end_time),
    recruitStartAt: normalizeDateTimeInput_(input.recruitStartAt || input.recruit_start_at),
    recruitEndAt: normalizeDateTimeInput_(input.recruitEndAt || input.recruit_end_at),
    answerDeadlineAt: normalizeDateTimeInput_(input.answerDeadlineAt || input.answer_deadline_at),
    note: normalizeString_(input.note),
    status: normalizeStatus_(input.status) || '有効',
  };
}

/**
 * 会議登録・更新の入力値を検証します。
 *
 * @param {Object} meeting 正規化済み会議情報。
 */
function validateMeetingPayload_(meeting) {
  if (!meeting.meetingName) {
    throw new Error('会議名を入力してください。');
  }
  if (!meeting.meetingDate) {
    throw new Error('会議日を入力してください。');
  }
  if (!meeting.startTime || !meeting.endTime) {
    throw new Error('開始時間と終了時間を入力してください。');
  }
  if (toMinutes_(meeting.startTime) >= toMinutes_(meeting.endTime)) {
    throw new Error('終了時間は開始時間より後にしてください。');
  }
  if (!meeting.recruitStartAt || !meeting.recruitEndAt || !meeting.answerDeadlineAt) {
    throw new Error('募集期間と回答締切日時を入力してください。');
  }
  if (meeting.recruitStartAt > meeting.recruitEndAt) {
    throw new Error('募集開始日時は募集終了日時以前にしてください。');
  }
  if (meeting.answerDeadlineAt > meeting.recruitEndAt) {
    throw new Error('回答締切日時は募集終了日時以前にしてください。');
  }
  if (meeting.status !== '有効' && meeting.status !== '無効') {
    throw new Error('有効状態は有効または無効を指定してください。');
  }
}

/**
 * 出欠回答から会議別、個人別、部局別、連続欠席の集計を再計算します。
 *
 * <p>未提出は出席扱いとして計算します。連続欠席は会議日が新しい順に見て、
 * 出席回答または未提出が出た時点で止めます。</p>
 *
 * @return {Object} 更新した集計行数。
 */
function refreshAttendanceAggregations_() {
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  const today = formatDate_(new Date(), 'yyyy-MM-dd');
  const settings = getSettings_();
  const members = selectActiveMembers_();
  const meetings = selectSheetObjects_(SHEET_NAMES.meetings)
    .map(normalizeMeetingRow_)
    .filter((meeting) => meeting.meeting_id && normalizeStatus_(meeting.status) === '有効' && meeting.meeting_date && meeting.meeting_date <= today)
    .sort((left, right) => `${left.meeting_date}${left.start_time}${left.meeting_id}`.localeCompare(`${right.meeting_date}${right.start_time}${right.meeting_id}`));
  const responses = selectSheetObjects_(SHEET_NAMES.attendanceResponses)
    .map(normalizeAttendanceRow_)
    .filter((row) => row.meeting_id && row.member_id);
  const responseMap = {};
  responses.forEach((response) => {
    responseMap[`${response.meeting_id}__${response.member_id}`] = response;
  });

  const meetingRows = [];
  const memberRows = [];
  const departmentStats = {};
  const streakRows = [];
  const alertStreakCount = Number(settings.ALERT_STREAK_COUNT || DEFAULT_SETTINGS.ALERT_STREAK_COUNT);
  const criticalStreakCount = Number(settings.CRITICAL_STREAK_COUNT || DEFAULT_SETTINGS.CRITICAL_STREAK_COUNT);
  const meetingStats = {};
  meetings.forEach((meeting) => {
    meetingStats[meeting.meeting_id] = { attendCount: 0, absenceCount: 0, unansweredCount: 0 };
  });

  members.forEach((member) => {
    let attendCount = 0;
    let absenceCount = 0;
    let unansweredCount = 0;
    let lastAnsweredAt = '';
    let streak = 0;
    let latestAbsenceMeetingName = '';
    let latestAbsenceReason = '';

    meetings.forEach((meeting) => {
      const response = responseMap[`${meeting.meeting_id}__${member.member_id}`];
      const meetingStat = meetingStats[meeting.meeting_id];
      if (!response) {
        unansweredCount += 1;
        meetingStat.unansweredCount += 1;
        streak = 0;
        return;
      }
      if (response.attendance === ATTENDANCE_OPTIONS.absence) {
        absenceCount += 1;
        meetingStat.absenceCount += 1;
        streak += 1;
        latestAbsenceMeetingName = meeting.meeting_name;
        latestAbsenceReason = response.absence_reason;
      } else {
        attendCount += 1;
        meetingStat.attendCount += 1;
        streak = 0;
      }
      if (response.updated_at && (!lastAnsweredAt || response.updated_at > lastAnsweredAt)) {
        lastAnsweredAt = response.updated_at;
      }
    });

    const effectiveAttendCount = attendCount + unansweredCount;
    memberRows.push([
      member.member_id,
      member.name,
      member.department,
      member.grade,
      member.role,
      meetings.length,
      attendCount,
      absenceCount,
      unansweredCount,
      effectiveAttendCount,
      meetings.length ? effectiveAttendCount / meetings.length : 0,
      streak,
      lastAnsweredAt,
      now,
    ]);

    const departmentName = member.department || '未設定';
    if (!departmentStats[departmentName]) {
      departmentStats[departmentName] = { department: departmentName, memberCount: 0, attendCount: 0, absenceCount: 0, unansweredCount: 0 };
    }
    departmentStats[departmentName].memberCount += 1;
    departmentStats[departmentName].attendCount += attendCount;
    departmentStats[departmentName].absenceCount += absenceCount;
    departmentStats[departmentName].unansweredCount += unansweredCount;

    streakRows.push([
      member.member_id,
      member.name,
      member.department,
      member.grade,
      member.role,
      streak,
      determineAlertLevel_(streak, alertStreakCount, criticalStreakCount),
      streak ? latestAbsenceMeetingName : '',
      streak ? latestAbsenceReason : '',
      now,
    ]);
  });

  meetings.forEach((meeting) => {
    const stat = meetingStats[meeting.meeting_id];
    const targetCount = members.length;
    const effectiveAttendCount = stat.attendCount + stat.unansweredCount;
    meetingRows.push([
      meeting.meeting_id,
      meeting.meeting_name,
      meeting.meeting_date,
      targetCount,
      stat.attendCount,
      stat.absenceCount,
      stat.unansweredCount,
      effectiveAttendCount,
      targetCount ? effectiveAttendCount / targetCount : 0,
      targetCount ? stat.absenceCount / targetCount : 0,
      now,
    ]);
  });

  const departmentRows = Object.keys(departmentStats).sort().map((departmentName) => {
    const stat = departmentStats[departmentName];
    const targetMeetingCount = meetings.length;
    const targetSlotCount = stat.memberCount * targetMeetingCount;
    const effectiveAttendCount = stat.attendCount + stat.unansweredCount;
    return [
      stat.department,
      stat.memberCount,
      targetMeetingCount,
      stat.attendCount,
      stat.absenceCount,
      stat.unansweredCount,
      effectiveAttendCount,
      targetSlotCount ? effectiveAttendCount / targetSlotCount : 0,
      targetSlotCount ? stat.absenceCount / targetSlotCount : 0,
      now,
    ];
  });

  replaceSheetData_(SHEET_NAMES.meetingAggregations, meetingRows);
  replaceSheetData_(SHEET_NAMES.memberAggregations, memberRows);
  replaceSheetData_(SHEET_NAMES.departmentAggregations, departmentRows);
  replaceSheetData_(SHEET_NAMES.streaks, streakRows);

  return {
    meetings: meetingRows.length,
    members: memberRows.length,
    departments: departmentRows.length,
    streaks: streakRows.length,
  };
}

/**
 * 連続欠席回数から注意レベルを決定します。
 *
 * @param {number} streak 連続欠席回数。
 * @param {number} alertStreakCount 注意扱いにする回数。
 * @param {number} criticalStreakCount 要確認扱いにする回数。
 * @return {string} 通常、注意、要確認のいずれか。
 */
function determineAlertLevel_(streak, alertStreakCount, criticalStreakCount) {
  if (streak >= criticalStreakCount) {
    return '要確認';
  }
  if (streak >= alertStreakCount) {
    return '注意';
  }
  return '通常';
}

/**
 * 集計シートの既存データを消して、新しい行に置き換えます。
 *
 * @param {string} sheetName 対象シート名。
 * @param {Array<Array<*>>} rows 書き込む行配列。
 */
function replaceSheetData_(sheetName, rows) {
  const sheet = getSheet_(sheetName);
  const headers = SHEET_HEADERS[getSchemaKeyBySheetName_(sheetName)];
  const maxColumns = headers.length;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, maxColumns).clearContent();
  }
  if (rows.length === 0) {
    return;
  }
  sheet.getRange(2, 1, rows.length, maxColumns).setValues(rows);
}

/**
 * 設定シートの管理者パスワードと入力値を照合します。
 *
 * @param {*} inputPassword 入力されたパスワード。
 */
function validateAdminPassword_(inputPassword) {
  const settings = getSettings_();
  const correctPassword = normalizeString_(settings.ADMIN_PASSWORD || DEFAULT_SETTINGS.ADMIN_PASSWORD);
  if (!correctPassword || normalizeString_(inputPassword) !== correctPassword) {
    throw new Error('パスワードが違います。');
  }
}

/**
 * 予約取消時に、対応する Google カレンダー予定も削除します。
 *
 * @param {string} roomId 会議室ID。
 * @param {string} calendarEventId カレンダー予定ID。
 */
function deleteCalendarEvent_(roomId, calendarEventId) {
  try {
    const room = createRoomMap_(selectActiveRooms_())[roomId];
    if (!room || !room.calendar_id || !calendarEventId) {
      return;
    }
    const calendar = CalendarApp.getCalendarById(room.calendar_id);
    if (!calendar) {
      return;
    }
    const event = calendar.getEventById(calendarEventId);
    if (event) {
      event.deleteEvent();
    }
  } catch (_) {
  }
}

/**
 * 有効状態の予約だけを取得します。
 *
 * @return {Array<Object>} 正規化済み予約一覧。
 */
function selectActiveReservations_() {
  return selectSheetObjects_(SHEET_NAMES.reservations)
    .map(normalizeReservationRow_)
    .filter((reservation) => reservation.reservation_id && reservation.status === RESERVATION_STATUS.active);
}

/**
 * 予約シートの行を内部処理用オブジェクトに正規化します。
 *
 * @param {Object} row シート行。
 * @return {Object} 正規化済み予約。
 */
function normalizeReservationRow_(row) {
  return {
    reservation_id: normalizeString_(row.reservation_id),
    line_user_id: normalizeString_(row.line_user_id),
    email: normalizeString_(row.email),
    organization_name: normalizeString_(row.organization_name),
    meeting_name: normalizeString_(row.meeting_name),
    user_name: normalizeString_(row.user_name),
    room_id: normalizeString_(row.room_id),
    room_name: normalizeString_(row.room_name),
    usage_date: normalizeDateString_(row.usage_date),
    start_time: normalizeTimeString_(row.start_time),
    end_time: normalizeTimeString_(row.end_time),
    calendar_event_id: normalizeString_(row.calendar_event_id),
    status: normalizeReservationStatus_(row.status),
    created_at: normalizeDateTimeString_(row.created_at),
  };
}

/**
 * 会議一覧シートの行を内部処理用オブジェクトに正規化します。
 *
 * @param {Object} row シート行。
 * @return {Object} 正規化済み会議。
 */
function normalizeMeetingRow_(row) {
  return {
    meeting_id: normalizeString_(row.meeting_id),
    meeting_name: normalizeString_(row.meeting_name),
    meeting_date: normalizeDateString_(row.meeting_date),
    start_time: normalizeTimeString_(row.start_time),
    end_time: normalizeTimeString_(row.end_time),
    recruit_start_at: normalizeDateTimeInput_(row.recruit_start_at),
    recruit_end_at: normalizeDateTimeInput_(row.recruit_end_at),
    answer_deadline_at: normalizeDateTimeInput_(row.answer_deadline_at),
    note: normalizeString_(row.note),
    status: normalizeStatus_(row.status),
    created_at: normalizeDateTimeString_(row.created_at),
    updated_at: normalizeDateTimeString_(row.updated_at),
  };
}

/**
 * 出欠回答シートの行を内部処理用オブジェクトに正規化します。
 *
 * @param {Object} row シート行。
 * @return {Object} 正規化済み出欠回答。
 */
function normalizeAttendanceRow_(row) {
  return {
    answer_id: normalizeString_(row.answer_id),
    meeting_id: normalizeString_(row.meeting_id),
    member_id: normalizeString_(row.member_id),
    name: normalizeString_(row.name),
    department: normalizeString_(row.department),
    grade: normalizeString_(row.grade),
    role: normalizeString_(row.role),
    attendance: normalizeString_(row.attendance),
    absence_reason: normalizeString_(row.absence_reason),
    absence_detail: normalizeString_(row.absence_detail),
    answered_at: normalizeDateTimeString_(row.answered_at),
    updated_at: normalizeDateTimeString_(row.updated_at),
    update_count: Number(row.update_count || 0),
  };
}

/**
 * 利用可能な会議室を表示順で取得します。
 *
 * @return {Array<Object>} 会議室一覧。
 */
function selectActiveRooms_() {
  return selectSheetObjects_(SHEET_NAMES.rooms)
    .map((row) => {
      const calendarId = normalizeString_(row.calendar_id);
      return {
        room_id: normalizeString_(row.room_id),
        room_name: normalizeString_(row.room_name),
        calendar_id: calendarId,
        calendar_url: createCalendarUrl_(calendarId),
        is_active: normalizeBoolean_(row.is_active),
        display_order: Number(row.display_order || 0),
      };
    })
    .filter((room) => room.room_id && room.room_name && room.is_active)
    .sort((left, right) => Number(left.display_order) - Number(right.display_order) || compareValues_(left.room_name, right.room_name));
}

/**
 * 会議室配列を会議室IDキーのマップに変換します。
 *
 * @param {Array<Object>} rooms 会議室一覧。
 * @return {Object} 会議室IDをキーにしたマップ。
 */
function createRoomMap_(rooms) {
  const roomMap = {};
  rooms.forEach((room) => {
    roomMap[room.room_id] = room;
  });
  return roomMap;
}

/**
 * 有効なメンバーを表示順で取得します。
 *
 * @return {Array<Object>} メンバー一覧。
 */
function selectActiveMembers_() {
  return selectSheetObjects_(SHEET_NAMES.members)
    .map((row) => ({
      member_id: normalizeString_(row.member_id),
      name: normalizeString_(row.name),
      department: normalizeString_(row.department),
      grade: normalizeString_(row.grade),
      role: normalizeString_(row.role),
      display_order: Number(row.display_order || 0),
      status: normalizeStatus_(row.status),
    }))
    .filter((member) => member.member_id && member.name && member.status === '有効')
    .sort((left, right) => Number(left.display_order) - Number(right.display_order) || compareValues_(left.name, right.name));
}

/**
 * メンバーIDから有効メンバーを1件取得します。
 *
 * @param {string} memberId メンバーID。
 * @return {Object|null} メンバー情報。見つからない場合は null。
 */
function getActiveMemberById_(memberId) {
  return selectActiveMembers_().find((member) => member.member_id === normalizeString_(memberId)) || null;
}

/**
 * 会議IDから会議情報を1件取得します。
 *
 * @param {string} meetingId 会議ID。
 * @return {Object|null} 会議情報。見つからない場合は null。
 */
function getMeetingById_(meetingId) {
  return selectSheetObjects_(SHEET_NAMES.meetings).map(normalizeMeetingRow_).find((meeting) => meeting.meeting_id === normalizeString_(meetingId)) || null;
}

/**
 * 現在時刻が回答締切日時以前かを判定します。
 *
 * @param {string} answerDeadlineAt 回答締切日時。
 * @return {boolean} 回答可能な場合は true。
 */
function isWithinAnswerDeadline_(answerDeadlineAt) {
  const deadline = parseDateTime_(answerDeadlineAt);
  if (!deadline) {
    return false;
  }
  return new Date().getTime() <= deadline.getTime();
}

/**
 * 会議日をもとに会議IDを採番します。
 *
 * @param {string} meetingDate yyyy-MM-dd 形式の会議日。
 * @return {string} MTG-yyyymmdd-連番形式の会議ID。
 */
function createMeetingId_(meetingDate) {
  return createSequentialIdFromSheet_('MTG', `${meetingDate}T00:00:00`, SHEET_NAMES.meetings, 'meeting_id');
}

/**
 * 予約IDを採番します。
 *
 * @return {string} R-yyyymmdd-連番形式の予約ID。
 */
function createReservationId_() {
  return createSequentialIdFromSheet_('R', formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"), SHEET_NAMES.reservations, 'reservation_id');
}

/**
 * 接頭辞と日付から用途別の連番IDを作成します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateTimeText 採番日付を含む文字列。
 * @return {string} 採番済みID。
 */
function createSequentialId_(prefix, dateTimeText) {
  const target = getSequentialIdTarget_(prefix);
  if (target) {
    return createSequentialIdFromSheet_(prefix, dateTimeText, target.sheetName, target.columnKey);
  }
  return createSequentialIdFromAllSheets_(prefix, dateTimeText);
}

/**
 * ID接頭辞に対応する採番対象シートと列を返します。
 *
 * @param {string} prefix ID接頭辞。
 * @return {Object|null} 採番対象情報。
 */
function getSequentialIdTarget_(prefix) {
  const targets = {
    R: { sheetName: SHEET_NAMES.reservations, columnKey: 'reservation_id' },
    MTG: { sheetName: SHEET_NAMES.meetings, columnKey: 'meeting_id' },
    ANS: { sheetName: SHEET_NAMES.attendanceResponses, columnKey: 'answer_id' },
    LOG: { sheetName: SHEET_NAMES.logs, columnKey: 'log_id' },
  };
  return targets[normalizeString_(prefix)] || null;
}

/**
 * 指定シートの既存IDを見て次の連番IDを作成します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateTimeText 採番日付を含む文字列。
 * @param {string} sheetName 対象シート名。
 * @param {string} columnKey ID列の内部キー。
 * @return {string} 採番済みID。
 */
function createSequentialIdFromSheet_(prefix, dateTimeText, sheetName, columnKey) {
  const dateText = createSequentialDateText_(dateTimeText);
  const nextNumber = getMaxSequentialNumberFromSheet_(prefix, dateText, sheetName, columnKey) + 1;
  return formatSequentialId_(prefix, dateText, nextNumber);
}

/**
 * 連続採番に使うクロージャを生成します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateTimeText 採番日付を含む文字列。
 * @param {string} sheetName 対象シート名。
 * @param {string} columnKey ID列の内部キー。
 * @return {Function} 呼び出すたびに次のIDを返す関数。
 */
function createSequentialIdGenerator_(prefix, dateTimeText, sheetName, columnKey) {
  const dateText = createSequentialDateText_(dateTimeText);
  let nextNumber = getMaxSequentialNumberFromSheet_(prefix, dateText, sheetName, columnKey) + 1;
  return () => {
    const id = formatSequentialId_(prefix, dateText, nextNumber);
    nextNumber += 1;
    return id;
  };
}

/**
 * 接頭辞の対象シートが未定義の場合に、全シートから最大連番を探して採番します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateTimeText 採番日付を含む文字列。
 * @return {string} 採番済みID。
 */
function createSequentialIdFromAllSheets_(prefix, dateTimeText) {
  const dateText = createSequentialDateText_(dateTimeText);
  const existingIds = [];
  Object.keys(SHEET_NAMES).forEach((schemaKey) => {
    const sheetName = SHEET_NAMES[schemaKey];
    const firstColumnKey = (SHEET_COLUMN_KEYS[schemaKey] || [])[0];
    if (!firstColumnKey) {
      return;
    }
    try {
      selectSheetObjects_(sheetName).forEach((row) => {
        const id = normalizeString_(row[firstColumnKey]);
        if (id) {
          existingIds.push(id);
        }
      });
    } catch (_) {
    }
  });
  return formatSequentialId_(prefix, dateText, getMaxSequentialNumber_(prefix, dateText, existingIds) + 1);
}

/**
 * 指定シートのID列から指定日付の最大連番を取得します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateText yyyymmdd 形式の日付。
 * @param {string} sheetName 対象シート名。
 * @param {string} columnKey ID列の内部キー。
 * @return {number} 最大連番。
 */
function getMaxSequentialNumberFromSheet_(prefix, dateText, sheetName, columnKey) {
  const columnIndex = getColumnIndex_(sheetName, columnKey);
  if (!columnIndex) {
    return 0;
  }
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }
  const values = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues().map((row) => row[0]);
  return getMaxSequentialNumber_(prefix, dateText, values);
}

/**
 * ID文字列配列から指定接頭辞・日付の最大連番を取得します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateText yyyymmdd 形式の日付。
 * @param {Array<string>} values ID文字列配列。
 * @return {number} 最大連番。
 */
function getMaxSequentialNumber_(prefix, dateText, values) {
  const escapedPrefix = normalizeString_(prefix).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`^${escapedPrefix}-${dateText}-(\\d{3,})$`);
  let max = 0;
  values.forEach((value) => {
    const match = normalizeString_(value).match(pattern);
    if (match) {
      max = Math.max(max, Number(match[1]));
    }
  });
  return max;
}

/**
 * 採番に使う yyyymmdd 形式の日付文字列を作成します。
 *
 * @param {string} dateTimeText 日付を含む文字列。
 * @return {string} yyyymmdd 形式の日付。
 */
function createSequentialDateText_(dateTimeText) {
  return normalizeString_(dateTimeText).slice(0, 10).replace(/-/g, '') || formatDate_(new Date(), 'yyyyMMdd');
}

/**
 * 接頭辞・日付・連番をID文字列に整形します。
 *
 * @param {string} prefix ID接頭辞。
 * @param {string} dateText yyyymmdd 形式の日付。
 * @param {number} number 連番。
 * @return {string} 整形済みID。
 */
function formatSequentialId_(prefix, dateText, number) {
  return `${normalizeString_(prefix)}-${dateText}-${String(number).padStart(3, '0')}`;
}

/**
 * 日付と時刻の文字列から Date を作成します。
 *
 * @param {string} usageDate yyyy-MM-dd 形式の日付。
 * @param {string} time HH:mm 形式の時刻。
 * @return {Date} 作成した日時。
 */
function createDateTime_(usageDate, time) {
  const dateParts = usageDate.split('-').map((value) => Number(value));
  const timeParts = time.split(':').map((value) => Number(value));
  return new Date(dateParts[0], dateParts[1] - 1, dateParts[2], timeParts[0], timeParts[1], 0);
}

/**
 * 2つの時間帯が重なっているかを判定します。
 *
 * @param {string} existingStart 既存開始時刻。
 * @param {string} existingEnd 既存終了時刻。
 * @param {string} newStart 新規開始時刻。
 * @param {string} newEnd 新規終了時刻。
 * @return {boolean} 重複している場合は true。
 */
function hasTimeOverlap_(existingStart, existingEnd, newStart, newEnd) {
  return toMinutes_(existingStart) < toMinutes_(newEnd) && toMinutes_(newStart) < toMinutes_(existingEnd);
}

/**
 * HH:mm 形式の時刻を 0:00 からの分数へ変換します。
 *
 * @param {string} time HH:mm 形式の時刻。
 * @return {number} 分数。
 */
function toMinutes_(time) {
  const normalizedTime = normalizeTimeString_(time);
  const parts = normalizedTime.split(':').map((value) => Number(value));
  if (parts.length !== 2 || Number.isNaN(parts[0]) || Number.isNaN(parts[1])) {
    return 0;
  }
  return parts[0] * 60 + parts[1];
}

/**
 * シートのデータ行を、ヘッダー名に対応したオブジェクト配列として取得します。
 *
 * @param {string} sheetName 対象シート名。
 * @return {Array<Object>} 行オブジェクト一覧。
 */
function selectSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const schemaKey = getSchemaKeyBySheetName_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn === 0) {
    return [];
  }
  const headers = getHeaderRow_(sheet);
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter((row) => row.some((cell) => normalizeString_(cell) !== ''))
    .map((row) => {
      const rowObject = {};
      columnKeys.forEach((columnKey, index) => {
        rowObject[columnKey] = row[index];
      });
      return rowObject;
    });
}

/**
 * シート名からスキーマキーを取得します。
 *
 * @param {string} sheetName 対象シート名。
 * @return {string} スキーマキー。
 */
function getSchemaKeyBySheetName_(sheetName) {
  const normalizedSheetName = normalizeString_(sheetName);
  const currentKey = Object.keys(SHEET_NAMES).find((key) => SHEET_NAMES[key] === normalizedSheetName);
  if (currentKey) {
    return currentKey;
  }
  const legacyKey = Object.keys(LEGACY_SHEET_NAMES).find((key) => LEGACY_SHEET_NAMES[key] === normalizedSheetName);
  return legacyKey || normalizedSheetName;
}

/**
 * 日本語ヘッダーを内部キーへ変換します。
 *
 * @param {string} schemaKey スキーマキー。
 * @param {Array<string>} headers シートのヘッダー配列。
 * @return {Array<string>} 内部キー配列。
 */
function getColumnKeysForHeaders_(schemaKey, headers) {
  const sheetHeaders = SHEET_HEADERS[schemaKey] || [];
  const columnKeys = SHEET_COLUMN_KEYS[schemaKey] || [];
  return headers.map((header) => {
    const headerIndex = sheetHeaders.indexOf(header);
    if (headerIndex !== -1) {
      return columnKeys[headerIndex] || header;
    }
    if (columnKeys.indexOf(header) !== -1) {
      return header;
    }
    return header;
  });
}

/**
 * 内部キーに対応する列番号を取得します。
 *
 * @param {string} sheetName 対象シート名。
 * @param {string} columnKey 内部列キー。
 * @return {number} 1始まりの列番号。見つからない場合は0。
 */
function getColumnIndex_(sheetName, columnKey) {
  const sheet = getSheet_(sheetName);
  const schemaKey = getSchemaKeyBySheetName_(sheetName);
  const headers = getHeaderRow_(sheet);
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);
  return columnKeys.indexOf(columnKey) + 1;
}

/**
 * 指定シートへ複数行を末尾追加します。
 *
 * @param {string} sheetName 対象シート名。
 * @param {Array<Array<*>>} rows 追加する行配列。
 */
function appendSheetRows_(sheetName, rows) {
  if (!rows || rows.length === 0) {
    return;
  }
  const sheet = getSheet_(sheetName);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * シート名から Sheet オブジェクトを取得します。
 *
 * @param {string} sheetName 対象シート名。
 * @return {Sheet} 対象シート。
 */
function getSheet_(sheetName) {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    const schemaKey = getSchemaKeyBySheetName_(sheetName);
    const legacySheetName = LEGACY_SHEET_NAMES[schemaKey];
    sheet = legacySheetName ? spreadsheet.getSheetByName(legacySheetName) : null;
  }
  if (!sheet) {
    throw new Error(`${sheetName}シートが見つかりません。setupSpreadsheet() を実行してください。`);
  }
  return sheet;
}

/**
 * 処理対象のスプレッドシートを取得します。
 *
 * @return {Spreadsheet} 対象スプレッドシート。
 */
function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY_KEY);
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }
  throw new Error(`スプレッドシートが見つかりません。スクリプトプロパティ ${SPREADSHEET_ID_PROPERTY_KEY} を設定してください。`);
}

/**
 * シートの1行目からヘッダー一覧を取得します。
 *
 * @param {Sheet} sheet 対象シート。
 * @return {Array<string>} ヘッダー一覧。
 */
function getHeaderRow_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return [];
  }
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((header) => normalizeString_(header));
}

/**
 * 設定シートの値を内部キーで扱えるオブジェクトに変換して取得します。
 *
 * @return {Object} 設定値マップ。
 */
function getSettings_() {
  const settings = {};
  Object.keys(DEFAULT_SETTINGS).forEach((key) => {
    settings[key] = DEFAULT_SETTINGS[key];
  });
  try {
    selectSheetObjects_(SHEET_NAMES.settings).forEach((row) => {
      const key = getSettingInternalKey_(row.setting_key);
      if (key) {
        settings[key] = normalizeString_(row.setting_value);
      }
    });
  } catch (_) {
  }
  return settings;
}

/**
 * 設定シート上の日本語キーまたは内部キーを内部キーへ変換します。
 *
 * @param {*} settingKey 設定キー。
 * @return {string} 内部キー。未対応の場合は空文字。
 */
function getSettingInternalKey_(settingKey) {
  const normalized = normalizeString_(settingKey);
  if (!normalized) {
    return '';
  }
  if (DEFAULT_SETTINGS[normalized] !== undefined) {
    return normalized;
  }
  const matched = Object.keys(SETTING_KEYS).find((key) => SETTING_KEYS[key] === normalized);
  return matched || '';
}

/**
 * 操作ログシートへ1行追加します。
 *
 * @param {string} actionType 操作種別。
 * @param {string} actor 操作者。
 * @param {string} targetId 対象ID。
 * @param {string} detail 操作内容。
 * @param {string} result 結果。
 * @param {string} errorMessage エラー内容。
 */
function writeOperationLog_(actionType, actor, targetId, detail, result, errorMessage) {
  try {
    appendSheetRows_(SHEET_NAMES.logs, [[
      createSequentialId_('LOG', formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss")),
      normalizeString_(actionType),
      normalizeString_(actor),
      normalizeString_(targetId),
      normalizeString_(detail),
      formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"),
      normalizeString_(result) || '成功',
      normalizeString_(errorMessage),
    ]]);
  } catch (_) {
  }
}

/**
 * カレンダーIDから表示用 URL を作成します。
 *
 * @param {string} calendarId Google カレンダーID。
 * @return {string} カレンダーURL。
 */
function createCalendarUrl_(calendarId) {
  const normalizedCalendarId = normalizeString_(calendarId);
  if (!normalizedCalendarId) {
    return '';
  }
  return `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(normalizedCalendarId)}`;
}

/**
 * 任意値を前後空白のない文字列へ変換します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み文字列。
 */
function normalizeString_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * 日付値を yyyy-MM-dd 形式へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み日付文字列。
 */
function normalizeDateString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, 'yyyy-MM-dd');
  }
  const stringValue = normalizeString_(value);
  if (!stringValue) {
    return '';
  }
  const match = stringValue.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (match) {
    return [match[1], pad2_(match[2]), pad2_(match[3])].join('-');
  }
  return stringValue.slice(0, 10);
}

/**
 * 時刻値を HH:mm 形式へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み時刻文字列。
 */
function normalizeTimeString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, 'HH:mm');
  }
  const stringValue = normalizeString_(value);
  if (!stringValue) {
    return '';
  }
  const match = stringValue.match(/^(\d{1,2}):(\d{1,2})/);
  if (match) {
    return `${pad2_(match[1])}:${pad2_(match[2])}`;
  }
  return stringValue;
}

/**
 * 日時値を保存用の yyyy-MM-dd'T'HH:mm:ss 形式へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み日時文字列。
 */
function normalizeDateTimeString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, "yyyy-MM-dd'T'HH:mm:ss");
  }
  return normalizeString_(value);
}

/**
 * フロント入力の日時値を保存用形式へ変換します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み日時文字列。
 */
function normalizeDateTimeInput_(value) {
  const text = normalizeString_(value);
  if (!text) {
    return '';
  }
  const parsed = parseDateTime_(text);
  return parsed ? formatDate_(parsed, "yyyy-MM-dd'T'HH:mm:ss") : text;
}

/**
 * 日時文字列や Date 値を Date オブジェクトへ変換します。
 *
 * @param {*} value 対象値。
 * @return {Date|null} 変換できた Date。失敗時は null。
 */
function parseDateTime_(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return value;
  }
  const normalized = normalizeString_(value).replace(/\//g, '-');
  const parsed = new Date(normalized);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed;
  }
  const plainMatch = normalized.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(?::(\d{2}))?$/);
  if (!plainMatch) {
    return null;
  }
  return new Date(`${plainMatch[1]}T${plainMatch[2]}:${plainMatch[3] || '00'}`);
}

/**
 * 予約ステータス表記を日本語の保存値へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済みステータス。
 */
function normalizeReservationStatus_(value) {
  const stringValue = normalizeString_(value).toLowerCase();
  if (!stringValue || stringValue === 'active' || stringValue === '有効') {
    return RESERVATION_STATUS.active;
  }
  if (stringValue === 'cancel_request' || stringValue === 'キャンセル依頼' || stringValue === 'キャンセル依頼中') {
    return RESERVATION_STATUS.cancelRequested;
  }
  if (stringValue === 'cancelled' || stringValue === 'canceled' || stringValue === '取消' || stringValue === 'キャンセル') {
    return RESERVATION_STATUS.cancelled;
  }
  return normalizeString_(value);
}

/**
 * 有効状態の表記を「有効」または「無効」へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 正規化済み状態。
 */
function normalizeStatus_(value) {
  const text = normalizeString_(value);
  if (!text) {
    return '';
  }
  if (text === '有効' || text.toLowerCase() === 'true') {
    return '有効';
  }
  if (text === '無効' || text.toLowerCase() === 'false') {
    return '無効';
  }
  return text;
}

/**
 * TRUE/FALSE、有効/無効などの表記を boolean へ変換します。
 *
 * @param {*} value 対象値。
 * @return {boolean} 真として扱える場合は true。
 */
function normalizeBoolean_(value) {
  const stringValue = normalizeString_(value).toLowerCase();
  return value === true
    || stringValue === 'true'
    || stringValue === '1'
    || stringValue === 'yes'
    || stringValue === '有効'
    || stringValue === '利用可';
}

/**
 * 数値や文字列を2桁の文字列へ整形します。
 *
 * @param {*} value 対象値。
 * @return {string} 2桁文字列。
 */
function pad2_(value) {
  return String(value).padStart(2, '0');
}

/**
 * 設定シートまたは GAS 設定からタイムゾーンを取得します。
 *
 * @return {string} タイムゾーンID。
 */
function getScriptTimeZone_() {
  const settings = (() => {
    try {
      return getSettings_();
    } catch (_) {
      return DEFAULT_SETTINGS;
    }
  })();
  return settings.TIMEZONE || Session.getScriptTimeZone() || 'Asia/Tokyo';
}

/**
 * Date を指定パターンで文字列化します。
 *
 * @param {Date} date 対象日時。
 * @param {string} pattern Utilities.formatDate の書式。
 * @return {string} フォーマット済み日時。
 */
function formatDate_(date, pattern) {
  return Utilities.formatDate(date, getScriptTimeZone_(), pattern);
}

/**
 * 文字列として値を比較します。
 *
 * @param {*} leftValue 左辺。
 * @param {*} rightValue 右辺。
 * @return {number} 比較結果。
 */
function compareValues_(leftValue, rightValue) {
  const leftText = normalizeString_(leftValue);
  const rightText = normalizeString_(rightValue);
  if (leftText < rightText) {
    return -1;
  }
  if (leftText > rightText) {
    return 1;
  }
  return 0;
}

/**
 * オブジェクトを JSON の TextOutput として返します。
 *
 * @param {Object} data レスポンスデータ。
 * @return {TextOutput} JSON レスポンス。
 */
function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
