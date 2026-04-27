/**
 * GAS 書き込みAPIエンドポイント
 * フロントエンドからPOSTリクエストを受け取り、スプレッドシートに書き込む
 *
 * 対応アクション:
 *   insert_reservations  - 予約登録
 *   cancel_request       - キャンセル依頼
 *   admin_delete         - 管理者による予約削除
 *   get_config           - 設定・会議室情報取得（初期化用）
 */

/**
 * GETリクエスト（フロントからの設定取得用）
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || '';

  if (action === 'get_config') {
    return jsonResponse_(handleGetConfig_());
  }

  return jsonResponse_({ ok: false, error: 'Unknown action' });
}

/**
 * POSTリクエスト処理
 */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonResponse_({ ok: false, error: 'Invalid JSON' });
  }

  const action = body.action || '';

  try {
    if (action === 'insert_reservations') {
      return jsonResponse_(handleInsertReservations_(body.payload));
    }
    if (action === 'cancel_request') {
      return jsonResponse_(handleCancelRequest_(body.payload));
    }
    if (action === 'admin_delete') {
      return jsonResponse_(handleAdminDelete_(body.payload));
    }
    return jsonResponse_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    writeLog_('ERROR', action, err.message, { stack: err.stack || '' });
    return jsonResponse_({ ok: false, error: err.message });
  }
}

// ---------------------------------------------------------------------------
// ハンドラ
// ---------------------------------------------------------------------------

/**
 * アプリ設定と会議室一覧を返す（初回ロード用）
 */
function handleGetConfig_() {
  const settings = getSettings_();
  const rooms = selectActiveRooms_();
  const today = formatDate_(new Date(), 'yyyy-MM-dd');

  return {
    ok: true,
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
    },
    rooms,
    today,
  };
}

/**
 * 予約登録
 */
function handleInsertReservations_(payload) {
  return insertReservations(payload);
}

/**
 * キャンセル依頼をスプレッドシートに記録する
 * ステータスを「キャンセル依頼」に更新する
 */
function handleCancelRequest_(payload) {
  const reservationId = normalizeString_(payload && payload.reservation_id);
  const reason = normalizeString_(payload && payload.reason);

  if (!reservationId) {
    return { ok: false, error: '予約IDが必要です。' };
  }

  const sheet = getSheet_(SHEET_NAMES.reservations);
  const schemaKey = 'reservations';
  const headers = getHeaderRow_(sheet);
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);

  const reservationIdCol = columnKeys.indexOf('reservation_id') + 1;
  const statusCol = columnKeys.indexOf('status') + 1;

  if (!reservationIdCol || !statusCol) {
    return { ok: false, error: 'シートの構造を確認してください。' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: false, error: '予約が見つかりません。' };
  }

  const idValues = sheet.getRange(2, reservationIdCol, lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < idValues.length; i++) {
    if (normalizeString_(idValues[i][0]) === reservationId) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    return { ok: false, error: '予約が見つかりません。' };
  }

  const currentStatus = normalizeString_(sheet.getRange(targetRow, statusCol).getValue());
  if (currentStatus !== RESERVATION_STATUS.active) {
    return { ok: false, error: 'この予約はすでにキャンセル済みまたは依頼済みです。' };
  }

  sheet.getRange(targetRow, statusCol).setValue('キャンセル依頼');

  writeLog_('INFO', 'cancelRequest', 'キャンセル依頼を受け付けました。', {
    reservation_id: reservationId,
    reason,
  });

  return { ok: true, reservation_id: reservationId };
}

/**
 * 管理者による予約削除（ステータスを「取消」に変更）
 */
function handleAdminDelete_(payload) {
  const adminPassword = normalizeString_(payload && payload.admin_password);
  const reservationIds = Array.isArray(payload && payload.reservation_ids)
    ? payload.reservation_ids.map((id) => normalizeString_(id))
    : [];

  // パスワード検証
  const settings = getSettings_();
  const correctPassword = normalizeString_(settings.ADMIN_PASSWORD || DEFAULT_SETTINGS.ADMIN_PASSWORD || '');
  if (!correctPassword || adminPassword !== correctPassword) {
    return { ok: false, error: 'パスワードが違います。' };
  }

  if (reservationIds.length === 0) {
    return { ok: false, error: '削除する予約IDが必要です。' };
  }

  const sheet = getSheet_(SHEET_NAMES.reservations);
  const schemaKey = 'reservations';
  const headers = getHeaderRow_(sheet);
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);

  const reservationIdCol = columnKeys.indexOf('reservation_id') + 1;
  const statusCol = columnKeys.indexOf('status') + 1;
  const calendarEventIdCol = columnKeys.indexOf('calendar_event_id') + 1;
  const roomIdCol = columnKeys.indexOf('room_id') + 1;

  if (!reservationIdCol || !statusCol) {
    return { ok: false, error: 'シートの構造を確認してください。' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: false, error: '予約が見つかりません。' };
  }

  const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const deletedIds = [];
  const notFoundIds = [];

  reservationIds.forEach((reservationId) => {
    let found = false;
    for (let i = 0; i < allData.length; i++) {
      if (normalizeString_(allData[i][reservationIdCol - 1]) === reservationId) {
        found = true;
        // カレンダーイベントを削除
        if (calendarEventIdCol && roomIdCol) {
          const calendarEventId = normalizeString_(allData[i][calendarEventIdCol - 1]);
          const roomId = normalizeString_(allData[i][roomIdCol - 1]);
          if (calendarEventId && roomId) {
            deleteCalendarEvent_(roomId, calendarEventId);
          }
        }
        // ステータスを「取消」に更新
        sheet.getRange(i + 2, statusCol).setValue(RESERVATION_STATUS.cancelled);
        allData[i][statusCol - 1] = RESERVATION_STATUS.cancelled;
        deletedIds.push(reservationId);
        break;
      }
    }
    if (!found) {
      notFoundIds.push(reservationId);
    }
  });

  writeLog_('INFO', 'adminDelete', '管理者が予約を削除しました。', {
    deleted: deletedIds,
    notFound: notFoundIds,
  });

  return {
    ok: true,
    deleted: deletedIds,
    notFound: notFoundIds,
  };
}

// ---------------------------------------------------------------------------
// ヘルパー
// ---------------------------------------------------------------------------

/**
 * Googleカレンダーのイベントを削除する
 */
function deleteCalendarEvent_(roomId, calendarEventId) {
  try {
    const rooms = selectActiveRooms_();
    const room = rooms.find((r) => r.room_id === roomId);
    if (!room || !room.calendar_id) return;

    const calendar = CalendarApp.getCalendarById(room.calendar_id);
    if (!calendar) return;

    const event = calendar.getEventById(calendarEventId);
    if (event) event.deleteEvent();
  } catch (err) {
    writeLog_('WARN', 'deleteCalendarEvent', err.message, { roomId, calendarEventId });
  }
}

/**
 * JSONレスポンスを返す
 */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
