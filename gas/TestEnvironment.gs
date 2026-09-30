/**
 * 予約 v2 の隔離テスト環境を初期化します。
 *
 * <p>名前が [TEST] で始まるスプレッドシートでのみ実行できます。
 * 本番カレンダーは参照せず、非公開のテスト用カレンダーを3件作成します。</p>
 *
 * @return {Object} テスト環境情報。
 */
function setupReservationV2TestEnvironment() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet || spreadsheet.getName().indexOf('[TEST]') !== 0) {
    throw new Error('名前が [TEST] で始まるテスト用スプレッドシートで実行してください。');
  }

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(SPREADSHEET_ID_PROPERTY_KEY, spreadsheet.getId());
  properties.setProperty('RESERVATION_V2_TEST_ENVIRONMENT', 'TRUE');
  properties.setProperty(RESERVATION_V2_MODE_PROPERTY_KEY, RESERVATION_V2_MODES.legacy);
  properties.setProperty(RESERVATION_V2_ENABLED_PROPERTY_KEY, 'FALSE');
  properties.setProperty(RESERVATION_V2_INDEX_READY_PROPERTY_KEY, 'FALSE');
  spreadsheetExecutionCache_ = spreadsheet;

  ['reservations', 'rooms', 'settings', 'logs'].forEach((schemaKey) => {
    ensureSheetAndHeader_(spreadsheet, schemaKey);
  });

  const reservationRows = selectSheetObjects_(SHEET_NAMES.reservations);
  if (reservationRows.length > 0) {
    throw new Error('予約シートに既存データがあるため初期化を中止しました。新規の空スプレッドシートを使用してください。');
  }

  ensureDefaultSettings_();
  const roomsSheet = getSheet_(SHEET_NAMES.rooms);
  const existingRooms = selectSheetObjects_(SHEET_NAMES.rooms);
  if (existingRooms.length > 0) {
    throw new Error('会議室シートに既存データがあるため初期化を中止しました。');
  }
  const calendarIds = ensureReservationV2TestCalendars_();
  roomsSheet.getRange(2, 1, 3, SHEET_HEADERS.rooms.length).setValues([
    ['TEST-ROOM-A', '[TEST] 会議室A', calendarIds.roomA, '有効', 1, '単体', ''],
    ['TEST-ROOM-B', '[TEST] 会議室B', calendarIds.roomB, '有効', 2, '単体', ''],
    ['TEST-ROOM-ALL', '[TEST] 大会議室', calendarIds.roomAll, '有効', 3, '大会議室', 'TEST-ROOM-A,TEST-ROOM-B'],
  ]);
  invalidateReservationStaticCacheV2_();
  const release = enableReservationV2Release();

  return {
    ok: true,
    spreadsheet_id: spreadsheet.getId(),
    spreadsheet_url: spreadsheet.getUrl(),
    calendars: calendarIds,
    rooms: ['TEST-ROOM-A', 'TEST-ROOM-B', 'TEST-ROOM-ALL'],
    reservation_release: release,
  };
}

/**
 * テスト用カレンダーを一度だけ作成します。
 *
 * @return {Object} 会議室別カレンダーID。
 */
function ensureReservationV2TestCalendars_() {
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = 'RESERVATION_V2_TEST_CALENDARS';
  const saved = properties.getProperty(propertyKey);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      if (parsed.roomA && parsed.roomB && parsed.roomAll) {
        return parsed;
      }
    } catch (_) {
    }
  }
  const suffix = Utilities.getUuid().slice(0, 8);
  const roomA = CalendarApp.createCalendar(`[TEST ${suffix}] 会議室A`, { timeZone: getScriptTimeZone_() });
  const roomB = CalendarApp.createCalendar(`[TEST ${suffix}] 会議室B`, { timeZone: getScriptTimeZone_() });
  const roomAll = CalendarApp.createCalendar(`[TEST ${suffix}] 大会議室`, { timeZone: getScriptTimeZone_() });
  const result = {
    roomA: roomA.getId(),
    roomB: roomB.getId(),
    roomAll: roomAll.getId(),
  };
  properties.setProperty(propertyKey, JSON.stringify(result));
  return result;
}

/**
 * 現在のスクリプトが隔離テスト環境を向いていることを検証します。
 *
 * @return {Object} シート・カレンダー・予約件数。
 */
function verifyReservationV2TestEnvironment() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('RESERVATION_V2_TEST_ENVIRONMENT') !== 'TRUE') {
    throw new Error('テスト環境フラグが設定されていません。');
  }
  const spreadsheet = getSpreadsheet_();
  if (spreadsheet.getName().indexOf('[TEST]') !== 0) {
    throw new Error('テスト用スプレッドシート以外を参照しています。');
  }
  const rooms = selectActiveRooms_();
  const nonTestRoom = rooms.find((room) => room.room_id.indexOf('TEST-') !== 0);
  if (nonTestRoom) {
    throw new Error(`テスト用以外の会議室を検出しました: ${nonTestRoom.room_id}`);
  }
  return {
    ok: true,
    spreadsheet_id: spreadsheet.getId(),
    spreadsheet_name: spreadsheet.getName(),
    room_count: rooms.length,
    reservation_count: selectSheetObjects_(SHEET_NAMES.reservations).length,
    calendars: rooms.map((room) => ({ room_id: room.room_id, calendar_id: room.calendar_id })),
  };
}

/**
 * 100回テスト後の予約行を、個人情報を含めず集計用に返します。
 *
 * @param {Object} params run_id を含む GET パラメータ。
 * @return {Object} 該当テストの予約行。
 */
function handleGetReservationV2TestSummary_(params) {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('RESERVATION_V2_TEST_ENVIRONMENT') !== 'TRUE') {
    throw new Error('テスト環境以外では実行できません。');
  }
  const runId = normalizeString_(params && (params.run_id || params.runId));
  if (!runId) {
    return { ok: false, error: 'run_id が必要です。' };
  }
  const prefix = `${runId}-`;
  const rows = selectSheetObjects_(SHEET_NAMES.reservations)
    .map(normalizeReservationRow_)
    .filter((reservation) => reservation.request_id.indexOf(prefix) === 0)
    .map((reservation) => ({
      reservation_id: reservation.reservation_id,
      request_id: reservation.request_id,
      room_id: reservation.room_id,
      usage_date: reservation.usage_date,
      start_time: reservation.start_time,
      end_time: reservation.end_time,
      status: reservation.status,
      event_ref_count: reservation.calendar_event_id
        ? reservation.calendar_event_id.split(';').filter(Boolean).length
        : 0,
      error_message: reservation.error_message,
    }));
  return { ok: true, run_id: runId, rows };
}
