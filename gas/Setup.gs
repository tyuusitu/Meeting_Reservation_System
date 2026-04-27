/**
 * スプレッドシート上のシート名を一元管理します。
 *
 * <p>フロントの config.js の SHEET_NAME と揃えておくことで、
 * CSV 読み込みと GAS 書き込みが同じシートを参照します。</p>
 */
const SHEET_NAMES = {
  reservations: '予約',
  rooms: '会議室',
  settings: '設定',
  members: 'メンバー一覧',
  meetings: '会議一覧',
  attendanceResponses: '出欠回答',
  meetingAggregations: '会議別集計',
  memberAggregations: '個人別集計',
  departmentAggregations: '部局別集計',
  streaks: '連続欠席チェック',
  logs: '操作ログ',
};

/**
 * 旧バージョンで使っていた英語シート名を日本語シート名へ移行するための対応表です。
 */
const LEGACY_SHEET_NAMES = {
  reservations: 'reservations',
  rooms: 'rooms',
  settings: 'settings',
  logs: 'ログ',
};

/**
 * 各シートの日本語ヘッダー定義です。
 *
 * <p>セットアップ時に1行目へ書き込み、非エンジニアがシートを直接編集しても
 * 意味が分かる列名にしています。</p>
 */
const SHEET_HEADERS = {
  reservations: [
    '予約ID',
    '利用者ID',
    'メールアドレス',
    '団体名',
    '会議名',
    'お名前',
    '会議室ID',
    '会議室名',
    '利用日',
    '開始時間',
    '終了時間',
    'カレンダー予定ID',
    'ステータス',
    '作成日時',
  ],
  rooms: [
    '会議室ID',
    '会議室名',
    'カレンダーID',
    '利用可能',
    '表示順',
  ],
  settings: [
    '設定キー',
    '設定値',
    '説明',
    '更新日時',
  ],
  members: [
    'メンバーID',
    '名前',
    '部局',
    '学年',
    '役職',
    '表示順',
    '有効状態',
    '登録日時',
    '更新日時',
  ],
  meetings: [
    '会議ID',
    '会議名',
    '会議日',
    '開始時間',
    '終了時間',
    '募集開始日時',
    '募集終了日時',
    '回答締切日時',
    '備考',
    '有効状態',
    '登録日時',
    '更新日時',
  ],
  attendanceResponses: [
    '回答ID',
    '会議ID',
    'メンバーID',
    '名前',
    '部局',
    '学年',
    '役職',
    '出欠',
    '欠席理由',
    '欠席理由詳細',
    '回答日時',
    '更新日時',
    '更新回数',
  ],
  meetingAggregations: [
    '会議ID',
    '会議名',
    '会議日',
    '対象人数',
    '出席回答数',
    '欠席回答数',
    '未提出数',
    '出席扱い人数',
    '参加率',
    '欠席率',
    '最終更新日時',
  ],
  memberAggregations: [
    'メンバーID',
    '名前',
    '部局',
    '学年',
    '役職',
    '対象会議数',
    '出席回答数',
    '欠席回答数',
    '未提出数',
    '出席扱い回数',
    '参加率',
    '連続欠席回数',
    '最終回答日時',
    '最終更新日時',
  ],
  departmentAggregations: [
    '部局',
    '所属人数',
    '対象会議数',
    '出席回答数',
    '欠席回答数',
    '未提出数',
    '出席扱い回数',
    '参加率',
    '欠席率',
    '最終更新日時',
  ],
  streaks: [
    'メンバーID',
    '名前',
    '部局',
    '学年',
    '役職',
    '連続欠席回数',
    '注意レベル',
    '直近欠席会議',
    '直近欠席理由',
    '最終更新日時',
  ],
  logs: [
    'ログID',
    '操作種別',
    '操作者',
    '対象ID',
    '内容',
    '操作日時',
    '結果',
    'エラー内容',
  ],
};

/**
 * 日本語ヘッダーを GAS 内部で扱いやすいキーへ変換する対応表です。
 */
const SHEET_COLUMN_KEYS = {
  reservations: [
    'reservation_id',
    'line_user_id',
    'email',
    'organization_name',
    'meeting_name',
    'user_name',
    'room_id',
    'room_name',
    'usage_date',
    'start_time',
    'end_time',
    'calendar_event_id',
    'status',
    'created_at',
  ],
  rooms: [
    'room_id',
    'room_name',
    'calendar_id',
    'is_active',
    'display_order',
  ],
  settings: [
    'setting_key',
    'setting_value',
    'description',
    'updated_at',
  ],
  members: [
    'member_id',
    'name',
    'department',
    'grade',
    'role',
    'display_order',
    'status',
    'created_at',
    'updated_at',
  ],
  meetings: [
    'meeting_id',
    'meeting_name',
    'meeting_date',
    'start_time',
    'end_time',
    'recruit_start_at',
    'recruit_end_at',
    'answer_deadline_at',
    'note',
    'status',
    'created_at',
    'updated_at',
  ],
  attendanceResponses: [
    'answer_id',
    'meeting_id',
    'member_id',
    'name',
    'department',
    'grade',
    'role',
    'attendance',
    'absence_reason',
    'absence_detail',
    'answered_at',
    'updated_at',
    'update_count',
  ],
  meetingAggregations: [
    'meeting_id',
    'meeting_name',
    'meeting_date',
    'target_count',
    'attend_count',
    'absence_count',
    'unanswered_count',
    'effective_attend_count',
    'attendance_rate',
    'absence_rate',
    'updated_at',
  ],
  memberAggregations: [
    'member_id',
    'name',
    'department',
    'grade',
    'role',
    'target_meeting_count',
    'attend_count',
    'absence_count',
    'unanswered_count',
    'effective_attend_count',
    'attendance_rate',
    'absence_streak',
    'last_answered_at',
    'updated_at',
  ],
  departmentAggregations: [
    'department',
    'member_count',
    'target_meeting_count',
    'attend_count',
    'absence_count',
    'unanswered_count',
    'effective_attend_count',
    'attendance_rate',
    'absence_rate',
    'updated_at',
  ],
  streaks: [
    'member_id',
    'name',
    'department',
    'grade',
    'role',
    'absence_streak',
    'alert_level',
    'latest_absence_meeting_name',
    'latest_absence_reason',
    'updated_at',
  ],
  logs: [
    'log_id',
    'action_type',
    'actor',
    'target_id',
    'detail',
    'acted_at',
    'result',
    'error_message',
  ],
};

/**
 * 設定シートで表示する日本語設定キーです。
 */
const SETTING_KEYS = {
  SYSTEM_NAME: 'システム名',
  ADMIN_PASSWORD: '管理者パスワード',
  OTHER_REASON_DETAIL_REQUIRED: 'その他理由詳細必須',
  ABSENCE_DETAIL_ALWAYS_REQUIRED: '欠席詳細常時必須',
  ALERT_STREAK_COUNT: '注意連続欠席回数',
  CRITICAL_STREAK_COUNT: '要確認連続欠席回数',
  TIMEZONE: 'タイムゾーン',
  BUSINESS_START_TIME: '利用開始時刻',
  BUSINESS_END_TIME: '利用終了時刻',
  TIME_SLOT_MINUTES: '時間刻み分',
  MAX_RESERVATIONS_PER_SUBMIT: '同時登録上限',
  LABEL_ORGANIZATION: '団体名ラベル',
  LABEL_USER_NAME: 'お名前ラベル',
  LABEL_MEETING_NAME: '会議名ラベル',
  LABEL_ROOM: '会議室ラベル',
  FIELD_ORGANIZATION: '団体名入力',
  FIELD_USER_NAME: 'お名前入力',
  SPREADSHEET_URL: 'スプレッドシートURL',
};

/**
 * 設定シートへ初期投入する既定値です。
 */
const DEFAULT_SETTINGS = {
  SYSTEM_NAME: '会議室予約・出欠管理システム',
  ADMIN_PASSWORD: 'admin1234',
  OTHER_REASON_DETAIL_REQUIRED: 'TRUE',
  ABSENCE_DETAIL_ALWAYS_REQUIRED: 'FALSE',
  ALERT_STREAK_COUNT: '2',
  CRITICAL_STREAK_COUNT: '3',
  TIMEZONE: 'Asia/Tokyo',
  BUSINESS_START_TIME: '09:00',
  BUSINESS_END_TIME: '22:00',
  TIME_SLOT_MINUTES: '5',
  MAX_RESERVATIONS_PER_SUBMIT: '10',
  LABEL_ORGANIZATION: '団体名',
  LABEL_USER_NAME: 'お名前',
  LABEL_MEETING_NAME: '会議名',
  LABEL_ROOM: '会議室',
  FIELD_ORGANIZATION: '有効',
  FIELD_USER_NAME: '有効',
  SPREADSHEET_URL: '',
};

/**
 * 設定シートの説明列へ初期投入する説明文です。
 */
const DEFAULT_SETTING_DESCRIPTIONS = {
  SYSTEM_NAME: '画面上に表示するシステム名です。',
  ADMIN_PASSWORD: '管理者画面に入るための簡易パスワードです。必ず変更してください。',
  OTHER_REASON_DETAIL_REQUIRED: '欠席理由がその他の場合に詳細入力を必須にするか。TRUE/FALSE。',
  ABSENCE_DETAIL_ALWAYS_REQUIRED: '欠席時に詳細入力を常に必須にするか。TRUE/FALSE。',
  ALERT_STREAK_COUNT: '連続欠席チェックで「注意」にする回数です。',
  CRITICAL_STREAK_COUNT: '連続欠席チェックで「要確認」にする回数です。',
  TIMEZONE: '日時判定に使うタイムゾーンです。通常は Asia/Tokyo。',
  BUSINESS_START_TIME: '予約画面で使う開始時刻です。HH:mm 形式。',
  BUSINESS_END_TIME: '予約画面で使う終了時刻です。HH:mm 形式。',
  TIME_SLOT_MINUTES: '予約画面の時間刻み分です。',
  MAX_RESERVATIONS_PER_SUBMIT: '予約画面で一度に登録できる件数上限です。',
  LABEL_ORGANIZATION: '予約フォームの団体名欄ラベルです。',
  LABEL_USER_NAME: '予約フォームのお名前欄ラベルです。',
  LABEL_MEETING_NAME: '予約フォームの会議名欄ラベルです。',
  LABEL_ROOM: '予約画面で表示する会議室の呼び名です。',
  FIELD_ORGANIZATION: '予約フォームで団体名欄を表示するか。有効/無効。',
  FIELD_USER_NAME: '予約フォームでお名前欄を表示するか。有効/無効。',
  SPREADSHEET_URL: 'このスプレッドシートの URL です。必要に応じて記入してください。',
};

const DEFAULT_ROOM_ROWS = [
  ['room-1', '中執内応接室', '', '有効', 1],
  ['room-2', '階段下会議室', '', '有効', 2],
  ['room-3', '中執前会議室', '', '有効', 3],
];

const DEFAULT_MEMBER_ROWS = [
  ['MEM-001', '山田太郎', '企画管理部', '3回生', '部長', 1, '有効', '', ''],
  ['MEM-002', '佐藤花子', '広報部', '2回生', '副部長', 2, '有効', '', ''],
  ['MEM-003', '鈴木一郎', '会計部', '1回生', '部員', 3, '有効', '', ''],
];

const DEFAULT_MEETING_ROWS = [
  ['MTG-20260430-001', '4月定例会', '2026-04-30', '18:00', '19:30', '2026-04-20T00:00:00', '2026-04-30T18:00:00', '2026-04-30T17:00:00', '必ず回答してください', '有効', '', ''],
];

const SPREADSHEET_ID_PROPERTY_KEY = 'SPREADSHEET_ID';
const ATTENDANCE_OPTIONS = { attend: '出席', absence: '欠席' };
const ABSENCE_REASONS = ['授業', 'アルバイト', '体調不良', '家庭の都合', 'その他'];
const RESERVATION_STATUS = { active: '有効', cancelRequested: 'キャンセル依頼', cancelled: '取消' };

/**
 * セットアップ関数の短い別名です。
 *
 * @return {Object} 初期化結果。
 */
function setupSpreadsheet() {
  return initializeSpreadsheet();
}

/**
 * 必要なシート、ヘッダー、初期データ、書式、集計をまとめて作成します。
 *
 * @return {Object} 初期化結果。
 */
function initializeSpreadsheet() {
  const spreadsheet = getSpreadsheet_();
  Object.keys(SHEET_NAMES).forEach((schemaKey) => ensureSheetAndHeader_(spreadsheet, schemaKey));
  ensureDefaultSettings_();
  ensureDefaultRooms_();
  ensureDefaultMembers_();
  ensureDefaultMeetings_();
  applySheetFormatting_();
  refreshAttendanceAggregations_();
  writeOperationLog_('セットアップ', '管理者', spreadsheet.getId(), 'スプレッドシート初期化を実行しました。', '成功', '');

  return {
    ok: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheets: Object.keys(SHEET_NAMES).map((schemaKey) => SHEET_NAMES[schemaKey]),
  };
}

/**
 * 現在開いているスプレッドシートIDをスクリプトプロパティへ保存します。
 *
 * @return {Object} 保存結果。
 */
function setSpreadsheetIdToCurrent() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('アクティブなスプレッドシートが見つかりません。スプレッドシートから Apps Script を開いて実行してください。');
  }
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY_KEY, spreadsheet.getId());
  return { ok: true, spreadsheetId: spreadsheet.getId(), spreadsheetUrl: spreadsheet.getUrl() };
}

/**
 * 会議室シートでカレンダーIDが未設定の有効会議室に Google カレンダーを作成します。
 *
 * @return {Object} 作成したカレンダー一覧。
 */
function createRoomCalendarsAndUpdateRooms() {
  const sheet = getSheet_(SHEET_NAMES.rooms);
  const roomNameColumn = getColumnIndex_(SHEET_NAMES.rooms, 'room_name');
  const calendarIdColumn = getColumnIndex_(SHEET_NAMES.rooms, 'calendar_id');
  const isActiveColumn = getColumnIndex_(SHEET_NAMES.rooms, 'is_active');
  const lastRow = sheet.getLastRow();
  const createdCalendars = [];

  if (!roomNameColumn || !calendarIdColumn || !isActiveColumn) {
    throw new Error('会議室シートのヘッダーを確認してください。');
  }

  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber += 1) {
    const roomName = normalizeString_(sheet.getRange(rowNumber, roomNameColumn).getValue());
    const calendarId = normalizeString_(sheet.getRange(rowNumber, calendarIdColumn).getValue());
    const isActive = normalizeBoolean_(sheet.getRange(rowNumber, isActiveColumn).getValue());
    if (!roomName || calendarId || !isActive) {
      continue;
    }
    const calendar = CalendarApp.createCalendar(`${roomName} 予約`, {
      timeZone: getScriptTimeZone_(),
    });
    sheet.getRange(rowNumber, calendarIdColumn).setValue(calendar.getId());
    createdCalendars.push({ room_name: roomName, calendar_id: calendar.getId() });
  }

  if (createdCalendars.length > 0) {
    writeOperationLog_('会議室カレンダー作成', '管理者', '', `${createdCalendars.length}件のカレンダーを作成しました。`, '成功', '');
  }
  return { ok: true, createdCalendars };
}

/**
 * 指定スキーマのシートを作成し、ヘッダーと基本書式を設定します。
 *
 * @param {Spreadsheet} spreadsheet 対象スプレッドシート。
 * @param {string} schemaKey SHEET_NAMES のキー。
 */
function ensureSheetAndHeader_(spreadsheet, schemaKey) {
  const sheetName = SHEET_NAMES[schemaKey];
  const headers = SHEET_HEADERS[schemaKey];
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    const legacySheetName = LEGACY_SHEET_NAMES[schemaKey];
    const legacySheet = legacySheetName ? spreadsheet.getSheetByName(legacySheetName) : null;
    if (legacySheet) {
      legacySheet.setName(sheetName);
      sheet = legacySheet;
    } else {
      sheet = spreadsheet.insertSheet(sheetName);
    }
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#EAF2FF');
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).setNumberFormat('@');
  }
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 設定シートに不足している初期設定を追加し、説明や更新日時を補完します。
 */
function ensureDefaultSettings_() {
  const sheet = getSheet_(SHEET_NAMES.settings);
  const rows = selectSheetObjects_(SHEET_NAMES.settings);
  const existingMap = {};
  rows.forEach((row) => {
    const internalKey = getSettingInternalKey_(row.setting_key);
    if (internalKey) {
      existingMap[internalKey] = row;
    }
  });

  const appendRows = [];
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  Object.keys(DEFAULT_SETTINGS).forEach((key) => {
    if (existingMap[key]) {
      return;
    }
    appendRows.push([
      SETTING_KEYS[key] || key,
      DEFAULT_SETTINGS[key],
      DEFAULT_SETTING_DESCRIPTIONS[key] || '',
      now,
    ]);
  });

  if (appendRows.length > 0) {
    appendSheetRows_(SHEET_NAMES.settings, appendRows);
  }

  const keyColumn = getColumnIndex_(SHEET_NAMES.settings, 'setting_key');
  const descColumn = getColumnIndex_(SHEET_NAMES.settings, 'description');
  const updatedAtColumn = getColumnIndex_(SHEET_NAMES.settings, 'updated_at');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  values.forEach((row, index) => {
    const internalKey = getSettingInternalKey_(row[keyColumn - 1]);
    if (!internalKey) {
      return;
    }
    if (row[keyColumn - 1] !== SETTING_KEYS[internalKey]) {
      sheet.getRange(index + 2, keyColumn).setValue(SETTING_KEYS[internalKey]);
    }
    if (!normalizeString_(row[descColumn - 1])) {
      sheet.getRange(index + 2, descColumn).setValue(DEFAULT_SETTING_DESCRIPTIONS[internalKey] || '');
    }
    if (!normalizeString_(row[updatedAtColumn - 1])) {
      sheet.getRange(index + 2, updatedAtColumn).setValue(now);
    }
  });
}

/**
 * 会議室シートが空の場合だけ、初期会議室を追加します。
 */
function ensureDefaultRooms_() {
  const existingRows = selectSheetObjects_(SHEET_NAMES.rooms).filter((row) => normalizeString_(row.room_id));
  if (existingRows.length === 0) {
    appendSheetRows_(SHEET_NAMES.rooms, DEFAULT_ROOM_ROWS);
  }
}

/**
 * メンバー一覧シートが空の場合だけ、サンプルメンバーを追加します。
 */
function ensureDefaultMembers_() {
  const existingRows = selectSheetObjects_(SHEET_NAMES.members).filter((row) => normalizeString_(row.member_id));
  if (existingRows.length > 0) {
    return;
  }
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  appendSheetRows_(SHEET_NAMES.members, DEFAULT_MEMBER_ROWS.map((row) => [row[0], row[1], row[2], row[3], row[4], row[5], row[6], now, now]));
}

/**
 * 会議一覧シートが空の場合だけ、サンプル会議を追加します。
 */
function ensureDefaultMeetings_() {
  const existingRows = selectSheetObjects_(SHEET_NAMES.meetings).filter((row) => normalizeString_(row.meeting_id));
  if (existingRows.length > 0) {
    return;
  }
  const now = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
  appendSheetRows_(SHEET_NAMES.meetings, DEFAULT_MEETING_ROWS.map((row) => row.slice(0, 10).concat([now, now])));
}

/**
 * 日付、日時、パーセント列の表示形式をまとめて適用します。
 */
function applySheetFormatting_() {
  applyFormattingForSheet_(SHEET_NAMES.settings, ['updated_at']);
  applyFormattingForSheet_(SHEET_NAMES.members, ['created_at', 'updated_at']);
  applyFormattingForSheet_(SHEET_NAMES.meetings, ['meeting_date'], ['recruit_start_at', 'recruit_end_at', 'answer_deadline_at', 'created_at', 'updated_at']);
  applyFormattingForSheet_(SHEET_NAMES.attendanceResponses, [], ['answered_at', 'updated_at']);
  applyFormattingForSheet_(SHEET_NAMES.meetingAggregations, [], ['updated_at'], ['attendance_rate', 'absence_rate']);
  applyFormattingForSheet_(SHEET_NAMES.memberAggregations, [], ['last_answered_at', 'updated_at'], ['attendance_rate']);
  applyFormattingForSheet_(SHEET_NAMES.departmentAggregations, [], ['updated_at'], ['attendance_rate', 'absence_rate']);
  applyFormattingForSheet_(SHEET_NAMES.streaks, [], ['updated_at']);
  applyFormattingForSheet_(SHEET_NAMES.logs, [], ['acted_at']);
  applyFormattingForSheet_(SHEET_NAMES.reservations, ['usage_date'], ['created_at']);
}

/**
 * 指定シートの列キーに応じて日付・日時・パーセント形式を設定します。
 *
 * @param {string} sheetName 対象シート名。
 * @param {Array<string>} dateKeys 日付列の内部キー一覧。
 * @param {Array<string>} dateTimeKeys 日時列の内部キー一覧。
 * @param {Array<string>} percentKeys パーセント列の内部キー一覧。
 */
function applyFormattingForSheet_(sheetName, dateKeys, dateTimeKeys, percentKeys) {
  const sheet = getSheet_(sheetName);
  const lastRow = Math.max(sheet.getMaxRows(), 2);
  (dateKeys || []).forEach((key) => {
    const column = getColumnIndex_(sheetName, key);
    if (column) {
      sheet.getRange(2, column, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  (dateTimeKeys || []).forEach((key) => {
    const column = getColumnIndex_(sheetName, key);
    if (column) {
      sheet.getRange(2, column, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
  });
  (percentKeys || []).forEach((key) => {
    const column = getColumnIndex_(sheetName, key);
    if (column) {
      sheet.getRange(2, column, lastRow - 1, 1).setNumberFormat('0.0%');
    }
  });
}
