/** 初期投入する会議室データ */
const DEFAULT_ROOM_ROWS = [
  ['room-1', '中執内応接室', '', '有効', 1],
  ['room-2', '階段下会議室', '', '有効', 2],
  ['room-3', '中執前会議室', '', '有効', 3],
];

/** 設定シートへ初期投入する説明文 */
const DEFAULT_SETTING_DESCRIPTIONS = {
  SYSTEM_NAME: '画面上に表示するシステム名です。',
  BUSINESS_START_TIME: 'タイムライン表示で使う開始時刻です。HH:mm形式で入力します。',
  BUSINESS_END_TIME: 'タイムライン表示で使う終了時刻です。HH:mm形式で入力します。',
  TIME_SLOT_MINUTES: '入力フォームの時間刻みです。',
  MAX_RESERVATIONS_PER_SUBMIT: '一度にまとめて登録できる予約数の上限です。',
  LABEL_ORGANIZATION: '予約フォームの団体名欄のラベルです。例：組織名、学校名、会社名',
  LABEL_USER_NAME: '予約フォームのお名前欄のラベルです。例：担当者名、申請者名',
  LABEL_MEETING_NAME: '予約フォームの会議名欄のラベルです。例：イベント名、目的、用途',
  LABEL_ROOM: '会議室の呼び方です。例：スペース、部屋、施設',
  FIELD_ORGANIZATION: '予約フォームの団体名欄の表示です。有効または無効を入力します。',
  FIELD_USER_NAME: '予約フォームのお名前欄の表示です。有効または無効を入力します。',
  ADMIN_PASSWORD: '管理者画面のパスワードです。必ず変更してください。',
  LINE_NOTIFY_URL: 'キャンセル依頼時にLINEへ送るWebhook URL（任意）。',
  SPREADSHEET_URL: 'このスプレッドシートのURL。キャンセル依頼LINEメッセージに添付されます。',
};

/**
 * 必要なシート、ヘッダー、初期設定、初期会議室を作成します。
 *
 * @return {Object} 初期化結果。
 */
function initializeSpreadsheet() {
  /** 対象スプレッドシート */
  const spreadsheet = getSpreadsheet_();

  Object.keys(SHEET_NAMES).forEach((schemaKey) => {
    ensureSheetAndHeader_(spreadsheet, schemaKey);
  });
  renameUnusedLegacySheet_(spreadsheet, 'users', '旧過去入力', [
    '利用者ID',
    'メールアドレス',
    '団体名',
    'お名前',
    '更新日時',
  ]);

  ensureDefaultSettings_();
  ensureDefaultRooms_();
  writeLog_('INFO', 'initializeSpreadsheet', 'スプレッドシート初期化を実行しました。', {
    spreadsheet_id: spreadsheet.getId(),
  });

  return {
    ok: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    sheets: Object.keys(SHEET_NAMES).map((schemaKey) => SHEET_NAMES[schemaKey]),
  };
}

/**
 * 現在開いているスプレッドシートIDをスクリプトプロパティに保存します。
 *
 * @return {Object} 保存結果。
 */
function setSpreadsheetIdToCurrent() {
  /** アクティブスプレッドシート */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('アクティブなスプレッドシートが見つかりません。スプレッドシートからApps Scriptを開いて実行してください。');
  }

  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY_KEY, spreadsheet.getId());

  return {
    ok: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
  };
}

/**
 * 会議室シートのカレンダーIDが空の会議室にGoogleカレンダーを作成してIDを書き込みます。
 *
 * @return {Object} 作成結果。
 */
function createRoomCalendarsAndUpdateRooms() {
  /** 会議室シート */
  const sheet = getSheet_(SHEET_NAMES.rooms);
  /** 会議室名列 */
  const roomNameColumn = getColumnIndex_(SHEET_NAMES.rooms, 'room_name');
  /** カレンダーID列 */
  const calendarIdColumn = getColumnIndex_(SHEET_NAMES.rooms, 'calendar_id');
  /** 利用可否列 */
  const isActiveColumn = getColumnIndex_(SHEET_NAMES.rooms, 'is_active');
  /** 最終行 */
  const lastRow = sheet.getLastRow();
  /** 作成したカレンダー一覧 */
  const createdCalendars = [];

  if (!roomNameColumn || !calendarIdColumn || !isActiveColumn) {
    throw new Error('会議室シートのヘッダーを確認してください。');
  }

  if (lastRow < 2) {
    return {
      ok: true,
      createdCalendars,
      message: '会議室シートに会議室がありません。',
    };
  }

  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber += 1) {
    /** 会議室名 */
    const roomName = normalizeString_(sheet.getRange(rowNumber, roomNameColumn).getValue());
    /** 現在のカレンダーID */
    const calendarId = normalizeString_(sheet.getRange(rowNumber, calendarIdColumn).getValue());
    /** 利用可能か */
    const isActive = normalizeBoolean_(sheet.getRange(rowNumber, isActiveColumn).getValue());

    if (!roomName || calendarId || !isActive) {
      continue;
    }

    /** 作成するカレンダー名 */
    const calendarName = `${roomName} 予約`;
    /** 作成されたカレンダー */
    const calendar = CalendarApp.createCalendar(calendarName, {
      timeZone: Session.getScriptTimeZone() || 'Asia/Tokyo',
    });
    sheet.getRange(rowNumber, calendarIdColumn).setValue(calendar.getId());

    createdCalendars.push({
      room_name: roomName,
      calendar_name: calendarName,
      calendar_id: calendar.getId(),
    });
  }

  writeLog_('INFO', 'createRoomCalendarsAndUpdateRooms', '会議室カレンダー作成を実行しました。', {
    count: createdCalendars.length,
  });

  return {
    ok: true,
    createdCalendars,
  };
}

/**
 * 指定シートを作成し、ヘッダーを設定します。
 *
 * @param {Spreadsheet} spreadsheet 対象スプレッドシート。
 * @param {string} schemaKey スキーマキー。
 */
function ensureSheetAndHeader_(spreadsheet, schemaKey) {
  /** 現行シート名 */
  const sheetName = SHEET_NAMES[schemaKey];
  /** ヘッダー一覧 */
  const headers = SHEET_HEADERS[schemaKey];
  /** 対象シート */
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    /** 旧シート */
    const legacySheet = LEGACY_SHEET_NAMES[schemaKey]
      ? spreadsheet.getSheetByName(LEGACY_SHEET_NAMES[schemaKey])
      : null;
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
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#EAF2FF');
  sheet.setFrozenRows(1);

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).setNumberFormat('@');
  }

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * 現在使わない旧シートを日本語名に変更します。
 *
 * @param {Spreadsheet} spreadsheet 対象スプレッドシート。
 * @param {string} legacySheetName 旧シート名。
 * @param {string} newSheetName 新シート名。
 * @param {Array<string>} headers ヘッダー一覧。
 */
function renameUnusedLegacySheet_(spreadsheet, legacySheetName, newSheetName, headers) {
  /** 旧シート */
  const legacySheet = spreadsheet.getSheetByName(legacySheetName);
  if (!legacySheet || spreadsheet.getSheetByName(newSheetName)) {
    return;
  }

  legacySheet.setName(newSheetName);
  if (legacySheet.getMaxColumns() < headers.length) {
    legacySheet.insertColumnsAfter(legacySheet.getMaxColumns(), headers.length - legacySheet.getMaxColumns());
  }
  legacySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  legacySheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#EAF2FF');
  legacySheet.setFrozenRows(1);
  legacySheet.autoResizeColumns(1, headers.length);
}

/**
 * 設定シートへ未登録の初期設定を追加します。
 */
function ensureDefaultSettings_() {
  /** 設定シート */
  const sheet = getSheet_(SHEET_NAMES.settings);
  /** 設定キー列 */
  const settingKeyColumn = getColumnIndex_(SHEET_NAMES.settings, 'setting_key') || 1;
  /** 最終行 */
  const lastRow = sheet.getLastRow();
  /** 登録済み設定キー */
  const existingKeys = [];
  /** 追加する設定行 */
  const settingRows = [];

  if (lastRow >= 2) {
    /** 設定キーの値一覧 */
    const settingKeyValues = sheet.getRange(2, settingKeyColumn, lastRow - 1, 1).getValues();
    settingKeyValues.forEach((row, index) => {
      /** コード内設定キー */
      const internalKey = getSettingInternalKey_(row[0]);
      if (!internalKey) {
        return;
      }
      existingKeys.push(internalKey);
      if (SETTING_LABELS[internalKey] && normalizeString_(row[0]) !== SETTING_LABELS[internalKey]) {
        sheet.getRange(index + 2, settingKeyColumn).setValue(SETTING_LABELS[internalKey]);
      }
    });
  }

  Object.keys(DEFAULT_SETTINGS).forEach((key) => {
    if (existingKeys.indexOf(key) === -1) {
      settingRows.push([
        SETTING_LABELS[key] || key,
        DEFAULT_SETTINGS[key],
        DEFAULT_SETTING_DESCRIPTIONS[key] || '',
      ]);
    }
  });

  if (settingRows.length > 0) {
    appendSheetRows_(SHEET_NAMES.settings, settingRows);
  }
}

/**
 * 会議室シートが空の場合に初期会議室を追加します。
 */
function ensureDefaultRooms_() {
  /** 登録済み会議室 */
  const existingRooms = selectSheetObjects_(SHEET_NAMES.rooms)
    .filter((row) => normalizeString_(row.room_id));

  if (existingRooms.length > 0) {
    return;
  }

  appendSheetRows_(SHEET_NAMES.rooms, DEFAULT_ROOM_ROWS);
}
