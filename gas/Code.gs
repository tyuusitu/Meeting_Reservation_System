/** スプレッドシートのシート名定義 */
const SHEET_NAMES = {
  reservations: '予約',
  rooms: '会議室',
  settings: '設定',
  logs: 'ログ',
};

/** 旧シート名から日本語シート名へ移行するための定義 */
const LEGACY_SHEET_NAMES = {
  reservations: 'reservations',
  rooms: 'rooms',
  settings: 'settings',
  logs: 'logs',
};

/** 各シートのヘッダー定義 */
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
  ],
  logs: [
    '記録日時',
    'レベル',
    '操作',
    'メッセージ',
    '詳細',
  ],
};

/** スプレッドシート列をコード内キーへ変換する定義 */
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
  ],
  logs: [
    'logged_at',
    'level',
    'action',
    'message',
    'payload',
  ],
};

/** システムで使う既定設定 */
const DEFAULT_SETTINGS = {
  SYSTEM_NAME: '会議室予約システム',
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
  ADMIN_PASSWORD: 'admin1234',
  LINE_NOTIFY_URL: '',
  SPREADSHEET_URL: '',
};

/** 設定キーのスプレッドシート表示名 */
const SETTING_LABELS = {
  SYSTEM_NAME: 'システム名',
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
  ADMIN_PASSWORD: '管理者パスワード',
  LINE_NOTIFY_URL: 'LINE通知URL',
  SPREADSHEET_URL: 'スプレッドシートURL',
};

/** 予約ステータスの有効値 */
const RESERVATION_STATUS = {
  active: '有効',
  cancelled: '取消',
};

/** スクリプトプロパティで使うスプレッドシートIDのキー */
const SPREADSHEET_ID_PROPERTY_KEY = 'SPREADSHEET_ID';

/**
 * 複数予約を確定し、Googleカレンダーとスプレッドシートに登録します。
 *
 * @param {Object} payload 登録する予約情報。
 * @return {Object} 登録結果。
 */
function insertReservations(payload) {
  /** 同時登録を防ぐためのロック */
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の予約処理が実行中です。少し待ってからもう一度お試しください。');
  }

  /** 登録済みカレンダーイベント */
  const createdEvents = [];
  try {
    /** 登録直前の検証結果 */
    const validation = validateReservationsInternal_(payload);
    if (!validation.ok) {
      return validation;
    }

    /** 共通入力情報 */
    const common = validation.common;
    /** 有効な会議室マップ */
    const roomMap = createRoomMap_(selectActiveRooms_());
    /** 作成日時 */
    const createdAt = formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss");
    /** 保存する予約行 */
    const reservationRows = [];
    /** 登録結果として返す予約一覧 */
    const registeredReservations = [];

    validation.reservations.forEach((reservation) => {
      /** 対象会議室 */
      const room = roomMap[reservation.room_id];
      if (!room || !room.calendar_id) {
        throw new Error(`会議室「${room ? room.room_name : reservation.room_id}」のカレンダーIDが未設定です。会議室シートを確認してください。`);
      }

      /** 予約ID */
      const reservationId = createReservationId_();
      /** 開始日時 */
      const startDate = createDateTime_(reservation.usage_date, reservation.start_time);
      /** 終了日時 */
      const endDate = createDateTime_(reservation.usage_date, reservation.end_time);
      /** 登録先カレンダー */
      const calendar = CalendarApp.getCalendarById(room.calendar_id);
      if (!calendar) {
        throw new Error(`会議室「${room.room_name}」のGoogleカレンダーが見つかりません。カレンダーIDを確認してください。`);
      }

      /** カレンダー予定タイトル */
      const eventTitle = common.organization_name
        ? `【${common.organization_name}】${reservation.meeting_name}`
        : reservation.meeting_name;
      /** カレンダー予定説明 */
      const eventDescription = [
        common.organization_name ? `団体名：${common.organization_name}` : null,
        `会議名：${reservation.meeting_name}`,
        common.user_name ? `お名前：${common.user_name}` : null,
        `会議室：${room.room_name}`,
        `予約ID：${reservationId}`,
      ].filter((line) => line !== null).join('\n');
      /** 作成されたカレンダー予定 */
      const calendarEvent = calendar.createEvent(eventTitle, startDate, endDate, {
        description: eventDescription,
      });
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
    writeLog_('INFO', 'insertReservations', '予約を登録しました。', {
      count: registeredReservations.length,
    });

    return {
      ok: true,
      errors: [],
      reservations: registeredReservations,
    };
  } catch (error) {
    createdEvents.forEach((calendarEvent) => {
      try {
        calendarEvent.deleteEvent();
      } catch (deleteError) {
        writeLog_('ERROR', 'rollbackCalendarEvent', deleteError.message, {});
      }
    });
    writeLog_('ERROR', 'insertReservations', error.message, {
      stack: error.stack || '',
    });
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 予約登録前の検証処理を実行します。
 *
 * @param {Object} payload 登録予定の予約情報。
 * @return {Object} 検証結果。
 */
function validateReservationsInternal_(payload) {
  /** 入力全体 */
  const requestPayload = payload || {};
  /** 共通入力情報 */
  const common = normalizeCommonInput_(requestPayload.common || {});
  /** 予約入力一覧 */
  const reservations = Array.isArray(requestPayload.reservations)
    ? requestPayload.reservations.map((reservation) => normalizeReservationInput_(reservation))
    : [];
  /** 検証エラー一覧 */
  const errors = [];
  /** システム設定 */
  const settings = getSettings_();
  /** 1回で登録できる最大予約数 */
  const maxReservationCount = Number(settings.MAX_RESERVATIONS_PER_SUBMIT || DEFAULT_SETTINGS.MAX_RESERVATIONS_PER_SUBMIT);
  /** フィールド表示設定 */
  const fieldConfig = {
    organizationName: normalizeBoolean_(settings.FIELD_ORGANIZATION || DEFAULT_SETTINGS.FIELD_ORGANIZATION),
    userName: normalizeBoolean_(settings.FIELD_USER_NAME || DEFAULT_SETTINGS.FIELD_USER_NAME),
  };

  validateCommonInput_(common, fieldConfig).forEach((error) => errors.push(error));

  if (reservations.length === 0) {
    errors.push({
      index: null,
      field: 'reservations',
      message: '予約を1件以上入力してください。',
    });
  }

  if (reservations.length > maxReservationCount) {
    errors.push({
      index: null,
      field: 'reservations',
      message: `一度に登録できる予約は${maxReservationCount}件までです。`,
    });
  }

  /** 有効な会議室一覧 */
  const rooms = selectActiveRooms_();
  /** 有効な会議室マップ */
  const roomMap = createRoomMap_(rooms);
  /** 登録済み有効予約 */
  const existingReservations = selectActiveReservations_();

  reservations.forEach((reservation, index) => {
    validateSingleReservationInput_(reservation, index, roomMap, settings).forEach((error) => errors.push(error));
  });

  reservations.forEach((reservation, index) => {
    if (!reservation.usage_date || !reservation.start_time || !reservation.end_time || !reservation.room_id) {
      return;
    }

    /** 既存予約との重複 */
    const duplicatedExistingReservation = existingReservations.find((existingReservation) => {
      return existingReservation.status === RESERVATION_STATUS.active
        && existingReservation.room_id === reservation.room_id
        && existingReservation.usage_date === reservation.usage_date
        && hasTimeOverlap_(existingReservation.start_time, existingReservation.end_time, reservation.start_time, reservation.end_time);
    });
    if (duplicatedExistingReservation) {
      errors.push({
        index,
        field: 'room_id',
        message: `${duplicatedExistingReservation.room_name}は${reservation.start_time}-${reservation.end_time}に既存予約があります。`,
      });
    }

    reservations.forEach((otherReservation, otherIndex) => {
      if (index >= otherIndex) {
        return;
      }
      if (
        reservation.room_id
        && otherReservation.room_id
        && reservation.room_id === otherReservation.room_id
        && reservation.usage_date === otherReservation.usage_date
        && hasTimeOverlap_(reservation.start_time, reservation.end_time, otherReservation.start_time, otherReservation.end_time)
      ) {
        errors.push({
          index,
          field: 'room_id',
          message: `${index + 1}件目と${otherIndex + 1}件目の予約時間が同じ会議室で重複しています。`,
        });
        errors.push({
          index: otherIndex,
          field: 'room_id',
          message: `${index + 1}件目と${otherIndex + 1}件目の予約時間が同じ会議室で重複しています。`,
        });
      }
    });
  });

  if (errors.length > 0) {
    return {
      ok: false,
      errors,
      common,
      reservations,
    };
  }

  return {
    ok: true,
    errors: [],
    common,
    reservations,
  };
}

/**
 * 共通入力を正規化します。
 *
 * @param {Object} commonInput 共通入力。
 * @return {Object} 正規化済み共通入力。
 */
function normalizeCommonInput_(commonInput) {
  return {
    line_user_id: normalizeString_(commonInput.line_user_id || commonInput.lineUserId),
    organization_name: normalizeString_(commonInput.organization_name || commonInput.organizationName),
    user_name: normalizeString_(commonInput.user_name || commonInput.userName),
  };
}

/**
 * 予約入力を正規化します。
 *
 * @param {Object} reservationInput 予約入力。
 * @return {Object} 正規化済み予約入力。
 */
function normalizeReservationInput_(reservationInput) {
  /** 予約入力 */
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
 * 共通入力の必須項目を検証します。
 *
 * @param {Object} common 共通入力。
 * @return {Array<Object>} エラー一覧。
 */
function validateCommonInput_(common, fieldConfig) {
  /** 検証エラー一覧 */
  const errors = [];
  /** フィールド設定（未指定はすべて有効） */
  const config = fieldConfig || { organizationName: true, userName: true };

  if (config.organizationName && !common.organization_name) {
    errors.push({
      index: null,
      field: 'organization_name',
      message: '団体名を入力してください。',
    });
  }

  if (config.userName && !common.user_name) {
    errors.push({
      index: null,
      field: 'user_name',
      message: 'お名前を入力してください。',
    });
  }

  return errors;
}

/**
 * 1件分の予約入力を検証します。
 *
 * @param {Object} reservation 予約入力。
 * @param {number} index 予約の位置。
 * @param {Object} roomMap 会議室マップ。
 * @return {Array<Object>} エラー一覧。
 */
function validateSingleReservationInput_(reservation, index, roomMap, settings) {
  /** 検証エラー一覧 */
  const errors = [];

  if (!reservation.meeting_name) {
    errors.push({
      index,
      field: 'meeting_name',
      message: '会議名を入力してください。',
    });
  }

  validateDateTimeInput_(reservation.usage_date, reservation.start_time, reservation.end_time, settings).forEach((message) => {
    errors.push({
      index,
      field: 'date_time',
      message,
    });
  });

  if (!reservation.room_id) {
    errors.push({
      index,
      field: 'room_id',
      message: '会議室を選択してください。',
    });
  } else if (!roomMap[reservation.room_id]) {
    errors.push({
      index,
      field: 'room_id',
      message: '選択された会議室は利用できません。',
    });
  } else if (!roomMap[reservation.room_id].calendar_id) {
    errors.push({
      index,
      field: 'room_id',
      message: `会議室「${roomMap[reservation.room_id].room_name}」のカレンダーIDが未設定です。`,
    });
  }

  return errors;
}

/**
 * 日付と時間の入力を検証します。
 *
 * @param {string} usageDate 利用日。
 * @param {string} startTime 開始時間。
 * @param {string} endTime 終了時間。
 * @return {Array<string>} エラーメッセージ一覧。
 */
function validateDateTimeInput_(usageDate, startTime, endTime, settings) {
  /** エラーメッセージ一覧 */
  const errors = [];

  if (!usageDate) {
    errors.push('利用日を入力してください。');
  } else if (!/^\d{4}-\d{2}-\d{2}$/.test(usageDate)) {
    errors.push('利用日の形式を確認してください。');
  }

  if (!startTime) {
    errors.push('開始時間を入力してください。');
  } else if (!/^\d{2}:\d{2}$/.test(startTime)) {
    errors.push('開始時間の形式を確認してください。');
  }

  if (!endTime) {
    errors.push('終了時間を入力してください。');
  } else if (!/^\d{2}:\d{2}$/.test(endTime)) {
    errors.push('終了時間の形式を確認してください。');
  }

  if (startTime && endTime && toMinutes_(startTime) >= toMinutes_(endTime)) {
    errors.push('終了時間は開始時間より後にしてください。');
  }

  if (settings && startTime && endTime) {
    /** 5分刻み */
    const slotMinutes = Number(settings.TIME_SLOT_MINUTES || DEFAULT_SETTINGS.TIME_SLOT_MINUTES || 5);
    /** 営業開始分 */
    const businessStartMin = toMinutes_(settings.BUSINESS_START_TIME || DEFAULT_SETTINGS.BUSINESS_START_TIME || '09:00');
    /** 営業終了分 */
    const businessEndMin = toMinutes_(settings.BUSINESS_END_TIME || DEFAULT_SETTINGS.BUSINESS_END_TIME || '22:00');
    /** 開始分 */
    const startMin = toMinutes_(startTime);
    /** 終了分 */
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
 * 有効な予約一覧を取得します。
 *
 * @return {Array<Object>} 有効予約一覧。
 */
function selectActiveReservations_() {
  /** 予約シートの全行 */
  const rows = selectSheetObjects_(SHEET_NAMES.reservations);

  return rows
    .map((row) => normalizeReservationRow_(row))
    .filter((reservation) => reservation.reservation_id && reservation.status === RESERVATION_STATUS.active);
}

/**
 * 予約行を画面用に正規化します。
 *
 * @param {Object} row シートから取得した予約行。
 * @return {Object} 正規化済み予約行。
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
 * 有効な会議室一覧を取得します。
 *
 * @return {Array<Object>} 会議室一覧。
 */
function selectActiveRooms_() {
  /** 会議室シートの全行 */
  const rows = selectSheetObjects_(SHEET_NAMES.rooms);

  return rows
    .map((row) => {
      /** カレンダーID */
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
    .sort((leftRoom, rightRoom) => {
      return Number(leftRoom.display_order) - Number(rightRoom.display_order)
        || compareValues_(leftRoom.room_name, rightRoom.room_name);
    });
}

/**
 * 会議室IDをキーにしたマップを作成します。
 *
 * @param {Array<Object>} rooms 会議室一覧。
 * @return {Object} 会議室マップ。
 */
function createRoomMap_(rooms) {
  /** 会議室マップ */
  const roomMap = {};
  rooms.forEach((room) => {
    roomMap[room.room_id] = room;
  });
  return roomMap;
}

/**
 * シートの行をオブジェクト配列として取得します。
 *
 * @param {string} sheetName シート名。
 * @return {Array<Object>} 行オブジェクト一覧。
 */
function selectSheetObjects_(sheetName) {
  /** 対象シート */
  const sheet = getSheet_(sheetName);
  /** スキーマキー */
  const schemaKey = getSchemaKeyBySheetName_(sheetName);
  /** 最終行番号 */
  const lastRow = sheet.getLastRow();
  /** 最終列番号 */
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn === 0) {
    return [];
  }

  /** ヘッダー行 */
  const headers = getHeaderRow_(sheet);
  /** コード内で使う列キー */
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);
  /** データ行 */
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values
    .filter((row) => row.some((cell) => normalizeString_(cell) !== ''))
    .map((row) => {
      /** 行オブジェクト */
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
 * @param {string} sheetName シート名。
 * @return {string} スキーマキー。
 */
function getSchemaKeyBySheetName_(sheetName) {
  /** 正規化済みシート名 */
  const normalizedSheetName = normalizeString_(sheetName);
  /** 現行シート名に対応するキー */
  const currentKey = Object.keys(SHEET_NAMES).find((key) => SHEET_NAMES[key] === normalizedSheetName);
  if (currentKey) {
    return currentKey;
  }

  /** 旧シート名に対応するキー */
  const legacyKey = Object.keys(LEGACY_SHEET_NAMES).find((key) => LEGACY_SHEET_NAMES[key] === normalizedSheetName);
  return legacyKey || normalizedSheetName;
}

/**
 * ヘッダー名からコード内の列キー一覧を取得します。
 *
 * @param {string} schemaKey スキーマキー。
 * @param {Array<string>} headers ヘッダー一覧。
 * @return {Array<string>} 列キー一覧。
 */
function getColumnKeysForHeaders_(schemaKey, headers) {
  /** 日本語ヘッダー一覧 */
  const sheetHeaders = SHEET_HEADERS[schemaKey] || [];
  /** コード内の列キー一覧 */
  const columnKeys = SHEET_COLUMN_KEYS[schemaKey] || [];

  return headers.map((header) => {
    /** 日本語ヘッダー位置 */
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
 * 指定列キーの列番号を取得します。
 *
 * @param {string} sheetName シート名。
 * @param {string} columnKey 列キー。
 * @return {number} 1始まりの列番号。見つからない場合は0。
 */
function getColumnIndex_(sheetName, columnKey) {
  /** 対象シート */
  const sheet = getSheet_(sheetName);
  /** スキーマキー */
  const schemaKey = getSchemaKeyBySheetName_(sheetName);
  /** ヘッダー行 */
  const headers = getHeaderRow_(sheet);
  /** 列キー一覧 */
  const columnKeys = getColumnKeysForHeaders_(schemaKey, headers);
  return columnKeys.indexOf(columnKey) + 1;
}

/**
 * 指定シートへ複数行を追加します。
 *
 * @param {string} sheetName シート名。
 * @param {Array<Array<*>>} rows 追加する行。
 */
function appendSheetRows_(sheetName, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  /** 対象シート */
  const sheet = getSheet_(sheetName);
  /** 追加開始行 */
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * 指定シートを取得します。
 *
 * @param {string} sheetName シート名。
 * @return {Sheet} 対象シート。
 */
function getSheet_(sheetName) {
  /** スプレッドシート */
  const spreadsheet = getSpreadsheet_();
  /** 対象シート */
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    /** 旧シート名 */
    const schemaKey = getSchemaKeyBySheetName_(sheetName);
    const legacySheetName = LEGACY_SHEET_NAMES[schemaKey];
    sheet = legacySheetName ? spreadsheet.getSheetByName(legacySheetName) : null;
  }
  if (!sheet) {
    throw new Error(`${sheetName}シートが見つかりません。initializeSpreadsheet()を実行してください。`);
  }
  return sheet;
}

/**
 * 使用するスプレッドシートを取得します。
 *
 * @return {Spreadsheet} スプレッドシート。
 */
function getSpreadsheet_() {
  /** スクリプトプロパティのスプレッドシートID */
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY_KEY);
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  /** コンテナバインド時のアクティブスプレッドシート */
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  throw new Error(`スプレッドシートが見つかりません。スクリプトプロパティ${SPREADSHEET_ID_PROPERTY_KEY}を設定してください。`);
}

/**
 * シートのヘッダー行を取得します。
 *
 * @param {Sheet} sheet 対象シート。
 * @return {Array<string>} ヘッダー名一覧。
 */
function getHeaderRow_(sheet) {
  /** 最終列番号 */
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return [];
  }

  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map((header) => normalizeString_(header));
}

/**
 * 設定シートから設定値を取得します。
 *
 * @return {Object} 設定値。
 */
function getSettings_() {
  /** 既定値で初期化した設定 */
  const settings = {};
  Object.keys(DEFAULT_SETTINGS).forEach((key) => {
    settings[key] = DEFAULT_SETTINGS[key];
  });

  try {
    /** 設定シートの行一覧 */
    const rows = selectSheetObjects_(SHEET_NAMES.settings);
    rows.forEach((row) => {
      /** 設定キー */
      const key = getSettingInternalKey_(row.setting_key);
      if (key) {
        settings[key] = normalizeString_(row.setting_value);
      }
    });
  } catch (error) {
    return settings;
  }

  return settings;
}

/**
 * 設定シート上のキーをコード内キーへ変換します。
 *
 * @param {*} settingKey 設定キー。
 * @return {string} コード内設定キー。
 */
function getSettingInternalKey_(settingKey) {
  /** 正規化済み設定キー */
  const normalizedSettingKey = normalizeString_(settingKey);
  if (!normalizedSettingKey) {
    return '';
  }
  if (DEFAULT_SETTINGS[normalizedSettingKey] !== undefined) {
    return normalizedSettingKey;
  }

  /** 日本語設定キーに対応するコード内キー */
  const matchedKey = Object.keys(SETTING_LABELS).find((key) => SETTING_LABELS[key] === normalizedSettingKey);
  return matchedKey || '';
}

/**
 * 操作ログをログシートへ書き込みます。
 *
 * @param {string} level ログレベル。
 * @param {string} action 操作名。
 * @param {string} message メッセージ。
 * @param {Object} payload 追加情報。
 */
function writeLog_(level, action, message, payload) {
  try {
    /** ログシート */
    const sheet = getSheet_(SHEET_NAMES.logs);
    sheet.appendRow([
      formatDate_(new Date(), "yyyy-MM-dd'T'HH:mm:ss"),
      level,
      action,
      message,
      JSON.stringify(payload || {}),
    ]);
  } catch (error) {
    console.error(error);
  }
}

/**
 * Googleカレンダー確認用URLを作成します。
 *
 * @param {string} calendarId カレンダーID。
 * @return {string} GoogleカレンダーURL。
 */
function createCalendarUrl_(calendarId) {
  /** 正規化済みカレンダーID */
  const normalizedCalendarId = normalizeString_(calendarId);
  if (!normalizedCalendarId) {
    return '';
  }
  return `https://calendar.google.com/calendar/embed?src=${encodeURIComponent(normalizedCalendarId)}`;
}

/**
 * 予約IDを生成します。
 *
 * @return {string} 予約ID。
 */
function createReservationId_() {
  /** 現在日時 */
  const now = new Date();
  /** 日時部分 */
  const timestamp = formatDate_(now, 'yyyyMMddHHmmss');
  /** ランダム部分 */
  const randomText = Math.random().toString(36).slice(2, 8).toUpperCase();
  return `R-${timestamp}-${randomText}`;
}

/**
 * 日付と時刻の文字列からDateを作成します。
 *
 * @param {string} usageDate 利用日。
 * @param {string} time 時刻。
 * @return {Date} Dateオブジェクト。
 */
function createDateTime_(usageDate, time) {
  /** 日付の分解値 */
  const dateParts = usageDate.split('-').map((value) => Number(value));
  /** 時刻の分解値 */
  const timeParts = time.split(':').map((value) => Number(value));
  return new Date(dateParts[0], dateParts[1] - 1, dateParts[2], timeParts[0], timeParts[1], 0);
}

/**
 * 2つの時間帯が重複しているか判定します。
 *
 * @param {string} existingStart 既存予約の開始時間。
 * @param {string} existingEnd 既存予約の終了時間。
 * @param {string} newStart 新規予約の開始時間。
 * @param {string} newEnd 新規予約の終了時間。
 * @return {boolean} 重複している場合はtrue。
 */
function hasTimeOverlap_(existingStart, existingEnd, newStart, newEnd) {
  return toMinutes_(existingStart) < toMinutes_(newEnd)
    && toMinutes_(newStart) < toMinutes_(existingEnd);
}

/**
 * HH:mmを分に変換します。
 *
 * @param {string} time 時刻文字列。
 * @return {number} 0:00からの分数。
 */
function toMinutes_(time) {
  /** 時刻文字列 */
  const normalizedTime = normalizeTimeString_(time);
  /** 時刻の分解値 */
  const parts = normalizedTime.split(':').map((value) => Number(value));
  if (parts.length !== 2 || Number.isNaN(parts[0]) || Number.isNaN(parts[1])) {
    return 0;
  }
  return parts[0] * 60 + parts[1];
}

/**
 * 値を文字列として正規化します。
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
 * 日付文字列をyyyy-MM-ddへ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 日付文字列。
 */
function normalizeDateString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, 'yyyy-MM-dd');
  }

  /** 文字列化した値 */
  const stringValue = normalizeString_(value);
  if (!stringValue) {
    return '';
  }

  /** yyyy-MM-dd形式の値 */
  const dateMatch = stringValue.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (dateMatch) {
    return [
      dateMatch[1],
      pad2_(dateMatch[2]),
      pad2_(dateMatch[3]),
    ].join('-');
  }

  return stringValue;
}

/**
 * 時刻文字列をHH:mmへ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 時刻文字列。
 */
function normalizeTimeString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, 'HH:mm');
  }

  /** 文字列化した値 */
  const stringValue = normalizeString_(value);
  if (!stringValue) {
    return '';
  }

  /** HH:mm形式の値 */
  const timeMatch = stringValue.match(/^(\d{1,2}):(\d{1,2})$/);
  if (timeMatch) {
    return `${pad2_(timeMatch[1])}:${pad2_(timeMatch[2])}`;
  }

  return stringValue;
}

/**
 * 日時文字列を保存用に正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 日時文字列。
 */
function normalizeDateTimeString_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDate_(value, "yyyy-MM-dd'T'HH:mm:ss");
  }
  return normalizeString_(value);
}

/**
 * 予約ステータスを保存用の日本語値へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {string} 予約ステータス。
 */
function normalizeReservationStatus_(value) {
  /** 文字列化した値 */
  const stringValue = normalizeString_(value).toLowerCase();
  if (!stringValue || stringValue === 'active' || stringValue === '有効') {
    return RESERVATION_STATUS.active;
  }
  if (stringValue === 'cancelled' || stringValue === 'canceled' || stringValue === '取消' || stringValue === 'キャンセル') {
    return RESERVATION_STATUS.cancelled;
  }
  return normalizeString_(value);
}

/**
 * 値を真偽値へ正規化します。
 *
 * @param {*} value 対象値。
 * @return {boolean} 真として扱える場合はtrue。
 */
function normalizeBoolean_(value) {
  /** 文字列化した値 */
  const stringValue = normalizeString_(value).toLowerCase();
  return value === true
    || stringValue === 'true'
    || stringValue === '1'
    || stringValue === 'yes'
    || stringValue === '有効'
    || stringValue === '利用可';
}

/**
 * 2桁文字列へ変換します。
 *
 * @param {*} value 対象値。
 * @return {string} 2桁文字列。
 */
function pad2_(value) {
  return String(value).padStart(2, '0');
}

/**
 * Dateを指定形式で文字列化します。
 *
 * @param {Date} date 対象日付。
 * @param {string} pattern フォーマット。
 * @return {string} 日付文字列。
 */
function formatDate_(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Tokyo', pattern);
}

/**
 * 値を比較します。
 *
 * @param {*} leftValue 左辺。
 * @param {*} rightValue 右辺。
 * @return {number} 比較結果。
 */
function compareValues_(leftValue, rightValue) {
  /** 左辺文字列 */
  const leftText = normalizeString_(leftValue);
  /** 右辺文字列 */
  const rightText = normalizeString_(rightValue);
  if (leftText < rightText) {
    return -1;
  }
  if (leftText > rightText) {
    return 1;
  }
  return 0;
}
