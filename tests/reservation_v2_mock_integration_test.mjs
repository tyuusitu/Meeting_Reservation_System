#!/usr/bin/env node

import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import vm from 'node:vm';

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

class MockRange {
  constructor(sheet, row, column, rowCount, columnCount) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  getValues() {
    return Array.from({ length: this.rowCount }, (_, rowOffset) =>
      Array.from({ length: this.columnCount }, (_, columnOffset) =>
        this.sheet.rows[this.row - 1 + rowOffset]?.[this.column - 1 + columnOffset] ?? ''));
  }

  setValues(values) {
    values.forEach((sourceRow, rowOffset) => {
      const targetIndex = this.row - 1 + rowOffset;
      while (this.sheet.rows.length <= targetIndex) this.sheet.rows.push([]);
      sourceRow.forEach((value, columnOffset) => {
        this.sheet.rows[targetIndex][this.column - 1 + columnOffset] = value;
      });
    });
    return this;
  }

  setValue(value) {
    return this.setValues([[value]]);
  }

  setFontWeight() { return this; }
  setBackground() { return this; }
  setNumberFormat() { return this; }
}

class MockSheet {
  constructor(name, rows) {
    this.name = name;
    this.rows = clone(rows);
  }

  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map((row) => row.length)); }
  getMaxRows() { return Math.max(100, this.rows.length); }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  getDataRange() {
    if (this.name === '予約' && lockHeld && forbidReservationReadWhileLocked) {
      throw new Error('予約シート全件読込をロック中に実行してはいけません。');
    }
    return new MockRange(this, 1, 1, this.getLastRow(), this.getLastColumn());
  }
  getRange(row, column, rowCount = 1, columnCount = 1) {
    return new MockRange(this, row, column, rowCount, columnCount);
  }
  appendRow(row) { this.rows.push(row.slice()); }
  insertColumnsAfter() {}
  setFrozenRows() {}
  autoResizeColumns() {}
}

class MockSpreadsheet {
  constructor(sheets) {
    this.sheets = new Map(sheets.map((sheet) => [sheet.getName(), sheet]));
  }

  getId() { return 'TEST-SPREADSHEET-ID'; }
  getName() { return '[TEST] mock integration'; }
  getSheetByName(name) { return this.sheets.get(name) || null; }
}

class MockEvent {
  constructor(id, calendarId, eventStore) {
    this.id = id;
    this.calendarId = calendarId;
    this.eventStore = eventStore;
    this.deleted = false;
  }

  getId() { return this.id; }
  deleteEvent() {
    this.deleted = true;
    this.eventStore.delete(this.id);
  }
}

let uuidCounter = 0;
let eventCounter = 0;
let lockHeld = false;
let forbidReservationReadWhileLocked = true;
const cache = new Map();
const properties = new Map([
  ['SPREADSHEET_ID', 'TEST-SPREADSHEET-ID'],
  ['RESERVATION_V2_MODE', 'V2'],
  ['RESERVATION_V2_ENABLED', 'TRUE'],
  ['RESERVATION_V2_INDEX_READY', 'TRUE'],
]);
const eventStores = new Map([
  ['calendar-a', new Map()],
  ['calendar-b', new Map()],
  ['calendar-all', new Map()],
]);
const failingCalendars = new Set();
let spreadsheet;

const sandbox = {
  console,
  Utilities: {
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    Charset: { UTF_8: 'UTF_8' },
    getUuid() {
      uuidCounter += 1;
      return `00000000-0000-4000-8000-${String(uuidCounter).padStart(12, '0')}`;
    },
    computeDigest(_algorithm, value) {
      return [...crypto.createHash('sha256').update(value, 'utf8').digest()]
        .map((byte) => (byte > 127 ? byte - 256 : byte));
    },
    formatDate(_date, _timeZone, format) {
      if (format === 'yyyyMMdd') return '20260930';
      if (format === 'yyyy-MM-dd') return '2026-09-30';
      return '2026-09-30T12:34:56';
    },
  },
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: (key) => properties.get(key) || null,
      setProperty: (key, value) => properties.set(key, String(value)),
      deleteProperty: (key) => properties.delete(key),
      getProperties: () => Object.fromEntries(properties),
      setProperties: (values) => Object.entries(values).forEach(([key, value]) => properties.set(key, String(value))),
    }),
  },
  CacheService: {
    getScriptCache: () => ({
      get: (key) => cache.get(key) || null,
      put: (key, value) => cache.set(key, value),
      remove: (key) => cache.delete(key),
    }),
  },
  LockService: {
    getScriptLock: () => ({
      tryLock: () => {
        if (lockHeld) return false;
        lockHeld = true;
        return true;
      },
      releaseLock: () => { lockHeld = false; },
    }),
  },
  SpreadsheetApp: {
    openById: () => spreadsheet,
    getActiveSpreadsheet: () => spreadsheet,
    flush: () => {},
  },
  CalendarApp: {
    getCalendarById(calendarId) {
      if (!eventStores.has(calendarId)) return null;
      return {
        createEvent() {
          if (failingCalendars.has(calendarId)) throw new Error(`mock calendar failure: ${calendarId}`);
          eventCounter += 1;
          const id = `event-${eventCounter}`;
          const event = new MockEvent(id, calendarId, eventStores.get(calendarId));
          eventStores.get(calendarId).set(id, event);
          return event;
        },
      };
    },
  },
};

const context = vm.createContext(sandbox);
for (const file of ['gas/Setup.gs', 'gas/Api.gs', 'gas/ReservationV2.gs']) {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
}
const headers = vm.runInContext('JSON.parse(JSON.stringify(SHEET_HEADERS))', context);
const settingKeys = vm.runInContext('JSON.parse(JSON.stringify(SETTING_KEYS))', context);

spreadsheet = new MockSpreadsheet([
  new MockSheet('予約', [headers.reservations]),
  new MockSheet('会議室', [
    headers.rooms,
    ['TEST-ROOM-A', '会議室A', 'calendar-a', '有効', 1, '単体', ''],
    ['TEST-ROOM-B', '会議室B', 'calendar-b', '有効', 2, '単体', ''],
    ['TEST-ROOM-ALL', '大会議室', 'calendar-all', '有効', 3, '大会議室', 'TEST-ROOM-A,TEST-ROOM-B'],
  ]),
  new MockSheet('設定', [
    headers.settings,
    [settingKeys.BUSINESS_START_TIME, '09:00', '', ''],
    [settingKeys.BUSINESS_END_TIME, '22:00', '', ''],
    [settingKeys.TIME_SLOT_MINUTES, '5', '', ''],
    [settingKeys.MAX_RESERVATIONS_PER_SUBMIT, '10', '', ''],
    [settingKeys.FIELD_ORGANIZATION, '有効', '', ''],
    [settingKeys.FIELD_USER_NAME, '有効', '', ''],
  ]),
  new MockSheet('操作ログ', [headers.logs]),
]);

function submit(requestId, roomId, startTime, endTime) {
  sandbox.payload = {
    request_id: requestId,
    common: { organization_name: 'テスト団体', user_name: 'テスト担当者' },
    reservations: [{
      meeting_name: `会議 ${requestId}`,
      room_id: roomId,
      usage_date: '2027-02-10',
      start_time: startTime,
      end_time: endTime,
    }],
  };
  return vm.runInContext('insertReservationsV2_(payload)', context);
}

properties.set('RESERVATION_V2_INDEX_READY', 'FALSE');
const maintenance = submit('request-maintenance', 'TEST-ROOM-A', '09:00', '09:30');
assert.equal(maintenance.ok, false);
assert.equal(maintenance.state, 'maintenance');
assert.equal(maintenance.code, 'RESERVATION_V2_NOT_READY');
properties.set('RESERVATION_V2_INDEX_READY', 'TRUE');

const first = submit('request-first', 'TEST-ROOM-A', '10:00', '10:30');
assert.equal(first.ok, true);
assert.equal(first.state, 'active');
assert.equal(eventStores.get('calendar-a').size, 1);
assert.equal(spreadsheet.getSheetByName('予約').rows[1][12], '有効');

const duplicate = submit('request-first', 'TEST-ROOM-A', '10:00', '10:30');
assert.equal(duplicate.ok, true);
assert.equal(duplicate.state, 'active');
assert.equal(duplicate.duplicate, true);
assert.equal(eventStores.get('calendar-a').size, 1);
assert.equal(spreadsheet.getSheetByName('予約').rows.length, 2);

const conflict = submit('request-conflict', 'TEST-ROOM-ALL', '10:15', '10:45');
assert.equal(conflict.ok, false);
assert.equal(conflict.state, 'conflict');
assert.equal(spreadsheet.getSheetByName('予約').rows.length, 2);
assert.equal(eventStores.get('calendar-all').size, 0);

const independent = submit('request-independent', 'TEST-ROOM-B', '10:00', '10:30');
assert.equal(independent.ok, true);
assert.equal(independent.state, 'active');
assert.equal(eventStores.get('calendar-b').size, 1);

failingCalendars.add('calendar-b');
const failed = submit('request-failure', 'TEST-ROOM-ALL', '11:00', '11:30');
assert.equal(failed.ok, false);
assert.equal(failed.state, 'failed');
assert.equal(failed.code, 'CALENDAR_OR_FINALIZE_FAILED');
assert.equal(eventStores.get('calendar-all').size, 0);
assert.equal(eventStores.get('calendar-a').size, 1);
assert.equal(eventStores.get('calendar-b').size, 1);
const failedRow = spreadsheet.getSheetByName('予約').rows.at(-1);
assert.equal(failedRow[12], '登録失敗');
assert.equal(failedRow[11], '');
assert.match(failedRow[18], /mock calendar failure/);

sandbox.cancelledRow = vm.runInContext("selectSheetObjects_(SHEET_NAMES.reservations).map(normalizeReservationRow_).find((row) => row.request_id === 'request-first')", context);
sandbox.cancelledRow.status = '取消';
spreadsheet.getSheetByName('予約').rows[1][12] = '取消';
assert.equal(vm.runInContext('synchronizeCancelledReservationsV2_([cancelledRow])', context), true);
const cancelledRequest = submit('request-first', 'TEST-ROOM-A', '10:00', '10:30');
assert.equal(cancelledRequest.ok, true);
assert.equal(cancelledRequest.state, 'cancelled');
const reusedSlot = submit('request-after-cancel', 'TEST-ROOM-A', '10:00', '10:30');
assert.equal(reusedSlot.ok, true);
assert.equal(reusedSlot.state, 'active');

forbidReservationReadWhileLocked = false;
const disabled = vm.runInContext('disableReservationV2Release()', context);
assert.equal(disabled.mode, 'LEGACY');
assert.equal(disabled.enabled, false);
assert.equal(disabled.index_ready, false);
sandbox.payload = {
  request_id: 'stale-v2-client',
  common: { organization_name: 'テスト団体', user_name: 'テスト担当者' },
  reservations: [{
    meeting_name: '旧画面切替テスト',
    room_id: 'TEST-ROOM-A',
    usage_date: '2027-02-10',
    start_time: '12:00',
    end_time: '12:30',
  }],
};
const staleClient = vm.runInContext('handleInsertReservations_(payload)', context);
assert.equal(staleClient.ok, false);
assert.equal(staleClient.state, 'mode_changed');
const enabled = vm.runInContext('enableReservationV2Release()', context);
assert.equal(enabled.mode, 'V2');
assert.equal(enabled.enabled, true);
assert.equal(enabled.index_ready, true);
const afterEnable = submit('request-after-enable', 'TEST-ROOM-A', '12:00', '12:30');
assert.equal(afterEnable.ok, true);
assert.equal(afterEnable.state, 'active');

const corruptedSlotKey = vm.runInContext("getReservationSlotPropertyKeyV2_('2027-02-10', 'TEST-ROOM-A')", context);
properties.set(corruptedSlotKey, '{invalid-json');
assert.throws(
  () => submit('request-corrupt-index', 'TEST-ROOM-A', '13:00', '13:30'),
  /予約枠索引が破損しています/
);
assert.equal(properties.get('RESERVATION_V2_INDEX_READY'), 'FALSE');
const repaired = vm.runInContext('rebuildReservationV2Indexes()', context);
assert.equal(repaired.index_ready, true);

process.stdout.write('reservation_v2_mock_integration_test: 44 assertions passed\n');
