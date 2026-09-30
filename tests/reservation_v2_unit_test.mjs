#!/usr/bin/env node

import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import vm from 'node:vm';

let uuidCounter = 0;
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
};
const context = vm.createContext(sandbox);
for (const file of ['gas/Setup.gs', 'gas/Api.gs', 'gas/ReservationV2.gs']) {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
}

function run(expression) {
  return vm.runInContext(expression, context);
}

const roomMap = {
  'TEST-ROOM-A': {
    room_id: 'TEST-ROOM-A',
    room_name: '会議室A',
    room_type: '単体',
    component_room_ids: [],
    calendar_id: 'calendar-a',
    calendar_url: 'calendar-a-url',
  },
  'TEST-ROOM-B': {
    room_id: 'TEST-ROOM-B',
    room_name: '会議室B',
    room_type: '単体',
    component_room_ids: [],
    calendar_id: 'calendar-b',
    calendar_url: 'calendar-b-url',
  },
  'TEST-ROOM-ALL': {
    room_id: 'TEST-ROOM-ALL',
    room_name: '大会議室',
    room_type: '大会議室',
    component_room_ids: ['TEST-ROOM-A', 'TEST-ROOM-B'],
    calendar_id: 'calendar-all',
    calendar_url: 'calendar-all-url',
  },
};
sandbox.testRoomMap = roomMap;

assert.equal(run("hasTimeOverlap_('10:00', '11:00', '10:30', '11:30')"), true);
assert.equal(run("hasTimeOverlap_('10:00', '11:00', '11:00', '12:00')"), false);
assert.equal(run("hasReservationRoomConflict_('TEST-ROOM-A', 'TEST-ROOM-B', testRoomMap)"), false);
assert.equal(run("hasReservationRoomConflict_('TEST-ROOM-A', 'TEST-ROOM-ALL', testRoomMap)"), true);
assert.equal(run("hasReservationRoomConflict_('TEST-ROOM-B', 'TEST-ROOM-ALL', testRoomMap)"), true);

sandbox.newReservation = {
  meeting_name: '新規会議',
  room_id: 'TEST-ROOM-A',
  usage_date: '2027-02-03',
  start_time: '10:00',
  end_time: '10:30',
};
sandbox.existingProcessing = [{
  reservation_id: 'R-PENDING',
  room_id: 'TEST-ROOM-ALL',
  room_name: '大会議室',
  usage_date: '2027-02-03',
  start_time: '09:45',
  end_time: '10:15',
  status: '処理中',
}];
sandbox.existingFailed = [{ ...sandbox.existingProcessing[0], status: '登録失敗' }];
assert.equal(run('findReservationConflictsV2_([newReservation], existingProcessing, testRoomMap).length'), 1);
assert.equal(run('findReservationConflictsV2_([newReservation], existingFailed, testRoomMap).length'), 0);

sandbox.validationContext = {
  settings: {
    MAX_RESERVATIONS_PER_SUBMIT: '10',
    FIELD_ORGANIZATION: '有効',
    FIELD_USER_NAME: '有効',
    BUSINESS_START_TIME: '09:00',
    BUSINESS_END_TIME: '22:00',
    TIME_SLOT_MINUTES: '5',
  },
  roomMap,
};
sandbox.validPayload = {
  common: { organization_name: 'テスト団体', user_name: 'テスト担当者' },
  reservations: [sandbox.newReservation],
};
assert.equal(run('validateReservationRequestV2_(validPayload, validationContext).ok'), true);

sandbox.overlapPayload = {
  common: sandbox.validPayload.common,
  reservations: [
    sandbox.newReservation,
    { ...sandbox.newReservation, room_id: 'TEST-ROOM-ALL', start_time: '10:15', end_time: '10:45' },
  ],
};
assert.equal(run('validateReservationRequestV2_(overlapPayload, validationContext).ok'), false);
assert.equal(run('validateReservationRequestV2_(overlapPayload, validationContext).errors.length'), 2);

sandbox.common = { line_user_id: '', organization_name: 'テスト団体', user_name: 'テスト担当者' };
sandbox.normalizedReservations = [sandbox.newReservation];
const firstHash = run('createReservationRequestHashV2_(common, normalizedReservations)');
const secondHash = run('createReservationRequestHashV2_(common, normalizedReservations)');
assert.equal(firstHash, secondHash);
assert.equal(firstHash.length, 64);

const pending = run("createPendingReservationRowsV2_('request-1', 'hash-1', common, normalizedReservations, testRoomMap, '2026-09-30T12:34:56')");
assert.equal(pending.rows.length, 1);
assert.equal(pending.rows[0].length, 19);
assert.equal(pending.rows[0][12], '処理中');
assert.equal(pending.rows[0][14], 'request-1');
assert.equal(pending.rows[0][15], 'hash-1');

sandbox.previousRows = [{
  ...pending.reservations[0],
  request_hash: 'hash-1',
  status: '処理中',
  error_message: '',
}];
assert.equal(run("buildPreviousReservationResultV2_('request-1', 'hash-1', previousRows, {}).state"), 'processing');
sandbox.previousRows[0].status = '有効';
assert.equal(run("buildPreviousReservationResultV2_('request-1', 'hash-1', previousRows, {}).state"), 'active');
assert.equal(run("buildPreviousReservationResultV2_('request-1', 'different-hash', previousRows, {}).code"), 'REQUEST_ID_MISMATCH');
sandbox.previousRows[0].status = '登録失敗';
sandbox.previousRows[0].error_message = 'calendar failed';
assert.equal(run("buildPreviousReservationResultV2_('request-1', 'hash-1', previousRows, {}).state"), 'failed');
sandbox.previousRows[0].status = '取消';
sandbox.previousRows[0].error_message = '';
assert.equal(run("buildPreviousReservationResultV2_('request-1', 'hash-1', previousRows, {}).state"), 'cancelled');

process.stdout.write('reservation_v2_unit_test: 21 assertions passed\n');
