#!/usr/bin/env node

import { mkdirSync, writeFileSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { dirname, resolve } from 'node:path';

const gasUrl = process.env.TEST_GAS_URL || '';
const runId = process.env.TEST_RUN_ID || `v2-${new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14)}-${randomUUID().slice(0, 6)}`;
const outputPath = resolve(process.env.TEST_RESULT_PATH || `test-results/reservation-v2-${runId}.json`);
const dryRun = process.argv.includes('--dry-run');
const baseDate = process.env.TEST_BASE_DATE || '2027-02-01';

function addDays(dateText, days) {
  const [year, month, day] = dateText.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day + days));
  return date.toISOString().slice(0, 10);
}

function timeText(totalMinutes) {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

function makeCase({ scenario, index, requestId, roomId, usageDate, startMinutes, durationMinutes = 10, meetingName }) {
  return {
    scenario,
    index,
    requestId,
    payload: {
      request_id: requestId,
      common: {
        organization_name: '[TEST] 100回検証',
        user_name: '[TEST] 自動検証',
      },
      reservations: [{
        meeting_name: meetingName || `[TEST] ${scenario}-${index}`,
        room_id: roomId,
        usage_date: usageDate,
        start_time: timeText(startMinutes),
        end_time: timeText(startMinutes + durationMinutes),
      }],
    },
  };
}

function buildCases() {
  const unique = [];
  for (let index = 0; index < 20; index += 1) {
    unique.push(makeCase({
      scenario: 'unique-a',
      index,
      requestId: `${runId}-unique-a-${index}`,
      roomId: 'TEST-ROOM-A',
      usageDate: addDays(baseDate, 0),
      startMinutes: 9 * 60 + index * 15,
    }));
    unique.push(makeCase({
      scenario: 'unique-b',
      index,
      requestId: `${runId}-unique-b-${index}`,
      roomId: 'TEST-ROOM-B',
      usageDate: addDays(baseDate, 0),
      startMinutes: 9 * 60 + index * 15,
    }));
    unique.push(makeCase({
      scenario: 'unique-all',
      index,
      requestId: `${runId}-unique-all-${index}`,
      roomId: 'TEST-ROOM-ALL',
      usageDate: addDays(baseDate, 1),
      startMinutes: 9 * 60 + index * 15,
    }));
  }

  const contention = Array.from({ length: 20 }, (_, index) => makeCase({
    scenario: 'contention',
    index,
    requestId: `${runId}-contention-${index}`,
    roomId: index % 2 === 0 ? 'TEST-ROOM-A' : 'TEST-ROOM-ALL',
    usageDate: addDays(baseDate, 2),
    startMinutes: 10 * 60,
    durationMinutes: 30,
  }));

  const idempotency = [];
  for (let index = 0; index < 10; index += 1) {
    const requestId = `${runId}-idempotent-${index}`;
    const original = makeCase({
      scenario: 'idempotency',
      index,
      requestId,
      roomId: index % 2 === 0 ? 'TEST-ROOM-A' : 'TEST-ROOM-B',
      usageDate: addDays(baseDate, 3),
      startMinutes: 9 * 60 + index * 20,
      meetingName: `[TEST] idempotency-${index}`,
    });
    idempotency.push(original, JSON.parse(JSON.stringify(original)));
  }
  return { unique, contention, idempotency };
}

async function submit(testCase) {
  const startedAt = Date.now();
  const attempts = [];
  for (let attempt = 1; attempt <= 6; attempt += 1) {
    try {
      const response = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ action: 'insert_reservations', payload: testCase.payload }),
        redirect: 'follow',
      });
      const text = await response.text();
      let data;
      try {
        data = JSON.parse(text);
      } catch (_) {
        data = { ok: false, state: 'invalid_response', error: text.slice(0, 500) };
      }
      attempts.push({ attempt, http_status: response.status, state: data.state || '', metrics: data.metrics || null });
      if (data.state === 'processing') {
        return recoverViaStatus(testCase, startedAt, attempts, response.status, data);
      }
      const isTransient = data.state === 'busy'
        || data.state === 'invalid_response'
        || (!data.state && data.ok !== true);
      if (!isTransient || attempt === 6) {
        if (isTransient) {
          return recoverViaStatus(testCase, startedAt, attempts, response.status, data);
        }
        return {
          scenario: testCase.scenario,
          index: testCase.index,
          request_id: testCase.requestId,
          http_status: response.status,
          client_ms: Date.now() - startedAt,
          transport_attempts: attempts,
          response: data,
        };
      }
      const retryAfterMs = Math.max(250, Number(data.retry_after_ms || 500)) * Math.min(attempt, 3);
      await new Promise((resolveWait) => setTimeout(resolveWait, retryAfterMs + Math.floor(Math.random() * 350)));
    } catch (error) {
      attempts.push({ attempt, http_status: 0, state: 'network_error', error: error.message });
      if (attempt === 6) {
        return recoverViaStatus(
          testCase,
          startedAt,
          attempts,
          0,
          { ok: false, state: 'network_error', error: error.message }
        );
      }
      await new Promise((resolveWait) => setTimeout(resolveWait, 500 * attempt + Math.floor(Math.random() * 350)));
    }
  }
  throw new Error('到達不能な状態です。');
}

async function recoverViaStatus(testCase, startedAt, attempts, fallbackHttpStatus, fallbackResponse) {
  const statusChecks = [];
  for (let check = 1; check <= 8; check += 1) {
    try {
      const url = new URL(gasUrl);
      url.searchParams.set('action', 'get_reservation_status');
      url.searchParams.set('request_id', testCase.requestId);
      const data = await fetchJsonWithRetry(url, 3);
      statusChecks.push({ check, state: data.state || '', ok: data.ok === true });
      if (data.state === 'active' || data.state === 'failed') {
        return {
          scenario: testCase.scenario,
          index: testCase.index,
          request_id: testCase.requestId,
          http_status: 200,
          client_ms: Date.now() - startedAt,
          transport_attempts: attempts,
          status_checks: statusChecks,
          recovered_via_status: true,
          response: data,
        };
      }
    } catch (error) {
      statusChecks.push({ check, state: 'network_error', ok: false, error: error.message });
    }
    await new Promise((resolveWait) => setTimeout(resolveWait, 500 + check * 250));
  }
  return {
    scenario: testCase.scenario,
    index: testCase.index,
    request_id: testCase.requestId,
    http_status: fallbackHttpStatus,
    client_ms: Date.now() - startedAt,
    transport_attempts: attempts,
    status_checks: statusChecks,
    recovered_via_status: false,
    response: fallbackResponse,
  };
}

async function runInBatches(cases, concurrency) {
  const results = [];
  for (let index = 0; index < cases.length; index += concurrency) {
    const batch = cases.slice(index, index + concurrency);
    results.push(...await Promise.all(batch.map(submit)));
  }
  return results;
}

function percentile(values, rate) {
  if (!values.length) return 0;
  const sorted = values.slice().sort((left, right) => left - right);
  return sorted[Math.min(sorted.length - 1, Math.ceil(sorted.length * rate) - 1)];
}

function effectiveRoomIds(roomId) {
  return roomId === 'TEST-ROOM-ALL' ? ['TEST-ROOM-A', 'TEST-ROOM-B'] : [roomId];
}

function hasOverlap(left, right) {
  if (left.usage_date !== right.usage_date) return false;
  if (!effectiveRoomIds(left.room_id).some((id) => effectiveRoomIds(right.room_id).includes(id))) return false;
  const toMinutes = (text) => {
    const [hours, minutes] = text.split(':').map(Number);
    return hours * 60 + minutes;
  };
  return toMinutes(left.start_time) < toMinutes(right.end_time)
    && toMinutes(right.start_time) < toMinutes(left.end_time);
}

async function fetchJsonWithRetry(url, attempts = 6) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetch(url, { redirect: 'follow', cache: 'no-store' });
      const text = await response.text();
      try {
        return JSON.parse(text);
      } catch (_) {
        lastError = new Error(`JSON以外の応答です（HTTP ${response.status}）`);
      }
    } catch (error) {
      lastError = error;
    }
    if (attempt < attempts) {
      await new Promise((resolveWait) => setTimeout(resolveWait, 500 * attempt + Math.floor(Math.random() * 350)));
    }
  }
  throw lastError || new Error('JSON応答を取得できませんでした。');
}

async function fetchSummary() {
  const url = new URL(gasUrl);
  url.searchParams.set('action', 'get_reservation_test_summary');
  url.searchParams.set('run_id', runId);
  return fetchJsonWithRetry(url);
}

async function verifyRemoteTestEnvironment() {
  const summary = await fetchSummary();
  if (!summary.ok || summary.run_id !== runId || !Array.isArray(summary.rows)) {
    throw new Error('接続先が予約v2テスト環境であることを確認できません。予約送信前に停止しました。');
  }
  if (summary.rows.length !== 0) {
    throw new Error(`run_id ${runId} はすでに使用されています。予約送信前に停止しました。`);
  }

  const configUrl = new URL(gasUrl);
  configUrl.searchParams.set('action', 'get_config');
  const config = await fetchJsonWithRetry(configUrl);
  const roomIds = Array.isArray(config.rooms) ? config.rooms.map((room) => room.room_id) : [];
  if (!config.ok || roomIds.length !== 3 || roomIds.some((roomId) => !String(roomId).startsWith('TEST-'))) {
    throw new Error('接続先にテスト用以外の会議室が含まれています。予約送信前に停止しました。');
  }
  return { summary_ok: true, room_ids: roomIds };
}

function validateResults(results, summary) {
  const issues = [];
  const rows = summary.rows || [];
  const activeRows = rows.filter((row) => row.status === '有効');
  const uniqueResults = results.filter((item) => item.scenario.startsWith('unique-'));
  const contentionResults = results.filter((item) => item.scenario === 'contention');
  const idempotencyResults = results.filter((item) => item.scenario === 'idempotency');

  if (results.length !== 100) issues.push(`HTTP試行数が100ではありません: ${results.length}`);
  if (uniqueResults.filter((item) => item.response.state === 'active').length !== 60) {
    issues.push('固有枠60件のすべてがactiveになっていません。');
  }
  if (contentionResults.filter((item) => item.response.state === 'active').length !== 1) {
    issues.push('同一枠20件でactiveレスポンスが1件になっていません。');
  }
  if (contentionResults.filter((item) => item.response.state === 'conflict').length !== 19) {
    issues.push('同一枠20件でconflictレスポンスが19件になっていません。');
  }
  if (idempotencyResults.filter((item) => item.response.state === 'active').length !== 20) {
    issues.push('冪等テスト20試行の最終状態がすべてactiveになっていません。');
  }

  const idempotentRows = activeRows.filter((row) => row.request_id.includes('-idempotent-'));
  const idempotentRequestCounts = new Map();
  idempotentRows.forEach((row) => idempotentRequestCounts.set(row.request_id, (idempotentRequestCounts.get(row.request_id) || 0) + 1));
  if (idempotentRequestCounts.size !== 10 || [...idempotentRequestCounts.values()].some((count) => count !== 1)) {
    issues.push('冪等テスト10組がrequest_idごとに1行になっていません。');
  }

  if (activeRows.length !== 71) issues.push(`有効予約行が期待値71件ではありません: ${activeRows.length}`);
  if (rows.some((row) => row.status === '処理中')) issues.push('処理中のまま残った予約があります。');
  if (rows.some((row) => row.status === '登録失敗')) issues.push('登録失敗の予約行があります。');

  for (let leftIndex = 0; leftIndex < activeRows.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < activeRows.length; rightIndex += 1) {
      if (hasOverlap(activeRows[leftIndex], activeRows[rightIndex])) {
        issues.push(`二重予約を検出: ${activeRows[leftIndex].reservation_id} / ${activeRows[rightIndex].reservation_id}`);
      }
    }
  }

  const expectedEventRefsMinimum = 111;
  const expectedEventRefsMaximum = 113;
  const eventRefCount = activeRows.reduce((sum, row) => sum + Number(row.event_ref_count || 0), 0);
  if (eventRefCount < expectedEventRefsMinimum || eventRefCount > expectedEventRefsMaximum) {
    issues.push(`カレンダー予定参照数が期待範囲外です: ${eventRefCount}`);
  }

  return {
    passed: issues.length === 0,
    issues,
    counts: {
      attempts: results.length,
      unique_active_responses: uniqueResults.filter((item) => item.response.state === 'active').length,
      contention_active_responses: contentionResults.filter((item) => item.response.state === 'active').length,
      contention_conflict_responses: contentionResults.filter((item) => item.response.state === 'conflict').length,
      idempotency_attempts: idempotencyResults.length,
      idempotency_active_responses: idempotencyResults.filter((item) => item.response.state === 'active').length,
      physical_http_attempts: results.reduce((sum, item) => sum + item.transport_attempts.length, 0),
      busy_responses_retried: results.reduce((sum, item) => sum + item.transport_attempts.filter((attempt) => attempt.state === 'busy').length, 0),
      transient_responses_retried: results.reduce((sum, item) => sum + item.transport_attempts.filter((attempt) => attempt.state === 'invalid_response' || attempt.state === 'network_error').length, 0),
      recovered_via_status: results.filter((item) => item.recovered_via_status).length,
      status_check_attempts: results.reduce((sum, item) => sum + (item.status_checks || []).length, 0),
      active_rows: activeRows.length,
      event_refs: eventRefCount,
    },
  };
}

const cases = buildCases();
const plan = {
  run_id: runId,
  total_attempts: cases.unique.length + cases.contention.length + cases.idempotency.length,
  scenarios: {
    unique: cases.unique.length,
    contention: cases.contention.length,
    idempotency: cases.idempotency.length,
  },
  expected_active_rows: 71,
  expected_conflicts: 19,
  base_date: baseDate,
};

if (dryRun) {
  process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`);
  process.exit(0);
}

if (!gasUrl || !/^https:\/\/script\.google\.com\/macros\/s\//.test(gasUrl)) {
  throw new Error('TEST_GAS_URL にテスト用GAS WebアプリURLを設定してください。');
}

const startedAt = new Date().toISOString();
const environmentVerification = await verifyRemoteTestEnvironment();
const uniqueResults = await runInBatches(cases.unique, 10);
const contentionResults = await runInBatches(cases.contention, 20);
const idempotencyResults = await runInBatches(cases.idempotency, 20);
const results = [...uniqueResults, ...contentionResults, ...idempotencyResults];
const summary = await fetchSummary();
const validation = validateResults(results, summary);
const clientDurations = results.map((item) => item.client_ms);
const serverTotals = results.map((item) => Number(item.response.metrics?.total_ms || 0)).filter((value) => value > 0);
const report = {
  ...plan,
  started_at: startedAt,
  finished_at: new Date().toISOString(),
  environment_verification: environmentVerification,
  validation,
  latency_ms: {
    client: {
      min: Math.min(...clientDurations),
      p50: percentile(clientDurations, 0.50),
      p95: percentile(clientDurations, 0.95),
      max: Math.max(...clientDurations),
    },
    server: serverTotals.length ? {
      min: Math.min(...serverTotals),
      p50: percentile(serverTotals, 0.50),
      p95: percentile(serverTotals, 0.95),
      max: Math.max(...serverTotals),
    } : null,
  },
  responses: results,
  sheet_summary: summary,
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
process.stdout.write(`${JSON.stringify({ output: outputPath, ...validation, latency_ms: report.latency_ms }, null, 2)}\n`);
process.exit(validation.passed ? 0 : 1);
