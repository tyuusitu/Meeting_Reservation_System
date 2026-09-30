#!/usr/bin/env node

import assert from 'node:assert/strict';
import fs from 'node:fs';

const reserveHtml = fs.readFileSync('reserve.html', 'utf8');
const apiSource = fs.readFileSync('gas/Api.gs', 'utf8');
const reservationV2Source = fs.readFileSync('gas/ReservationV2.gs', 'utf8');
const responsiveCss = fs.readFileSync('responsive.css', 'utf8');
const pages = ['index.html', 'reserve.html', 'status.html', 'attendance.html', 'admin.html'];

assert.doesNotMatch(reserveHtml, /RESERVATION_DONE_FALLBACK_MS/);
assert.doesNotMatch(reserveHtml, /createPendingReservations/);
assert.match(reserveHtml, /payload\.request_id = pending\.requestId/);
assert.match(reserveHtml, /action', 'get_reservation_status'/);
assert.match(reserveHtml, /keepalive: useReservationV2/);
assert.match(reserveHtml, /sessionStorage\.setItem/);
assert.match(reserveHtml, /data\.state === 'active'/);

assert.match(apiSource, /mode === RESERVATION_V2_MODES\.preparing/);
assert.match(apiSource, /return insertReservations\(payload\)/);
assert.match(reservationV2Source, /function enableReservationV2Release\(/);
assert.match(reservationV2Source, /function disableReservationV2Release\(/);
assert.match(reservationV2Source, /RESERVATION_V2_INDEX_READY_PROPERTY_KEY/);
assert.match(reservationV2Source, /cancelRequested/);

pages.forEach((page) => {
  const html = fs.readFileSync(page, 'utf8');
  assert.match(html, /viewport-fit=cover/, `${page}: safe-area viewport がありません。`);
  assert.match(html, /href="responsive\.css"/, `${page}: responsive.css が読み込まれていません。`);
});
assert.match(responsiveCss, /@media \(max-width: 640px\)/);
assert.match(responsiveCss, /gap: 18px/);
assert.match(responsiveCss, /env\(safe-area-inset-bottom\)/);

const inlineScripts = [...reserveHtml.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
  .map((match) => match[1])
  .filter(Boolean);
inlineScripts.forEach((source) => new Function(source));

process.stdout.write('release_readiness_test: checks passed\n');
