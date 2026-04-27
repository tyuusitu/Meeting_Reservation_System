/**
 * フロントエンド全画面で共有する設定オブジェクト。
 *
 * <p>このファイルだけを書き換えれば、予約画面・予約確認画面・出欠登録画面・
 * 管理画面の読み込み先と書き込み先がまとめて切り替わります。</p>
 *
 * <p>SHEET_GID が設定されているシートは gid 指定の CSV を優先します。
 * 空欄の場合は SHEET_NAME のシート名指定にフォールバックします。</p>
 */
window.APP_CONFIG = {
  SPREADSHEET_ID: '1IRmvnXSuP_A44rOM4m1B4RTWcSQzzW1cZuxtizgJx_M',
  GAS_API_URL: 'https://script.google.com/macros/s/AKfycbyuklc-3W8N9J_pXZ9ockQnwkf_xxAF8WpwY6bvLYqC8HIQeXFLrnU6PCNpQAr0p-0Nkg/exec',
  SHEET_GID: {
    reservations: '929838420',
    rooms: '1209500136',
    settings: '1074437390',
    members: '',
    meetings: '',
    attendanceResponses: '',
    meetingAggregations: '',
    memberAggregations: '',
    departmentAggregations: '',
    streaks: '',
    logs: '',
  },
  SHEET_NAME: {
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
  },
  ABSENCE_REASONS: ['授業', 'アルバイト', '体調不良', '家庭の都合', 'その他'],
  CSV_CACHE_TTL_MS: 30000,
};

/**
 * 指定シートの CSV 取得 URL を生成します。
 *
 * @param {string} sheetKey SHEET_GID / SHEET_NAME に登録している論理キー。
 * @param {boolean} bustCache true の場合は現在時刻を付けてブラウザキャッシュを避けます。
 * @return {string} CSV 取得 URL。設定不足の場合は空文字。
 */
window.APP_CONFIG.buildSheetCsvUrl = function buildSheetCsvUrl(sheetKey, bustCache) {
  const gid = this.SHEET_GID && this.SHEET_GID[sheetKey];
  const sheetName = this.SHEET_NAME && this.SHEET_NAME[sheetKey];
  if (gid) {
    return `https://docs.google.com/spreadsheets/d/${this.SPREADSHEET_ID}/export?format=csv&gid=${encodeURIComponent(gid)}${bustCache ? `&t=${Date.now()}` : ''}`;
  }
  if (sheetName) {
    return `https://docs.google.com/spreadsheets/d/${this.SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}${bustCache ? `&t=${Date.now()}` : ''}`;
  }
  return '';
};

window.APP_CONFIG._csvTextCache = {};

window.APP_CONFIG.fetchSheetCsvText = async function fetchSheetCsvText(sheetKey, bustCache) {
  const url = this.buildSheetCsvUrl(sheetKey, bustCache);
  if (!url) return '';

  const cacheKey = sheetKey;
  const now = Date.now();
  const cached = this._csvTextCache[cacheKey];
  if (!bustCache && cached && cached.expiresAt > now) {
    return cached.text;
  }

  const res = await fetch(url, { cache: bustCache ? 'no-store' : 'default' });
  const text = await res.text();
  this._csvTextCache[cacheKey] = {
    text,
    expiresAt: now + Number(this.CSV_CACHE_TTL_MS || 30000),
  };
  return text;
};
