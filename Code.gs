/**
 * 알레르기 예방 액션카드 - 실천률 취합 (Google Apps Script)
 *
 * 학생이 '선생님께 보내기'를 누르면 [실천 현황] 탭이 자동으로 정리돼요.
 *
 *   [실천 현황]  이름 | 실천률(%)      ← 선생님이 보는 표 (학생마다 한 줄)
 *   [_기록]      계산용 보관 탭 (숨겨져 있어요. 날짜·이름·실천 수만 저장)
 *
 *   - 실천률 = 학생이 보낸 날들의 하루 실천률 평균
 *     (예: 1일차 80%, 2일차 100% → 90%)
 *   - 같은 날 여러 번 보내면 그날 '마지막' 기록만 계산해요.
 *   - 날짜는 학생 기기가 아니라 서버(한국 시간) 기준이에요.
 *   - 학생은 이름으로 구분해요. 이름이 같은 학생은 '김지민A'처럼 구분해서 쓰게 해 주세요.
 *   - 개인정보 보호: 수칙 내용은 받지 않아요. 이름과 '몇 개 중 몇 개'만 받아요.
 *
 * [설치·업데이트 방법]
 * 1. 구글 시트 → 확장 프로그램 → Apps Script 에 이 코드를 통째로 붙여 넣고 저장합니다.
 * 2. 함수 목록에서 setup 을 골라 ▶실행 → [실천 현황] 탭이 생깁니다.
 * 3. 처음이라면: 배포 > 새 배포 > 웹 앱 (실행 사용자: 나 / 액세스: 모든 사용자)
 *    이미 배포했다면: 배포 > 배포 관리 > ✏️ 수정 > 버전: 새 버전 > 배포  (주소가 그대로 유지돼요)
 *
 * 예전 버전에서 만든 [실천기록], [학생별 요약], [날짜별 현황] 탭은 더 이상 쓰지 않아요. 지워도 돼요.
 * 시트 위쪽 메뉴 [🛡️ 알레르기 카드 > 표 새로 고침]으로 언제든 다시 계산할 수 있어요.
 */

const TZ = 'Asia/Seoul';
const MAIN_SHEET = '실천 현황';
const RAW_SHEET = '_기록';
const RAW_HEADERS = ['받은 시각', '날짜', '이름', '실천 수', '전체 수칙 수'];

const COLOR_HIGH = '#C9F2E6';  // 80% 이상: 민트
const COLOR_MID = '#FFF3B8';   // 50~79%: 연노랑
const COLOR_LOW = '#FFDDE6';   // 50% 미만: 연분홍
const COLOR_HEAD = '#D4ECFF';

/* =========================================================
   웹앱 입구
   ========================================================= */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = clean_(data.name).slice(0, 20);
    if (!name) throw new Error('이름이 없습니다.');
    const total = toInt_(data.total);
    const done = Math.min(toInt_(data.done), total);
    if (!total) throw new Error('수칙 수가 없습니다.');

    const now = new Date();
    getRawSheet_().appendRow([
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
      Utilities.formatDate(now, TZ, 'yyyy-MM-dd'),
      name, done, total
    ]);
    rebuild_();
    return json_({ result: 'success' });
  } catch (err) {
    return json_({ result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ result: 'ok', message: '알레르기 예방 액션카드 서버가 작동 중입니다.' });
}

/** 처음 한 번 실행 */
function setup() {
  getRawSheet_();
  rebuild_();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛡️ 알레르기 카드')
    .addItem('표 새로 고침', 'rebuildSummaries')
    .addToUi();
}

function rebuildSummaries() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { rebuild_(); } finally { lock.releaseLock(); }
}

/* =========================================================
   취합: _기록 → 실천 현황 (이름 | 실천률)
   ========================================================= */

function rebuild_() {
  const values = getRawSheet_().getDataRange().getValues();
  const map = {}; // 이름 → { 날짜: 실천률 }
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const day = dayStr_(r[1]);
    const name = String(r[2] || '').trim();
    const done = Number(r[3]) || 0, total = Number(r[4]) || 0;
    if (!day || !name || !total) continue;
    if (!map[name]) map[name] = {};
    map[name][day] = Math.round((done / total) * 100); // 같은 날은 마지막 기록으로 덮어씀
  }
  const rows = Object.keys(map).sort(function (a, b) { return a.localeCompare(b, 'ko'); }).map(function (name) {
    const rates = Object.keys(map[name]).map(function (d) { return map[name][d]; });
    const avg = Math.round(rates.reduce(function (a, b) { return a + b; }, 0) / rates.length);
    return [name, avg];
  });

  const sheet = getOrCreate_(MAIN_SHEET);
  sheet.clear();
  const all = [['이름', '실천률(%)']].concat(rows);
  sheet.getRange(1, 1, all.length, 2).setValues(all);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground(COLOR_HEAD).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 110);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setHorizontalAlignment('center');
    sheet.getRange(2, 2, rows.length, 1).setBackgrounds(rows.map(function (r) { return [rateColor_(r[1])]; }));
  }
  sheet.getRange(all.length + 2, 1).setValue(
    '※ 실천률 = 보낸 날들의 평균 · 업데이트 ' + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm')
  ).setFontColor('#6E7D88');
}

/* =========================================================
   내부 도우미
   ========================================================= */

function getRawSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RAW_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RAW_SHEET);
    sheet.getRange('A:B').setNumberFormat('@'); // 날짜를 글자 그대로 보관
    sheet.appendRow(RAW_HEADERS);
    sheet.getRange(1, 1, 1, RAW_HEADERS.length).setFontWeight('bold').setBackground(COLOR_HEAD);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function getOrCreate_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, 0);
    ss.setActiveSheet(sheet);
  }
  return sheet;
}

function rateColor_(v) {
  const n = Number(v);
  if (n >= 80) return COLOR_HIGH;
  if (n >= 50) return COLOR_MID;
  return COLOR_LOW;
}

function dayStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function toInt_(v) {
  const n = parseInt(v, 10);
  return n > 0 && n < 1000 ? n : 0;
}

/** 시트 수식 실행 방지(=, +, -, @ 로 시작하면 앞에 ' 붙이기) + 공백 정리 */
function clean_(v) {
  let s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
