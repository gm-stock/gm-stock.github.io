/**
 * 시즌 배송 · 배송 이슈 → 구글 시트 기록 (출고 내역 시트에 붙이는 추가 파일)
 *
 * 하는 일
 *  - 5분마다 Firestore의 shipments(시즌 배송), claims(배송 이슈 접수)를 읽어
 *    이 스프레드시트의 '시즌 배송', '배송 이슈' 탭을 통째로 다시 씁니다. (읽기 전용 기록용)
 *  - 기존 출고 내역 동기화(apps_script_v4.gs)와는 함수 이름이 겹치지 않게 모두 ss_ 로 시작합니다.
 *
 * 처음 한 번만
 *  1) 출고 내역 시트 → 확장 프로그램 → Apps Script
 *  2) 왼쪽 '파일' 옆 ＋ → 스크립트 → 이름: season_claims_sync → 이 내용을 통째로 붙여넣고 저장
 *  3) 위쪽 함수 선택에서 ss_setup 을 고르고 ▶ 실행 → 권한 승인
 *     (탭 2개가 만들어지고, 5분마다 자동으로 갱신되는 예약이 걸립니다)
 *
 * Firestore 접속은 기존 동기화와 같은 방식(ScriptApp.getOAuthToken, datastore 권한)을 씁니다.
 */
const SS_PROJECT = 'gm-stock-management';
const SS_TAB_SHIP = '시즌 배송';
const SS_TAB_CLAIM = '배송 이슈';
const SS_CLAIM_DAYS = 7;

function ss_setup() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'ss_syncSeasonSheets');
  if (!exists) ScriptApp.newTrigger('ss_syncSeasonSheets').timeBased().everyMinutes(5).create();
  ss_syncSeasonSheets();
}

function ss_syncSeasonSheets() {
  // 예약 실행이 겹치면(앞 실행이 아직 안 끝났으면) 이번 차례는 건너뛴다
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { ss_syncSeasonSheetsInner_(); } finally { lock.releaseLock(); }
}
function ss_syncSeasonSheetsInner_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ships = ss_list_('shipments', null);
  // 사진(용량이 큼)은 빼고 필요한 칸만 받는다
  const claims = ss_list_('claims', ['shipmentId', 'storeName', 'itemName', 'optName', 'collection', 'type', 'reason',
    'qty', 'shipped', 'reporter', 'status', 'date', 'createdAt', 'handledAt', 'handler']);
  const shipById = {};
  ships.forEach(s => { shipById[s.id] = s; });

  // ── 시즌 배송 ──
  const shipRows = [];
  ships.sort((a, b) => String(b.openDate || '').localeCompare(String(a.openDate || '')));
  ships.forEach(s => {
    const deadline = ss_deadline_(s.openDate);
    Object.keys(s.stores || {}).sort().forEach(store => {
      (s.stores[store] || []).forEach(it => {
        shipRows.push([s.title || '', s.openDate || '', deadline, s.open ? '접수 중' : '접수 닫힘',
          store, it.collection || '', it.itemName || '', it.optName || '', Number(it.qty) || 0]);
      });
    });
  });
  ss_write_(ss, SS_TAB_SHIP,
    ['시즌', '오픈일', '접수 기한', '접수 상태', '매장', '컬렉션', '품목', '옵션', '배송 수량'], shipRows);

  // ── 배송 이슈 ──
  claims.sort((a, b) => ss_time_(b.createdAt) - ss_time_(a.createdAt));
  const claimRows = claims.map(c => [
    c.createdAt instanceof Date ? Utilities.formatDate(c.createdAt, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : (c.date || ''),
    (shipById[c.shipmentId] || {}).title || '', c.storeName || '', c.collection || '', c.itemName || '', c.optName || '',
    c.type || '', Number(c.qty) || 0, Number(c.shipped) || 0, c.reason || '', c.reporter || '', c.status || '',
    c.handledAt ? Utilities.formatDate(new Date(c.handledAt), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : '', c.handler || ''
  ]);
  const claimChanged = ss_write_(ss, SS_TAB_CLAIM,
    ['접수일시', '시즌', '매장', '컬렉션', '품목', '옵션', '유형', '재요청 수량', '배송 수량', '사유', '접수자', '처리 상태', '처리일시', '처리자'],
    claimRows);
  // 처리 상태 색 (검토중 주황 / 재발송 예정·완료 초록 / 반려 빨강)
  const sh = ss.getSheetByName(SS_TAB_CLAIM);
  if (claimChanged && claimRows.length) {
    const colors = claimRows.map(r => [r[11] === '검토중' ? '#fdebd9' : (r[11] === '반려' ? '#f8d7da' : '#e3f1df')]);
    ss_retry_(() => sh.getRange(2, 12, claimRows.length, 1).setBackgrounds(colors));
  }
}

// 시트 탭을 비우고 머리글 + 내용을 한 번에 쓴다
// - 내용이 지난번과 같으면 아무것도 하지 않는다 (5분마다 돌아도 시트에 부담이 없게)
// - 시트가 바빠서 시간 초과가 나면 잠깐 쉬었다가 최대 3번까지 다시 시도한다
function ss_write_(ss, name, header, rows) {
  const props = PropertiesService.getScriptProperties();
  const key = 'ss_hash_' + name;
  const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([header, rows])));
  let sh = ss.getSheetByName(name);
  if (sh && props.getProperty(key) === hash) return false;
  ss_retry_(() => {
    if (!sh) { sh = ss.insertSheet(name); }
    else { sh.clearContents(); sh.clearFormats(); }
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
    if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.setFrozenRows(1);
    SpreadsheetApp.flush();
  });
  props.setProperty(key, hash);
  return true;
}
function ss_retry_(fn) {
  for (let i = 0; i < 3; i++) {
    try { return fn(); }
    catch (e) {
      if (i === 2 || !/timed out|시간|Service/i.test(String(e))) throw e;
      Utilities.sleep(5000 * (i + 1));
    }
  }
}

function ss_deadline_(openDate) {
  if (!openDate) return '';
  const d = new Date(openDate + 'T00:00:00+09:00');
  d.setDate(d.getDate() + SS_CLAIM_DAYS);
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd') + ' 24:00';
}
function ss_time_(v) { return v instanceof Date ? v.getTime() : 0; }

// Firestore 컬렉션 전체 읽기 (fields: 받을 칸 목록, null이면 전부)
function ss_list_(col, fields) {
  const base = 'https://firestore.googleapis.com/v1/projects/' + SS_PROJECT + '/databases/(default)/documents/' + col;
  const mask = fields ? fields.map(f => '&mask.fieldPaths=' + encodeURIComponent(f)).join('') : '';
  const out = [];
  let token = '';
  do {
    const url = base + '?pageSize=300' + mask + (token ? '&pageToken=' + encodeURIComponent(token) : '');
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error(col + ' 읽기 실패 (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300));
    const body = JSON.parse(res.getContentText());
    (body.documents || []).forEach(d => {
      const o = ss_fields_(d.fields || {});
      o.id = d.name.split('/').pop();
      out.push(o);
    });
    token = body.nextPageToken || '';
  } while (token);
  return out;
}
function ss_fields_(f) { const o = {}; Object.keys(f).forEach(k => { o[k] = ss_val_(f[k]); }); return o; }
function ss_val_(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('timestampValue' in v) return new Date(v.timestampValue);
  if ('nullValue' in v) return null;
  if ('mapValue' in v) return ss_fields_(v.mapValue.fields || {});
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(ss_val_);
  return null;
}
