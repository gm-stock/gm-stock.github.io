/**
 * 출고 마스터 시트 → 재고 프로그램 연동 (출고 내역 시트에 붙이는 추가 파일)
 *
 * 하는 일 (5분마다)
 *  1) 구글 드라이브 '출고 마스터' 폴더의 "마스터 시트 _ (컬렉션명)" 파일들을 읽어
 *     Firestore masterSheets/{파일ID} 에 올린다.
 *       - title: 컬렉션 이름('최종 수량' B1), openDate: 공통 오픈 날짜(B2)
 *       - grid: 매장별 최종 수량 표 (첫 줄 = ["매장", 품목...], 다음 줄부터 매장별 수량) — JSON 글자
 *       - unsent: '미발송 현황' 탭의 매장별·품목별 아직 안 보낸 수량 { 매장: { 품목: 수량 } } — JSON 글자
 *     → 관리자 화면 '시즌 배송 → 새 시즌 등록'에서 컬렉션 이름을 누르면 이 내용으로 채워진다.
 *  2) 마스터 시트에서 등록한 시즌(shipments 문서의 master.fileId)이 있으면
 *     그 시즌 문서에 unsent(미발송 수량)를 같이 적어 둔다.
 *     → 매장 화면 '배송 이슈'와 관리자 접수 목록에 "아직 발송 전 N개"로 보인다.
 *  - 내용이 지난번과 같으면 아무것도 쓰지 않는다.
 *  - 함수 이름은 기존 파일들과 겹치지 않게 모두 ms_ 로 시작한다.
 *
 * 처음 한 번만
 *  1) 출고 내역 시트 → 확장 프로그램 → Apps Script
 *  2) 왼쪽 '파일' 옆 ＋ → 스크립트 → 이름: master_sheet_sync → 이 내용을 통째로 붙여넣고 저장
 *  3) 위쪽 함수 선택에서 ms_setup 을 고르고 ▶ 실행 → 권한 승인 (드라이브 읽기 권한이 새로 추가됨)
 *     ※ "권한이 없습니다(drive / spreadsheets)" 오류가 나면: 왼쪽 ⚙ 프로젝트 설정 →
 *        'appsscript.json 매니페스트 파일 표시' 체크 → appsscript.json 의 "oauthScopes" 목록에
 *        "https://www.googleapis.com/auth/drive.readonly",
 *        "https://www.googleapis.com/auth/spreadsheets" 두 줄을 추가하고 저장 → 다시 ms_setup 실행
 *
 * Firestore 접속은 기존 동기화와 같은 방식(ScriptApp.getOAuthToken, datastore 권한)을 쓴다.
 */
const MS_PROJECT = 'gm-stock-management';
const MS_FOLDER_ID = '1dVMAOGSBuRyGTuWkhPo2z5ufD-ajatmW'; // '출고 마스터' 폴더
const MS_FOLDER_NAME = '출고 마스터';
const MS_FILE_PREFIX = '마스터 시트 _';
const MS_TAB_QTY = '최종 수량';
const MS_TAB_UNSENT = '미발송 현황';

function ms_setup() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'ms_syncMasterSheets');
  if (!exists) ScriptApp.newTrigger('ms_syncMasterSheets').timeBased().everyMinutes(5).create();
  ms_syncMasterSheets();
}

function ms_syncMasterSheets() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { ms_syncInner_(); } finally { lock.releaseLock(); }
}

function ms_syncInner_() {
  const props = PropertiesService.getScriptProperties();
  const files = ms_masterFiles_();
  const unsentById = {};
  files.forEach(f => {
    let data;
    try { data = ms_readMaster_(f); }
    catch (e) { console.error('마스터 시트 읽기 실패: ' + f.getName() + ' / ' + e); return; }
    unsentById[f.getId()] = data.unsent;
    const key = 'ms_hash_' + f.getId();
    const hash = ms_hash_(data);
    if (props.getProperty(key) === hash) return;
    ms_patch_('masterSheets/' + f.getId(), {
      title: { stringValue: data.title },
      fileName: { stringValue: f.getName() },
      url: { stringValue: f.getUrl() },
      openDate: { stringValue: data.openDate },
      grid: { stringValue: data.grid },
      unsent: { stringValue: data.unsent },
      updatedAt: { timestampValue: new Date().toISOString() }
    }, false);
    props.setProperty(key, hash);
  });

  // 마스터 시트에서 등록한 시즌에 미발송 수량을 붙인다
  const ships = ms_listShipments_();
  ships.forEach(s => {
    const fid = s.masterFileId;
    if (!fid || !(fid in unsentById)) return;
    const u = unsentById[fid];
    if (s.unsent === u) return;
    ms_patch_('shipments/' + s.id, {
      unsent: { stringValue: u },
      unsentAt: { timestampValue: new Date().toISOString() }
    }, true);
  });
}

// '출고 마스터' 폴더 안의 "마스터 시트 _ ..." 스프레드시트들 (복사용 양식은 제외)
function ms_masterFiles_() {
  let folder = null;
  try { folder = DriveApp.getFolderById(MS_FOLDER_ID); } catch (e) { folder = null; }
  if (!folder) {
    const it = DriveApp.getFoldersByName(MS_FOLDER_NAME);
    if (!it.hasNext()) throw new Error("'" + MS_FOLDER_NAME + "' 폴더를 찾지 못했어요.");
    folder = it.next();
  }
  const out = [];
  const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    const f = it.next();
    if (f.isTrashed()) continue;
    if (String(f.getName()).indexOf(MS_FILE_PREFIX) !== 0) continue;
    out.push(f);
  }
  return out;
}

function ms_readMaster_(file) {
  const ss = SpreadsheetApp.openById(file.getId());
  const sh = ss.getSheetByName(MS_TAB_QTY);
  if (!sh) throw new Error("'" + MS_TAB_QTY + "' 탭이 없어요.");
  const values = sh.getRange(1, 1, sh.getLastRow(), Math.min(sh.getLastColumn(), 22)).getValues(); // A~V
  const title = String(values[0][1] || '').trim() || String(file.getName()).replace(MS_FILE_PREFIX, '').trim();
  const od = values[1] ? values[1][1] : '';
  const openDate = od instanceof Date ? Utilities.formatDate(od, 'Asia/Seoul', 'yyyy-MM-dd') : String(od || '').trim();

  // 5행: 품목 머리글 (H~V)
  const head = values[4] || [];
  const cols = [];
  for (let c = 7; c < 22; c++) { const h = String(head[c] || '').trim(); if (h) cols.push({ c: c, h: h }); }
  const grid = [['매장'].concat(cols.map(x => x.h))];
  // 6행부터: 국가(B열)가 있는 줄 = 매장 (소계·합계·빈 줄은 국가가 비어 있음)
  for (let r = 5; r < values.length; r++) {
    const store = String(values[r][0] || '').trim();
    const cc = String(values[r][1] || '').trim();
    if (!store || !cc) continue;
    grid.push([store].concat(cols.map(x => Number(values[r][x.c]) || 0)));
  }

  // 미발송 현황: 4행 머리글(D열부터), 5행부터 매장
  const unsent = {};
  const us = ss.getSheetByName(MS_TAB_UNSENT);
  if (us && us.getLastRow() >= 5) {
    const uv = us.getRange(4, 1, us.getLastRow() - 3, us.getLastColumn()).getValues();
    const uh = uv[0];
    for (let r = 1; r < uv.length; r++) {
      const store = String(uv[r][0] || '').trim();
      if (!store) continue;
      for (let c = 3; c < uh.length; c++) {
        const h = String(uh[c] || '').trim();
        const n = Number(uv[r][c]);
        if (h && n > 0) { (unsent[store] = unsent[store] || {})[h] = n; }
      }
    }
  }
  return { title: title, openDate: openDate, grid: JSON.stringify(grid), unsent: JSON.stringify(unsent) };
}

function ms_hash_(o) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(o)));
}

function ms_base_() {
  return 'https://firestore.googleapis.com/v1/projects/' + MS_PROJECT + '/databases/(default)/documents/';
}

// 문서의 일부 칸만 고쳐 쓴다 (mustExist=true 면 문서가 있을 때만)
function ms_patch_(path, fields, mustExist) {
  const mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + encodeURIComponent(k)).join('&');
  const url = ms_base_() + path + '?' + mask + (mustExist ? '&currentDocument.exists=true' : '');
  const res = UrlFetchApp.fetch(url, {
    method: 'patch', contentType: 'application/json', payload: JSON.stringify({ fields: fields }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error(path + ' 저장 실패 (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300));
}

// 시즌 문서 중 마스터 시트 연결 정보와 미발송 값만 읽는다
function ms_listShipments_() {
  const out = [];
  let token = '';
  do {
    const url = ms_base_() + 'shipments?pageSize=300&mask.fieldPaths=master&mask.fieldPaths=unsent'
      + (token ? '&pageToken=' + encodeURIComponent(token) : '');
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('shipments 읽기 실패 (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300));
    const body = JSON.parse(res.getContentText());
    (body.documents || []).forEach(d => {
      const f = d.fields || {};
      const m = (f.master && f.master.mapValue && f.master.mapValue.fields) || {};
      out.push({
        id: d.name.split('/').pop(),
        masterFileId: m.fileId ? m.fileId.stringValue : '',
        unsent: f.unsent ? f.unsent.stringValue : null
      });
    });
    token = body.nextPageToken || '';
  } while (token);
  return out;
}
