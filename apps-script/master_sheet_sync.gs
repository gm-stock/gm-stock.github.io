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
    try { mo_check_(f.getId()); } catch (e) { console.error('품목 순서 확인 실패: ' + f.getName() + ' / ' + e); }
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
  // 6행부터: 국가(B열)가 있는 줄 = 매장 (6행 사이즈 줄·소계·합계·빈 줄은 국가가 비어 있어 건너뜀)
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

/* =====================================================================
 * 품목 순서를 바꿔도 수량이 따라가게 (2026-11-23 추가, 함수 이름은 mo_ 로 시작)
 *
 *  - '품목 정보' 5~19행 = '최종 수량' H~V열. 4~6행(사진·품목명·사이즈)은 INDEX 수식이라
 *    품목 정보 줄을 옮겨도 "5번째 줄 → H열" 자리가 그대로다.
 *  - '품목 정보'의 숨긴 '고유번호(자동)' 열에 품목마다 번호를 붙이고,
 *    '최종 수량' H~V 각 열이 어느 품목 것인지 기억해 둔다(스크립트 속성 mo_order_파일ID).
 *  - 품목 줄을 옮기거나 끼워 넣으면 → 매장 줄의 수량 숫자와 칸 색을 그 품목 열로 같이 옮긴다.
 *    (소계·합계처럼 수식이 있는 줄은 건드리지 않는다)
 *  - 지운 품목(또는 16번째로 밀려난 품목)에 수량이 있었으면 '지운 품목 수량(백업)' 탭에 남기고
 *    '변경 이력'에도 한 줄 적는다. 수량을 그냥 버리지 않는다.
 *  - 실행 시점: 마스터 시트마다 '변경 시' 트리거(5분 동기화가 자동 설치) + 5분 동기화 때 한 번 더 확인.
 *  - 새 마스터 시트(복사본)는 '복사용' 양식에 이미 고유번호 열과 수식이 있어 따로 할 일 없음.
 *    옛 양식 파일은 mo_prepareFile('파일ID') 를 한 번 실행하면 준비된다.
 * ===================================================================== */
const MO_INFO_SID = 6;   // '품목 정보' 탭 고유 번호
const MO_QTY_SID = 1;    // '최종 수량'
const MO_LOG_SID = 5;    // '변경 이력'
const MO_FIRST = 5;      // 품목 정보 첫 품목 줄
const MO_SLOTS = 15;     // H~V
const MO_QTY_COL = 8;    // H
const MO_QTY_ROW = 7;    // 최종 수량 매장 첫 줄
const MO_ID_HEAD = '고유번호(자동)';
const MO_BACKUP = '지운 품목 수량(백업)';

function mo_sheet_(ss, sid) {
  const shs = ss.getSheets();
  for (let i = 0; i < shs.length; i++) if (shs[i].getSheetId() === sid) return shs[i];
  return null;
}

// '품목 정보' 4행에서 고유번호 열 찾기 (열을 끼워 넣어 위치가 바뀌어도 찾음). 없으면 0
function mo_idCol_(info) {
  const head = info.getRange(4, 1, 1, info.getMaxColumns()).getDisplayValues()[0];
  for (let c = 0; c < head.length; c++) if (String(head[c]).trim() === MO_ID_HEAD) return c + 1;
  return 0;
}

function mo_newId_() { return 'i' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/** 마스터 시트 '변경 시' 트리거 */
function mo_onChange(e) {
  const ss = e && e.source;
  if (!ss) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return; // 5분 동기화가 다시 확인함
  try { mo_realign_(ss); } catch (err) { console.error('품목 순서 맞춤 실패: ' + err); }
  finally { lock.releaseLock(); }
}

/** 5분 동기화에서 파일마다 호출: 트리거 설치 + 순서 확인 */
function mo_check_(fileId) {
  const has = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'mo_onChange' && t.getTriggerSourceId() === fileId);
  if (!has) ScriptApp.newTrigger('mo_onChange').forSpreadsheet(fileId).onChange().create();
  mo_realign_(SpreadsheetApp.openById(fileId));
}

function mo_realign_(ss) {
  const info = mo_sheet_(ss, MO_INFO_SID), qty = mo_sheet_(ss, MO_QTY_SID);
  if (!info || !qty) return;
  const idCol = mo_idCol_(info);
  if (!idCol) return; // 옛 양식: 아무것도 하지 않음
  const n = MO_SLOTS;
  const names = info.getRange(MO_FIRST, 1, n, 1).getDisplayValues().map(r => String(r[0]).trim());
  const ids = info.getRange(MO_FIRST, idCol, n, 1).getDisplayValues().map(r => String(r[0]).trim());

  const props = PropertiesService.getScriptProperties();
  const key = 'mo_order_' + ss.getId();
  let stored = null;
  try { stored = JSON.parse(props.getProperty(key) || 'null'); } catch (err) { stored = null; }
  const sIds = (stored && stored.ids && stored.ids.length === n) ? stored.ids : null;

  // 고유번호 붙이기 · 겹친 번호 정리 (줄을 복사해 붙이면 번호가 겹침 → 원래 자리 쪽이 번호를 유지)
  let idChanged = false;
  const count = {};
  ids.forEach(id => { if (id) count[id] = (count[id] || 0) + 1; });
  const keep = {};
  for (let i = 0; i < n; i++) {
    const id = ids[i];
    if (id && count[id] > 1 && sIds && sIds[i] === id) keep[id] = i;
  }
  const seen = {};
  for (let i = 0; i < n; i++) {
    let id = ids[i];
    if (id && count[id] > 1) {
      const owner = (id in keep) ? keep[id] : -1;
      if ((owner >= 0 && owner !== i) || (owner < 0 && seen[id])) id = '';
    }
    if (!id && names[i]) id = mo_newId_();
    if (id !== ids[i]) { ids[i] = id; idChanged = true; }
    if (id) seen[id] = true;
  }
  if (idChanged) info.getRange(MO_FIRST, idCol, n, 1).setValues(ids.map(x => [x]));

  const save = () => props.setProperty(key, JSON.stringify({ ids: ids, names: names }));
  if (!sIds) { save(); return; } // 처음: 지금 상태를 기준으로 기억
  if (sIds.join('|') === ids.join('|')) {
    if (stored.names.join('|') !== names.join('|')) save();
    return;
  }

  // 새 자리 j 의 수량을 가져올 옛 자리
  // (새 품목이 원래 빈 자리에 생긴 경우엔 그 자리 열을 그대로 씀 — 이름보다 수량을 먼저 적었어도 지켜짐)
  const src = ids.map((id, j) => {
    const i = id ? sIds.indexOf(id) : -1;
    return i >= 0 ? i : (sIds[j] === '' ? j : -1);
  });
  const last = qty.getLastRow();
  if (last >= MO_QTY_ROW) {
    const nr = last - MO_QTY_ROW + 1;
    const rng = qty.getRange(MO_QTY_ROW, MO_QTY_COL, nr, n);
    const v = rng.getValues(), f = rng.getFormulas(), bg = rng.getBackgrounds();
    const stores = qty.getRange(MO_QTY_ROW, 1, nr, 1).getDisplayValues();
    const used = {};
    src.forEach(i => { if (i >= 0) used[i] = true; });

    // 자리를 잃은 열(지운 품목)에 수량이 있으면 백업
    const lost = [];
    for (let i = 0; i < n; i++) {
      if (used[i]) continue;
      for (let r = 0; r < nr; r++) {
        if (f[r][i] || v[r][i] === '' || v[r][i] === null) continue;
        lost.push([new Date(), ss.getName(), stored.names[i] || '(이름 없음)', stores[r][0], v[r][i]]);
      }
    }
    if (lost.length) mo_backup_(ss, lost);

    const out = [], outBg = [];
    for (let r = 0; r < nr; r++) {
      const fRow = f[r].some(x => x);
      const row = [], rowBg = [];
      for (let j = 0; j < n; j++) {
        if (fRow) { row.push(f[r][j] || v[r][j]); rowBg.push(bg[r][j]); continue; }
        const i = src[j];
        row.push(i >= 0 ? v[r][i] : '');
        rowBg.push(i >= 0 ? bg[r][i] : (String(bg[r][j]).toLowerCase() === '#fff3a6' ? '#ffffff' : bg[r][j]));
      }
      out.push(row); outBg.push(rowBg);
    }
    rng.setValues(out);
    rng.setBackgrounds(outBg);
  }
  save();
  mo_log_(ss, stored.names, names);
  try { ss.toast('품목 순서가 바뀌어 매장 수량도 같은 품목 열로 옮겼어요.', '최종 수량', 6); } catch (err) {}
}

function mo_backup_(ss, rows) {
  let sh = ss.getSheetByName(MO_BACKUP);
  if (!sh) {
    sh = ss.insertSheet(MO_BACKUP);
    sh.getRange(1, 1, 1, 5).setValues([['백업 시각', '파일', '품목', '매장', '수량']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function mo_log_(ss, before, after) {
  const log = mo_sheet_(ss, MO_LOG_SID);
  if (!log) return;
  const fmt = a => a.map(x => String(x || '-').replace(/\n/g, ' ')).join(' / ');
  log.insertRowsBefore(5, 1);
  log.getRange(5, 1, 1, 9).setValues([[new Date(), '자동 (품목 순서 맞춤)', '최종 수량', 'H~V', '',
    '품목 순서 변경 → 수량 열 이동', fmt(before), fmt(after), '매장 수량을 같은 품목 열로 옮김']])
    .setFontWeight('normal').setBackground(null).setFontColor('#000000');
  log.getRange(5, 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

/** 옛 양식 파일 준비: 고유번호 열 추가(숨김) + 최종 수량 4~6행을 자리 고정 수식으로. 한 번만 실행 */
function mo_prepareFile(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  const info = mo_sheet_(ss, MO_INFO_SID), qty = mo_sheet_(ss, MO_QTY_SID);
  if (!info || !qty) throw new Error('품목 정보 / 최종 수량 탭을 찾지 못했어요.');
  // 지금 '최종 수량' 품목명이 '품목 정보' 줄 순서와 같은지 먼저 확인 (다르면 수량이 어긋날 수 있어 멈춤)
  const want = info.getRange(MO_FIRST, 1, MO_SLOTS, 1).getDisplayValues().map(r => String(r[0]).trim());
  const have = qty.getRange(5, MO_QTY_COL, 1, MO_SLOTS).getDisplayValues()[0].map(x => String(x).trim());
  if (want.join('|') !== have.join('|')) throw new Error('최종 수량 품목명이 품목 정보 순서와 달라요: ' + have.join(', '));
  if (!mo_idCol_(info)) {
    info.insertColumnAfter(info.getMaxColumns());
    const c = info.getMaxColumns();
    info.getRange(4, c).setValue(MO_ID_HEAD);
    info.hideColumns(c);
  }
  const P = "'" + info.getName().replace(/'/g, "''") + "'!";
  const photo = [], name = [], size = [];
  for (let r = MO_FIRST; r < MO_FIRST + MO_SLOTS; r++) {
    const c = 'INDEX(' + P + '$C:$C,' + r + ')';
    photo.push('=IFERROR(LET(c,' + c + ',t,IFERROR(c&"",""),IF(t="",IF(ISBLANK(' + c + '),"",c),'
      + 'IF(REGEXMATCH(t,"drive\\.google\\.com"),IMAGE("https://drive.google.com/thumbnail?sz=w400&id="&REGEXEXTRACT(t,"(?:/d/|id=)([\\w-]+)")),'
      + 'IF(REGEXMATCH(t,"^https?://"),IMAGE(t),c)))),"")');
    name.push('=IF(INDEX(' + P + '$A:$A,' + r + ')="","",INDEX(' + P + '$A:$A,' + r + '))');
    size.push('=IF(INDEX(' + P + '$B:$B,' + r + ')="","",INDEX(' + P + '$B:$B,' + r + '))');
  }
  qty.getRange(4, MO_QTY_COL, 3, MO_SLOTS).setFormulas([photo, name, size]);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().deleteProperty('mo_order_' + fileId);
  mo_realign_(ss); // 고유번호 붙이고 지금 상태를 기준으로 기억
}
