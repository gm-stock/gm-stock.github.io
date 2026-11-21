// ── 젠틀몬스터 출고 내역 - 전체 상태 실시간 동기화 ──
// 출고 내역 시트 Apps Script의 'Code.gs' 파일 내용 (저장소에는 기록용 사본)
// 2026-11-21: API 키 대신 에이도스 계정 인증으로 읽음 (로그인 도입 후 키 방식은 거절됨)
// 2026-11-21: 내용이 바뀌었을 때만, 한 번에 쓰도록 변경 (예전에는 5분마다 시트를 비우고 한 줄씩 2~3분 동안 써서
//             그 사이 시트가 비어 보였음)
// ※ 맨 아래 setupTrigger 는 이 프로젝트의 모든 트리거(시즌 배송·마스터 시트 동기화 포함)를 지우므로 실행하지 말 것.
const SHEET_ID = "19l5pjxridh8Svtl0yUkpLimibDLS_81QcUDoLWf2fRc";
const SHEET_NAME = "출고 내역";
const FIREBASE_PROJECT = "gm-stock-management";

const HEADERS = ["날짜","시간","매장명","요청유형","컬렉션","상품명","옵션","수량","요청 사유","담당자","처리 상태","취소/반려 사유"];

const COLL_COLORS = {
  "VEGGIE":"#e8f5e9","BOUQUET":"#fce4ec","BOLD":"#e3f2fd",
  "MM":"#fff8e1","POCKET X BRATZ":"#f3e5f5","FALL SEASON":"#fbe9e7",
  "JEWELRY":"#e8eaf6","CIRCUIT X DISNEY":"#e0f7fa"
};

const STATUS_STYLES = {
  "검토중":   {bg:"#ede9fe", color:"#5b21b6"},
  "승인":     {bg:"#dbeafe", color:"#1e40af"},
  "준비중":   {bg:"#fef3c7", color:"#92400e"},
  "출고대기": {bg:"#ffedd5", color:"#9a3412"},
  "출고완료": {bg:"#d1fae5", color:"#065f46"},
  "반려":     {bg:"#fee2e2", color:"#991b1b"},
  "취소":     {bg:"#f3f4f6", color:"#374151"},
};

// ── 메인 동기화 (5분마다 자동 실행) ──
function syncFromFirebase() {
  // 앞 실행이 아직 안 끝났으면 이번 차례는 건너뜀
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const orders = fetchFirebaseOrders();
    if (!orders || orders.length === 0) {
      Logger.log("주문 없음 (또는 읽기 실패) — 시트는 그대로 둠");
      return;
    }

    // 1) 화면에 쓸 내용을 먼저 전부 만든다
    const rows = [], merges = [], firstOf = [];
    orders.forEach(order => {
      const items = order.items && order.items.length ? order.items : [null];
      const start = rows.length + 2;
      items.forEach((item, idx) => { rows.push(rowOf_(order, item, idx)); firstOf.push(idx === 0 ? order : null); });
      if (items.length > 1) merges.push([start, items.length]);
    });

    // 2) 지난번과 같으면 아무것도 하지 않음
    const props = PropertiesService.getScriptProperties();
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rows)));
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (sheet && props.getProperty("orders_hash") === hash && sheet.getLastRow() === rows.length + 1) {
      Logger.log("변경 없음: " + orders.length + "건");
      return;
    }
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // 3) 한 번에 쓴다
    const n = rows.length, W = HEADERS.length;
    const bgs = [], fcs = [], fws = [], has = [];
    rows.forEach((r, i) => {
      const rowNum = i + 2, order = firstOf[i];
      const base = rowNum % 2 === 0 ? "#f8fafc" : "#ffffff";
      const bg = new Array(W).fill(base), fc = new Array(W).fill("#000000"), fw = new Array(W).fill("normal"), ha = new Array(W).fill("left");
      ha[0] = "center"; ha[1] = "center"; ha[7] = "center"; fw[7] = "bold";
      const coll = r[4];
      if (COLL_COLORS[coll]) { bg[4] = COLL_COLORS[coll]; fw[4] = "bold"; }
      if (order) {
        const st = STATUS_STYLES[order.status] || {bg:"#f3f4f6", color:"#374151"};
        bg[10] = st.bg; fc[10] = st.color; fw[10] = "bold"; ha[10] = "center";
        if (order.orderType === "파손") { fc[3] = "#dc2626"; fw[3] = "bold"; ha[3] = "center"; }
        else if (order.orderType === "부족") { fc[3] = "#d97706"; fw[3] = "bold"; ha[3] = "center"; }
        if (order.cancelReason || order.rejectReason) fc[11] = "#dc2626";
      }
      bgs.push(bg); fcs.push(fc); fws.push(fw); has.push(ha);
    });

    sheet.getDataRange().breakApart();
    sheet.clearContents();
    sheet.clearFormats();
    setupHeader(sheet);
    if (sheet.getMaxRows() < n + 1) sheet.insertRowsAfter(sheet.getMaxRows(), n + 1 - sheet.getMaxRows());
    const body = sheet.getRange(2, 1, n, W);
    body.setValues(rows);
    body.setBackgrounds(bgs).setFontColors(fcs).setFontWeights(fws).setHorizontalAlignments(has)
        .setFontSize(10).setVerticalAlignment("middle")
        .setBorder(null, null, true, null, null, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(2, 12, n, 1).setFontSize(9);
    // 여러 품목 주문은 날짜·시간·매장명·요청유형·사유·담당자·상태·취소사유 칸 병합
    merges.forEach(([start, cnt]) => {
      [1,2,3,4,9,10,11,12].forEach(col => { try { sheet.getRange(start, col, cnt, 1).merge(); } catch (e) {} });
    });
    sheet.autoResizeColumns(1, W);
    for (let col = 1; col <= W; col++) sheet.setColumnWidth(col, sheet.getColumnWidth(col) + 30);
    sheet.autoResizeRows(2, n);
    SpreadsheetApp.flush();
    props.setProperty("orders_hash", hash);
    Logger.log(`동기화 완료: 총 ${orders.length}건`);
  } catch (e) {
    Logger.log("오류: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function rowOf_(order, item, idx) {
  const isFirst = idx === 0;
  const cancelNote = order.cancelReason || order.rejectReason || "";
  return [
    isFirst ? order.date : "",
    isFirst ? order.time : "",
    isFirst ? order.storeName : "",
    isFirst ? (order.orderType || "") : "",
    item ? item.collection : "",
    item ? item.itemName : "",
    item ? (item.optName || "-") : "",
    item ? item.qty : "",
    isFirst ? order.reason : "",
    isFirst ? order.handler : "",
    isFirst ? order.status : "",
    isFirst ? cancelNote : ""
  ];
}

// ── Firebase REST API로 orders 가져오기 (에이도스 계정 인증) ──
function fetchFirebaseOrders() {
  let allDocs = [];
  let pageToken = null;
  let guard = 0;
  do {
    let url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT}/databases/(default)/documents/orders?pageSize=300`;
    if (pageToken) url += "&pageToken=" + encodeURIComponent(pageToken);
    const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const json = JSON.parse(res.getContentText());
    if (json.error) { Logger.log("Firebase 오류: " + JSON.stringify(json.error)); return []; }
    if (json.documents) allDocs = allDocs.concat(json.documents);
    pageToken = json.nextPageToken || null;
    guard++;
  } while (pageToken && guard < 50);

  Logger.log("Firebase에서 읽은 출고요청 수: " + allDocs.length);
  if (!allDocs.length) return [];

  return allDocs.map(doc => {
    const f = doc.fields || {};
    const id = doc.name.split("/").pop();
    let dateStr = getStr(f.date);
    let timeStr = getStr(f.time);
    if (!dateStr && f.createdAt && f.createdAt.timestampValue) {
      const d = new Date(f.createdAt.timestampValue);
      dateStr = Utilities.formatDate(d, "Asia/Seoul", "yyyy/MM/dd");
      timeStr = Utilities.formatDate(d, "Asia/Seoul", "HH:mm");
    }
    return {
      id,
      date: dateStr,
      time: timeStr,
      storeName: getStr(f.storeName),
      status: getStr(f.status) || "검토중",
      orderType: getStr(f.orderType),
      reason: getStr(f.reason),
      handler: getStr(f.handler),
      cancelReason: getStr(f.cancelReason),
      rejectReason: getStr(f.rejectReason),
      items: getArr(f.items)
    };
  }).sort((a,b) => (b.date + b.time).localeCompare(a.date + a.time));
}

function getStr(field) {
  if (!field) return "";
  return field.stringValue || field.integerValue || "";
}

function getArr(field) {
  if (!field || !field.arrayValue || !field.arrayValue.values) return [];
  return field.arrayValue.values.map(v => {
    const m = v.mapValue && v.mapValue.fields || {};
    return {
      collection: getStr(m.collection),
      itemName: getStr(m.itemName),
      optName: getStr(m.optName) || "-",
      qty: parseInt(m.qty && (m.qty.integerValue || m.qty.doubleValue) || 0),
      unit: getStr(m.unit) || "개"
    };
  });
}

// ── 헤더 설정 ──
function setupHeader(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  const r = sheet.getRange(1, 1, 1, HEADERS.length);
  r.setBackground("#0f172a")
   .setFontColor("#ffffff")
   .setFontWeight("bold")
   .setFontSize(10)
   .setHorizontalAlignment("center")
   .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 34);
  sheet.setFrozenRows(1);
}

// ── 트리거 설정 ──
// ⚠ 실행 금지: 이 프로젝트의 모든 트리거(ss_syncSeasonSheets, ms_syncMasterSheets 포함)를 지운다.
//   출고 내역 트리거만 다시 만들어야 할 때는 setupOrderTrigger 를 실행할 것.
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncFromFirebase").timeBased().everyMinutes(5).create();
  Logger.log("5분 트리거 설정 완료");
}
function setupOrderTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "syncFromFirebase");
  if (!exists) ScriptApp.newTrigger("syncFromFirebase").timeBased().everyMinutes(5).create();
}
