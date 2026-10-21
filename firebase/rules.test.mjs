// Firestore 보안 규칙 테스트 — 재고관리 앱(젠틀몬스터·템버린즈 공통)
// 실행: npx firebase emulators:exec --only firestore "node --test rules.test.mjs"
import { readFileSync } from "node:fs";
import { test, before, after, beforeEach } from "node:test";
import {
  initializeTestEnvironment, assertSucceeds, assertFails,
} from "@firebase/rules-unit-testing";
import {
  doc, getDoc, getDocs, setDoc, addDoc, deleteDoc, collection, serverTimestamp,
} from "firebase/firestore";

let env;
const ITEM_PLAIN = { id: "i1", name: "SET1 PLATE", collection: "MM2", year: "2026", quantity: 10, unit: "개", _order: 1 };
const ITEM_OPTS = {
  id: "i2", name: "CAUTION PAPER", collection: "MM2", year: "2026", quantity: 0, unit: "개", _order: 2,
  options: [{ name: "한국어", quantity: 5 }, { name: "영어", quantity: 3 }],
};
const ORDER_REVIEW = {
  storeName: "TOKYO GINZA _ JP", status: "검토중", orderType: "신규", reason: "[신규]",
  items: [{ itemId: "i1", itemName: "SET1 PLATE", optName: null, qty: 2, unit: "개" }],
  date: "2026-10-20", time: "10:00",
};
const ORDER_APPROVED = { ...ORDER_REVIEW, status: "승인", handler: "에이도스 스튜디오" };

before(async () => {
  env = await initializeTestEnvironment({
    projectId: "demo-stock",
    firestore: { rules: readFileSync("firestore.rules", "utf8"), host: "127.0.0.1", port: 8085 },
  });
});
after(async () => { await env.cleanup(); });
beforeEach(async () => {
  await env.clearFirestore();
  await env.withSecurityRulesDisabled(async (ctx) => {
    const db = ctx.firestore();
    await setDoc(doc(db, "admins", "owner@eidos.kr"), { name: "신", role: "owner" });
    await setDoc(doc(db, "admins", "mgr@gentlemonster.com"), { name: "담당자", role: "manager" });
    await setDoc(doc(db, "admins", "norole@eidos.kr"), { name: "역할없음" });
    await setDoc(doc(db, "admins", "typo@eidos.kr"), { name: "오타", role: "Owner " });
    await setDoc(doc(db, "stores", "gentlemonster.stock@eidosstudio.co.kr"), { name: "젠틀몬스터 매장 공용" });
    await setDoc(doc(db, "items", "i1"), ITEM_PLAIN);
    await setDoc(doc(db, "items", "i2"), ITEM_OPTS);
    await setDoc(doc(db, "orders", "o1"), ORDER_REVIEW);
    await setDoc(doc(db, "orders", "o2"), ORDER_APPROVED);
    await setDoc(doc(db, "stockLogs", "l1"), { itemId: "i1", from: 10, to: 9 });
  });
});

// 접속 주체
const anon = () => env.unauthenticatedContext().firestore();                 // 외부인 (로그인 없음)
const user = (email) => env.authenticatedContext(email.replace(/\W/g, "_"), { email }).firestore();
const stranger = () => user("stranger@x.com");                               // 로그인했지만 명단에 없음
const manager = () => user("mgr@gentlemonster.com");
const norole = () => user("norole@eidos.kr");
const typo = () => user("typo@eidos.kr");
const owner = () => user("owner@eidos.kr");
const store = () => user("gentlemonster.stock@eidosstudio.co.kr");            // 매장 공용 계정
const ownerUpper = () => user("Owner@Eidos.KR");                              // 대문자 이메일

// ── 매장 계정: 매장 화면이 실제로 하는 동작은 모두 통과해야 한다 ──
test("매장: 재고·출고요청 보기", async () => {
  await assertSucceeds(getDocs(collection(store(), "items")));
  await assertSucceeds(getDocs(collection(store(), "orders")));
});
test("매장: 발주 시 재고 차감 (문서 전체 다시 저장, 옵션 없는 품목)", async () => {
  await assertSucceeds(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, quantity: 8 }));
});
test("매장: 발주 시 재고 차감 (옵션 수량)", async () => {
  const options = [{ name: "한국어", quantity: 3 }, { name: "영어", quantity: 3 }];
  await assertSucceeds(setDoc(doc(store(), "items", "i2"), { ...ITEM_OPTS, options }));
});
test("매장: options가 비어 있는(null) 품목의 수량 차감", async () => {
  await env.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), "items", "i3"), { ...ITEM_PLAIN, id: "i3", options: null });
  });
  await assertSucceeds(setDoc(doc(store(), "items", "i3"), { ...ITEM_PLAIN, id: "i3", options: null, quantity: 4 }));
});
test("매장: 빈 options 목록 품목의 수량 차감", async () => {
  await env.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), "items", "i4"), { ...ITEM_PLAIN, id: "i4", options: [] });
  });
  await assertSucceeds(setDoc(doc(store(), "items", "i4"), { ...ITEM_PLAIN, id: "i4", options: [], quantity: 4 }));
});
test("매장 계정(세부 규칙): null인 options를 목록으로 바꾸기 차단", async () => {
  await env.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), "items", "i3"), { ...ITEM_PLAIN, id: "i3", options: null });
  });
  await assertFails(setDoc(doc(store(), "items", "i3"), { ...ITEM_PLAIN, id: "i3", options: [{ name: "x", quantity: 1 }] }));
});
test("매장: 차감 결과가 음수여도 저장 (앱 동작 유지)", async () => {
  await assertSucceeds(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, quantity: -2 }));
});
test("매장: 출고요청 만들기 (검토중)", async () => {
  await assertSucceeds(addDoc(collection(store(), "orders"), { ...ORDER_REVIEW, createdAt: serverTimestamp() }));
});
test("매장: 검토중 요청의 수량 수정 (문서 전체 + id 칸)", async () => {
  const items = [{ ...ORDER_REVIEW.items[0], qty: 5 }];
  await assertSucceeds(setDoc(doc(store(), "orders", "o1"), { ...ORDER_REVIEW, items, id: "o1" }, { merge: false }));
});
test("매장: 검토중 요청의 매장명 수정", async () => {
  await assertSucceeds(setDoc(doc(store(), "orders", "o1"), { ...ORDER_REVIEW, storeName: "OSAKA PARCO _ JP", id: "o1" }));
});

// ── 외부인(로그인 없음): 막혀야 하는 것 ──
test("매장 계정(세부 규칙): 품목 이름 변경 차단", async () => {
  await assertFails(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, name: "HACKED" }));
});
test("매장 계정(세부 규칙): 품목에 새 칸 추가 차단", async () => {
  await assertFails(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, imageUrl: "x" }));
});
test("매장 계정(세부 규칙): 품목 칸 삭제 차단", async () => {
  const { unit, ...rest } = ITEM_PLAIN;
  await assertFails(setDoc(doc(store(), "items", "i1"), rest));
});
test("매장 계정(세부 규칙): 옵션 개수 변경 차단", async () => {
  await assertFails(setDoc(doc(store(), "items", "i2"), { ...ITEM_OPTS, options: [ITEM_OPTS.options[0]] }));
});
test("매장 계정(세부 규칙): 수량을 글자로 바꾸기 차단", async () => {
  await assertFails(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, quantity: "999" }));
});
test("외부: 품목 추가·삭제 차단", async () => {
  await assertFails(setDoc(doc(anon(), "items", "new"), { ...ITEM_PLAIN, id: "new" }));
  await assertFails(deleteDoc(doc(anon(), "items", "i1")));
});
test("매장 계정(세부 규칙): 승인 상태로 출고요청 만들기 차단", async () => {
  await assertFails(addDoc(collection(store(), "orders"), { ...ORDER_REVIEW, status: "승인" }));
});
test("매장 계정(세부 규칙): 상태 없는 출고요청 만들기 차단", async () => {
  const { status, ...rest } = ORDER_REVIEW;
  await assertFails(addDoc(collection(store(), "orders"), rest));
});
test("매장 계정(세부 규칙): 출고요청 상태 변경(승인) 차단", async () => {
  await assertFails(setDoc(doc(store(), "orders", "o1"), { ...ORDER_REVIEW, status: "승인" }));
});
test("매장 계정(세부 규칙): 승인된 요청 수정 차단", async () => {
  const items = [{ ...ORDER_APPROVED.items[0], qty: 50 }];
  await assertFails(setDoc(doc(store(), "orders", "o2"), { ...ORDER_APPROVED, items }));
});
test("매장 계정(세부 규칙): 검토중 요청에 다른 칸 추가 차단 (예: handler)", async () => {
  await assertFails(setDoc(doc(store(), "orders", "o1"), { ...ORDER_REVIEW, handler: "x" }));
});
test("외부: 출고요청 삭제 차단", async () => {
  await assertFails(deleteDoc(doc(anon(), "orders", "o1")));
});
test("외부: 재고 기록 읽기·쓰기 차단", async () => {
  await assertFails(getDocs(collection(anon(), "stockLogs")));
  await assertFails(addDoc(collection(anon(), "stockLogs"), { a: 1 }));
});
test("외부: 관리자 명단 읽기·쓰기 차단", async () => {
  await assertFails(getDoc(doc(anon(), "admins", "owner@eidos.kr")));
  await assertFails(getDocs(collection(anon(), "admins")));
  await assertFails(setDoc(doc(anon(), "admins", "hacker@x.com"), { role: "owner" }));
});
test("외부: 다른 컬렉션 쓰기 차단", async () => {
  await assertFails(setDoc(doc(anon(), "anything", "x"), { a: 1 }));
});

test("외부: 재고·출고요청 보기 차단", async () => {
  await assertFails(getDocs(collection(anon(), "items")));
  await assertFails(getDoc(doc(anon(), "items", "i1")));
  await assertFails(getDocs(collection(anon(), "orders")));
});
test("외부: 재고 수량 변경 차단 (수량만 바꿔도)", async () => {
  await assertFails(setDoc(doc(anon(), "items", "i1"), { ...ITEM_PLAIN, quantity: 9999 }));
  const options = [{ name: "한국어", quantity: 999 }, { name: "영어", quantity: 999 }];
  await assertFails(setDoc(doc(anon(), "items", "i2"), { ...ITEM_OPTS, options }));
});
test("외부: 검토중 출고요청 만들기·수정 차단", async () => {
  await assertFails(addDoc(collection(anon(), "orders"), { ...ORDER_REVIEW, createdAt: serverTimestamp() }));
  await assertFails(setDoc(doc(anon(), "orders", "o1"), { ...ORDER_REVIEW, items: [{ ...ORDER_REVIEW.items[0], qty: 99 }] }));
});
test("외부: 매장 명단 읽기·쓰기 차단", async () => {
  await assertFails(getDoc(doc(anon(), "stores", "gentlemonster.stock@eidosstudio.co.kr")));
  await assertFails(setDoc(doc(anon(), "stores", "hacker@x.com"), { name: "x" }));
});

// ── 매장 계정: 막혀야 하는 것 ──
test("매장 계정: 상태 변경·품목 수정·삭제 차단", async () => {
  await assertFails(setDoc(doc(store(), "orders", "o1"), { ...ORDER_REVIEW, status: "승인" }));
  await assertFails(setDoc(doc(store(), "orders", "o2"), { ...ORDER_APPROVED, items: [{ ...ORDER_APPROVED.items[0], qty: 9 }] }));
  await assertFails(setDoc(doc(store(), "items", "i1"), { ...ITEM_PLAIN, name: "x" }));
  await assertFails(setDoc(doc(store(), "items", "new"), { ...ITEM_PLAIN, id: "new" }));
  await assertFails(deleteDoc(doc(store(), "items", "i1")));
  await assertFails(deleteDoc(doc(store(), "orders", "o1")));
  await assertFails(addDoc(collection(store(), "orders"), { ...ORDER_REVIEW, status: "승인" }));
});
test("매장 계정: 관리자 데이터 차단, 본인 매장 문서만 읽기", async () => {
  await assertFails(getDocs(collection(store(), "stockLogs")));
  await assertFails(getDoc(doc(store(), "admins", "owner@eidos.kr")));
  await assertSucceeds(getDoc(doc(store(), "stores", "gentlemonster.stock@eidosstudio.co.kr")));
  await assertFails(getDocs(collection(store(), "stores")));
  await assertFails(setDoc(doc(store(), "stores", "gentlemonster.stock@eidosstudio.co.kr"), { name: "x", role: "owner" }));
  await assertFails(setDoc(doc(store(), "admins", "gentlemonster.stock@eidosstudio.co.kr"), { role: "owner" }));
});
test("관리자도 매장 화면 사용 가능 (보기·발주)", async () => {
  await assertSucceeds(getDocs(collection(manager(), "items")));
  await assertSucceeds(addDoc(collection(manager(), "orders"), { ...ORDER_REVIEW, createdAt: serverTimestamp() }));
  await assertSucceeds(getDocs(collection(owner(), "orders")));
});

// ── 로그인했지만 명단에 없는 계정 = 외부와 같음 ──
test("명단 밖 계정: 보기·수량 변경·발주 차단", async () => {
  await assertFails(getDocs(collection(stranger(), "items")));
  await assertFails(setDoc(doc(stranger(), "items", "i1"), { ...ITEM_PLAIN, quantity: 1 }));
  await assertFails(addDoc(collection(stranger(), "orders"), { ...ORDER_REVIEW }));
});
test("명단 밖 계정: 상태 변경·품목 수정·삭제 차단", async () => {
  await assertFails(setDoc(doc(stranger(), "orders", "o1"), { ...ORDER_REVIEW, status: "승인" }));
  await assertFails(setDoc(doc(stranger(), "items", "i1"), { ...ITEM_PLAIN, name: "x" }));
  await assertFails(deleteDoc(doc(stranger(), "orders", "o1")));
  await assertFails(getDoc(doc(stranger(), "admins", "owner@eidos.kr")));
});
test("명단 밖 계정: 자기 이메일로 명단 문서 만들기 차단", async () => {
  await assertFails(setDoc(doc(stranger(), "admins", "stranger@x.com"), { role: "owner" }));
});

// ── 담당자: 출고요청 처리 가능, 품목·삭제 불가 ──
for (const [label, ctx] of [["담당자", manager], ["role 없음", norole], ["role 오타", typo]]) {
  test(`${label}: 승인 처리 (상태·처리자·처리시각)`, async () => {
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o1"),
      { ...ORDER_REVIEW, id: "o1", status: "승인", handler: "에이도스 스튜디오", handledAt: "2026-10-20T10:00:00Z" }, { merge: false }));
  });
  test(`${label}: 반려 + 재고 복원`, async () => {
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o1"),
      { ...ORDER_REVIEW, status: "반려", rejectReason: "중복", handler: "에이도스 스튜디오" }));
    await assertSucceeds(setDoc(doc(ctx(), "items", "i1"), { ...ITEM_PLAIN, quantity: 12 }));
  });
  test(`${label}: 출고완료·취소·매장명 수정`, async () => {
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o2"), { ...ORDER_APPROVED, status: "출고완료" }));
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o2"), { ...ORDER_APPROVED, status: "취소", cancelReason: "x" }));
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o1"), { ...ORDER_REVIEW, storeName: "PARIS _ FR" }));
  });
  test(`${label}: 품목 이름·사진 변경, 추가, 삭제 차단`, async () => {
    await assertFails(setDoc(doc(ctx(), "items", "i1"), { ...ITEM_PLAIN, name: "x" }));
    await assertFails(setDoc(doc(ctx(), "items", "i1"), { ...ITEM_PLAIN, imageUrl: "data:x" }));
    await assertFails(setDoc(doc(ctx(), "items", "new"), { ...ITEM_PLAIN, id: "new" }));
    await assertFails(deleteDoc(doc(ctx(), "items", "i1")));
  });
  test(`${label}: 출고요청 삭제·재고 기록 쓰기 차단, 기록 읽기 가능`, async () => {
    await assertFails(deleteDoc(doc(ctx(), "orders", "o1")));
    await assertFails(addDoc(collection(ctx(), "stockLogs"), { a: 1 }));
    await assertSucceeds(getDocs(collection(ctx(), "stockLogs")));
  });
}
test("담당자: 본인 명단 문서만 읽기, 명단 수정·목록 차단", async () => {
  await assertSucceeds(getDoc(doc(manager(), "admins", "mgr@gentlemonster.com")));
  await assertFails(getDoc(doc(manager(), "admins", "owner@eidos.kr")));
  await assertFails(getDocs(collection(manager(), "admins")));
  await assertFails(setDoc(doc(manager(), "admins", "mgr@gentlemonster.com"), { name: "담당자", role: "owner" }));
});

// ── 최고관리자: 전부 가능 (명단 수정만 콘솔에서) ──
for (const [label, ctx] of [["최고관리자", owner], ["최고관리자(대문자 이메일)", ownerUpper]]) {
  test(`${label}: 품목 추가·이름·사진·삭제`, async () => {
    await assertSucceeds(setDoc(doc(ctx(), "items", "new"), { ...ITEM_PLAIN, id: "new", name: "NEW" }));
    await assertSucceeds(setDoc(doc(ctx(), "items", "i1"), { ...ITEM_PLAIN, name: "RENAMED", imageUrl: "data:x" }));
    await assertSucceeds(deleteDoc(doc(ctx(), "items", "new")));
  });
  test(`${label}: 출고요청 처리·삭제, 엑셀 일괄 발주(승인 상태 생성)`, async () => {
    await assertSucceeds(setDoc(doc(ctx(), "orders", "o1"), { ...ORDER_REVIEW, status: "승인" }));
    await assertSucceeds(addDoc(collection(ctx(), "orders"), { ...ORDER_REVIEW, status: "승인", createdAt: serverTimestamp() }));
    await assertSucceeds(deleteDoc(doc(ctx(), "orders", "o2")));
  });
  test(`${label}: 재고 기록 쓰기`, async () => {
    await assertSucceeds(addDoc(collection(ctx(), "stockLogs"), { itemId: "i1", from: 10, to: 3 }));
  });
}
test("포장재(packaging): 외부 차단, 관리자 읽기, 최고관리자만 쓰기", async () => {
  await env.withSecurityRulesDisabled(async (ctx) => {
    await setDoc(doc(ctx.firestore(), "packaging", "p1"), { id: "p1", name: "에어캡 1m", quantity: 3, options: null });
  });
  await assertFails(getDoc(doc(anon(), "packaging", "p1")));
  await assertSucceeds(getDoc(doc(manager(), "packaging", "p1")));
  await assertFails(setDoc(doc(manager(), "packaging", "p1"), { id: "p1", name: "에어캡 1m", quantity: 1, options: null }));
  await assertSucceeds(setDoc(doc(owner(), "packaging", "p1"), { id: "p1", name: "에어캡 1m", quantity: 1, options: null }));
});
test("최고관리자: 관리자 명단은 보기만, 수정은 앱에서 불가 (콘솔 전용)", async () => {
  await assertSucceeds(getDoc(doc(owner(), "admins", "owner@eidos.kr")));
  await assertSucceeds(getDocs(collection(owner(), "admins")));
  await assertFails(setDoc(doc(owner(), "admins", "new@x.com"), { role: "manager" }));
});
test("최고관리자: 매장 계정 명단 보기·추가·삭제 (계정 관리 화면)", async () => {
  await assertSucceeds(getDocs(collection(owner(), "stores")));
  await assertSucceeds(setDoc(doc(owner(), "stores", "new.stock@eidosstudio.co.kr"), { name: "새 공용" }));
  await assertFails(setDoc(doc(owner(), "stores", "Upper@eidosstudio.co.kr"), { name: "대문자" }));
  await assertFails(setDoc(doc(owner(), "stores", "new.stock@eidosstudio.co.kr"), { name: "덮어쓰기" }));
  await assertSucceeds(deleteDoc(doc(owner(), "stores", "new.stock@eidosstudio.co.kr")));
});
test("발주자 이름: 매장은 저장만, 관리자만 읽기", async () => {
  await assertSucceeds(setDoc(doc(store(), "orderContacts", "o-new"), { name: "홍길동" }));
  await assertFails(setDoc(doc(store(), "orderContacts", "o-new"), { name: "덮어쓰기" }));
  await assertFails(getDoc(doc(store(), "orderContacts", "o-new")));
  await assertFails(setDoc(doc(store(), "orderContacts", "o-x"), { name: "홍길동", extra: 1 }));
  await assertFails(setDoc(doc(anon(), "orderContacts", "o-y"), { name: "외부인" }));
  await assertSucceeds(getDoc(doc(manager(), "orderContacts", "o-new")));
  await assertSucceeds(getDoc(doc(owner(), "orderContacts", "o-new")));
});
test("담당자: 매장 계정 명단 보기·추가·삭제 불가", async () => {
  await assertFails(getDocs(collection(manager(), "stores")));
  await assertFails(setDoc(doc(manager(), "stores", "x.stock@eidosstudio.co.kr"), { name: "x" }));
  await assertFails(deleteDoc(doc(manager(), "stores", "gentlemonster.stock@eidosstudio.co.kr")));
});
