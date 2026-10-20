# Firestore 보안 규칙

- `firestore.rules`: Firebase 콘솔 → Firestore → 규칙 탭에 게시하는 규칙 원본 (젠틀몬스터·템버린즈 공통)
- `rules.test.mjs`: 규칙 테스트 (Firestore 에뮬레이터 사용, 실제 데이터에는 영향 없음)

테스트 실행: `npm install && npm test` (Java 필요. Google Cloud Shell에는 기본 설치되어 있음)

이 폴더의 파일은 규칙의 기록용이며, 게시는 Firebase 콘솔에서 직접 한다.
GitHub Pages(Jekyll)는 이 폴더도 배포하지만, 규칙 내용은 비밀이 아니므로 문제없다.
