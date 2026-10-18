/*
 * 참석 의사 → 구글 시트 저장용 Apps Script
 *
 * 설정 방법
 * 1. 구글 드라이브에서 새 구글 시트를 만듭니다.
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script]를 열고 이 파일 내용을 전부 붙여넣고 저장합니다.
 * 3. 오른쪽 위 [배포] → [새 배포] → 유형 [웹 앱]
 *    - 다음 사용자 인증 정보로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] 후 권한 승인
 * 4. 나온 웹 앱 URL(https://script.google.com/macros/s/.../exec)을
 *    index.html CONFIG의 rsvpUrl 에 붙여넣습니다.
 *
 * 새 응답이 올 때 메일로도 받고 싶으면 아래 NOTIFY_EMAIL 에 주소를 넣으세요.
 */
const NOTIFY_EMAIL = "";

function doPost(e) {
  const p = e.parameter;
  // 시트 수식으로 해석되지 않도록 = + - @ 로 시작하는 값 앞에 ' 를 붙임
  const clean = v => {
    const s = String(v || "").slice(0, 100);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };
  const row = [new Date(), clean(p.side), clean(p.attend), clean(p.name), clean(p.phone)];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(["접수 시각", "하객 구분", "참석 여부", "성함", "휴대폰 뒤 4자리"]);
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    MailApp.sendEmail(NOTIFY_EMAIL, `[청첩장] 참석 의사 - ${row[3]} (${row[2]})`,
      `하객 구분: ${row[1]}\n참석 여부: ${row[2]}\n성함: ${row[3]}\n휴대폰 뒤 4자리: ${row[4]}`);
  }
  return ContentService.createTextOutput("ok");
}
