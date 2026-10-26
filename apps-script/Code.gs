/**
 * 온담 홈페이지 문의 → 구글 시트·드라이브 저장
 *
 * 홈페이지 문의 폼이 이 웹 앱으로 내용을 보내면
 *  1) "온담 문의 접수" 시트에 한 줄 추가
 *  2) 첨부 파일은 "첨부파일" 폴더에 저장하고 링크를 시트에 기록
 *  3) 담당자 메일로 새 문의 알림 발송
 */

const SHEET_ID = '1WFZEh6FyDzo4h7ocsw1WQ91JrvvoJzoQxbQs7ltvQLk';   // 온담 문의 접수
const FOLDER_ID = '11d2PaxJ0kItYmOV8VEJJRWF63S2S2L68';             // 첨부파일
const NOTIFY_EMAIL = '';   // 알림 받을 메일. 비워 두면 스크립트 소유자 메일로 보냄
const MAX_BYTES = 10 * 1024 * 1024;

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.company || !d.name || !d.phone || !d.message || !d.agree) {
      return json({ ok: false, error: '필수 항목 누락' });
    }

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    let total = 0;
    const links = (d.files || []).map(f => {
      const bytes = Utilities.base64Decode(f.data);
      total += bytes.length;
      if (total > MAX_BYTES) throw new Error('첨부 파일 용량 초과');
      const blob = Utilities.newBlob(bytes, f.type, `${stamp} ${clean(d.company)} - ${clean(f.name)}`);
      return folder.createFile(blob).getUrl();
    });

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    sheet.appendRow([
      stamp, clean(d.kind), clean(d.company), clean(d.name), clean(d.phone),
      clean(d.email), clean(d.message), links.join('\n'), '동의', '신규'
    ]);

    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(to, `[온담 홈페이지] 새 ${d.kind} 문의 - ${d.company}`,
      `접수: ${stamp}\n유형: ${d.kind}\n회사: ${d.company}\n담당자: ${d.name}\n연락처: ${d.phone}\n이메일: ${d.email || '-'}\n\n${d.message}\n\n첨부: ${links.join('\n') || '없음'}\n\n시트: https://docs.google.com/spreadsheets/d/${SHEET_ID}`);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

// 브라우저로 주소를 열었을 때 연결 확인용
function doGet() {
  return json({ ok: true, service: '온담 문의 접수' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 시트 수식 주입 방지 + 길이 제한
function clean(v) {
  let s = String(v || '').slice(0, 5000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
