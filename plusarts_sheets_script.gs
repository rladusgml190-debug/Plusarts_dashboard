// ─────────────────────────────────────────────────────────────
//  PLUSARTS 신입생 상담지 → 웹 연동 Apps Script
//  이 파일을 Google Apps Script에 붙여넣고 웹 앱으로 배포하세요.
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var rows = [];

  // 행 2부터 (1행은 헤더)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue; // 학생 이름(B열) 없으면 스킵

    // 휴대폰 뒷 4자리 추출
    var phoneRaw = String(row[2] || '').replace(/[^0-9]/g, '');
    var phone4 = phoneRaw.length >= 4 ? phoneRaw.slice(-4) : '';

    rows.push({
      timestamp:   String(row[0]  || ''),   // A: 타임스탬프
      name:        String(row[1]  || ''),   // B: 학생 성함
      phoneFull:   String(row[2]  || ''),   // C: 학생 연락처
      phone4:      phone4,                  // C에서 뒷 4자리 추출
      email:       String(row[3]  || ''),   // D: 학생 이메일
      parent:      String(row[4]  || ''),   // E: 학부모 성함/연락처
      grade:       String(row[5]  || ''),   // F: 현재 학년
      country:     String(row[6]  || ''),   // G: 희망 지원 국가
      major:       String(row[7]  || ''),   // H: 관심 전공
      artExp:      String(row[8]  || ''),   // I: 미술 학습 경험
      portfolio:   String(row[9]  || ''),   // J: 포트폴리오 유무
      personality: String(row[10] || ''),   // K: 학생 성향/특이사항
      message:     String(row[11] || ''),   // L: 학원에 전하고 싶은 말
      callTime:    String(row[12] || '')    // M: 통화 가능 시간대
    });
  }

  var output = JSON.stringify({
    rows: rows,
    total: rows.length,
    updated: new Date().toISOString()
  });

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
