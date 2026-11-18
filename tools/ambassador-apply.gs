/**
 * QUOC 앰버서더 지원서 접수 (Google Apps Script 웹앱)
 *
 * 설치
 *  1. 구글 드라이브에서 새 스프레드시트를 만들고, 확장 프로그램 → Apps Script를 엽니다.
 *  2. 이 파일 내용을 붙여 넣고 저장합니다.
 *  3. 배포 → 새 배포 → 유형: 웹 앱, 실행: 나, 액세스 권한: 모든 사용자 → 배포.
 *  4. 나온 웹 앱 주소(https://script.google.com/macros/s/.../exec)를
 *     ambassador/index.html 의 <form id="applyForm" data-endpoint=""> 따옴표 안에 넣습니다.
 *
 * 주소를 넣기 전에는 지원서가 help@rubi.kr 메일로 열립니다.
 */
const SHEET_NAME = '2027 지원서';
const NOTIFY_TO = 'help@rubi.kr';

function doPost(e) {
  const p = e.parameter || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const fields = Object.keys(p);

  let headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  if (!headers.length) {
    headers = ['접수 시각'].concat(fields);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  fields.forEach(function (f) {
    if (headers.indexOf(f) === -1) {
      headers.push(f);
      sheet.getRange(1, headers.length).setValue(f);
    }
  });

  const row = headers.map(function (h) { return h === '접수 시각' ? new Date() : (p[h] || ''); });
  sheet.appendRow(row);

  const body = fields.map(function (f) { return f + ': ' + (p[f] || '-'); }).join('\n');
  MailApp.sendEmail(NOTIFY_TO, '[2027 QUOC 앰버서더 지원] ' + (p['이름'] || ''), body + '\n\n' + ss.getUrl());

  return ContentService.createTextOutput('ok');
}
