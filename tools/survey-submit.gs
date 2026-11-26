/**
 * QUOC 라이더 서베이 2026 (My QUOC 2026) 접수 (Google Apps Script 웹앱)
 * 페이지: quoc.rubiworkshop.com/survey/
 *
 * 하는 일
 *  - 사진: 구글 드라이브 "My QUOC 2026 서베이" 폴더 안에 응답자별 폴더를 만들어 원본(또는 긴 변 3,000px)으로 저장
 *  - 답변: 이 스프레드시트의 "2026 서베이" 시트에 한 줄씩 추가 (사진 폴더 링크 포함)
 *  - 알림: NOTIFY_TO 로 메일 발송
 *
 * 설치
 *  1. 구글 드라이브에서 새 스프레드시트를 만들고, 확장 프로그램 → Apps Script를 엽니다.
 *  2. 이 파일 내용을 붙여 넣고 저장합니다.
 *  3. 편집기에서 함수 setup 을 한 번 실행해 드라이브 · 시트 · 메일 권한을 승인합니다.
 *  4. 배포 → 새 배포 → 유형: 웹 앱, 실행: 나, 액세스 권한: 모든 사용자 → 배포.
 *  5. 나온 웹 앱 주소(https://script.google.com/macros/s/.../exec)를
 *     survey/index.html 의 <form id="surveyForm" data-endpoint=""> 따옴표 안에 넣습니다.
 *
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 수정 → 버전: 새 버전 으로 다시 배포해야 반영됩니다 (주소는 그대로).
 * 주소를 넣기 전에는 설문이 help@rubi.kr 메일로 열립니다.
 */
const SHEET_NAME = '2026 서베이';
const FOLDER_NAME = 'My QUOC 2026 서베이';
const NOTIFY_TO = 'help@rubi.kr';
const MAX_B64 = 45 * 1024 * 1024;   // 사진 한 장 base64 상한 (원본 약 33MB)

function setup() {
  rootFolder_();
  sheet_();
  Logger.log('준비 완료: ' + rootFolder_().getUrl());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const req = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!/^[a-z0-9]{6,12}-[a-z0-9]{3,8}$/.test(req.id || '')) return out_('bad id');
    if (req.type === 'image') return saveImage_(req);
    if (req.type === 'response') return saveResponse_(req);
    return out_('unknown type');
  } catch (err) {
    console.error(err);
    return out_('error');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out_('QUOC survey endpoint');
}

/* 사진 한 장 저장 */
function saveImage_(req) {
  if (!req.data || req.data.length > MAX_B64) return out_('too large');
  const mime = /^image\/[\w.+-]+$/.test(req.mime || '') ? req.mime : 'image/jpeg';
  const idx = ('0' + (parseInt(req.index, 10) || 0)).slice(-2);
  const name = idx + '_' + String(req.filename || 'photo').replace(/[\\/:*?"<>|]/g, '_').slice(-80);
  const blob = Utilities.newBlob(Utilities.base64Decode(req.data), mime, name);
  subFolder_(req.id).createFile(blob);
  return out_('ok');
}

/* 답변 저장 + 알림 */
function saveResponse_(req) {
  const fields = (req.fields || []).filter(function (f) { return f && f.length === 2; });
  const get = function (k) { for (var i = 0; i < fields.length; i++) if (fields[i][0] === k) return String(fields[i][1] || ''); return ''; };

  // 응답자 폴더 이름 정리: 날짜_@계정_이름
  const folder = subFolder_(req.id);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd');
  folder.setName(stamp + '_' + (get('인스타그램') || '-') + '_' + (get('이름') || '-'));
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });

  const sheet = sheet_();
  const extra = [['사진 수', files.length + ' / ' + (req.photos || 0)], ['사진 폴더', folder.getUrl()], ['게시 상태', '대기'], ['제출 ID', req.id]];
  const all = fields.concat(extra);

  let headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  if (!headers.length || !headers[0]) {
    headers = ['접수 시각'];
    sheet.getRange(1, 1).setValue('접수 시각');
    sheet.setFrozenRows(1);
  }
  all.forEach(function (f) {
    if (headers.indexOf(f[0]) === -1) {
      headers.push(f[0]);
      sheet.getRange(1, headers.length).setValue(f[0]);
    }
  });
  const row = headers.map(function (h) {
    if (h === '접수 시각') return new Date();
    for (var i = 0; i < all.length; i++) if (all[i][0] === h) return all[i][1];
    return '';
  });
  sheet.appendRow(row);

  const missing = (req.photos || 0) - files.length;
  const body = fields.map(function (f) { return f[0] + ': ' + (f[1] || '-'); }).join('\n')
    + '\n\n사진 ' + files.length + '장' + (missing > 0 ? ' (보낸 사진 중 ' + missing + '장 누락, 응답자에게 다시 받아 주세요)' : '')
    + '\n' + folder.getUrl()
    + '\n\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
  MailApp.sendEmail(NOTIFY_TO, '[My QUOC 2026 설문] ' + (get('인스타그램') || '') + ' ' + (get('이름') || ''), body);
  return out_('ok');
}

/* 드라이브 폴더 */
function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('ROOT_FOLDER', f.getId());
  return f;
}

function subFolder_(sid) {
  const props = PropertiesService.getScriptProperties();
  const key = 'F_' + sid;
  const id = props.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = rootFolder_().createFolder(sid);
  props.setProperty(key, f.getId());
  return f;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function out_(t) {
  return ContentService.createTextOutput(t);
}
