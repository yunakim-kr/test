/**
 * 콜드메일 자동 발송 도구
 * - 시트에 입력한 명단으로 콜드메일을 발송하고,
 *   3일 뒤 답장이 없으면 후속 메일을 한 번 더 보낸다.
 */

// 시트 열(컬럼) 번호 정의: A=1, B=2 ... 순서
var COL = {
  SENDER: 1,   // 보내는 사람
  NAME: 2,     // 이름
  COMPANY: 3,  // 회사
  TITLE: 4,    // 직책
  EMAIL: 5,    // 이메일
  MEMO: 6,     // 메모(나만 봄)
  STATUS: 7    // 상태
};

/**
 * 시트를 열면 실행되어 '콜드메일' 메뉴를 만든다.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('콜드메일')
    .addItem('처음 설정하기 (헤더 만들기)', 'setUpHeaders')
    .addItem('지금 발송하기', 'sendMailsNow')
    .addItem('자동 후속 메일 켜기', 'turnOnAutoFollowUp')
    .addItem('자동 후속 메일 끄기', 'turnOffAutoFollowUp')
    .addToUi();
}

/**
 * '처음 설정하기': 1행에 헤더를 만들고 서식을 적용한다.
 */
function setUpHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = ['보내는 사람', '이름', '회사', '직책', '이메일', '메모(나만 봄)', '상태'];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');      // 헤더 글자 굵게
  headerRange.setBackground('#d9d9d9');   // 헤더 배경 회색
  sheet.setFrozenRows(1);                 // 1행 고정

  SpreadsheetApp.getActiveSpreadsheet().toast('헤더 설정이 완료되었습니다.', '처음 설정하기', 5);
}

/**
 * 메일 제목과 본문을 만드는 함수.
 * {{sender}}, {{name}}, {{company}}, {{title}} 자리표시자를
 * 실제 값으로 바꿔서 반환한다.
 */
function buildMail(sender, name, company, title) {
  // 수강생 수정: 메일 제목 템플릿을 원하는 문구로 바꾸세요
  var subjectTemplate = '{{name}}님, {{company}}의 업무 자동화를 제안드립니다';

  // 수강생 수정: 메일 본문 템플릿을 원하는 문구로 자유롭게 수정하세요
  var bodyTemplate =
    '{{name}} {{title}}님, 안녕하세요.\n\n' +
    '저는 {{sender}}입니다. {{company}}에서 반복적으로 처리하시는 업무를 자동화로 줄여드릴 수 있을 것 같아 연락드립니다.\n\n' +
    '지금 보고 계신 이 메일도 구글 시트와 연동해 자동으로 발송된 것입니다\n\n' +
    '혹시 관심이 있으시면 짧게 통화나 미팅으로 자세히 설명드리고 싶습니다. 편하신 시간을 알려주시면 감사하겠습니다.\n\n' +
    '감사합니다.\n{{sender}} 드림';

  return {
    subject: fillPlaceholders_(subjectTemplate, sender, name, company, title),
    body: fillPlaceholders_(bodyTemplate, sender, name, company, title)
  };
}

/**
 * 문자열 안의 {{sender}}, {{name}}, {{company}}, {{title}} 자리표시자를
 * 실제 값으로 바꿔주는 내부 함수.
 */
function fillPlaceholders_(text, sender, name, company, title) {
  return text
    .replace(/{{sender}}/g, sender)
    .replace(/{{name}}/g, name)
    .replace(/{{company}}/g, company)
    .replace(/{{title}}/g, title);
}

/**
 * '지금 발송하기': 확인 팝업을 먼저 보여주고,
 * 사용자가 '예'를 눌렀을 때만 메일을 발송한다.
 */
function sendMailsNow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();

  if (lastRow < 2) {
    ui.alert('발송할 데이터가 없습니다.');
    return;
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  var data = dataRange.getValues();

  var targetIndexes = [];   // 발송 대상이 되는 data 배열 인덱스
  var alreadySentCount = 0; // 이미 발송완료라서 제외되는 건수

  for (var i = 0; i < data.length; i++) {
    var email = data[i][COL.EMAIL - 1];
    var status = String(data[i][COL.STATUS - 1]);

    if (!email) continue; // 이메일이 없는 행은 건너뜀

    if (status.indexOf('발송완료') === 0) {
      alreadySentCount++;
      continue; // 이미 첫 메일을 보낸 행은 다시 보내지 않음
    }

    targetIndexes.push(i);
  }

  var remainingQuota = MailApp.getRemainingDailyQuota();
  var sendableCount = Math.min(targetIndexes.length, remainingQuota);

  var message =
    '이번에 보낼 수 있는 메일: ' + targetIndexes.length + '건\n' +
    '오늘 남은 발송 한도: ' + remainingQuota + '건\n' +
    '이미 발송완료로 제외되는 건수: ' + alreadySentCount + '건';

  if (targetIndexes.length > remainingQuota) {
    message += '\n\n남은 한도(' + remainingQuota + '건)만큼만 발송됩니다.';
  }

  var response = ui.alert('콜드메일 발송 확인', message, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    return; // '아니오'를 누르면 아무것도 보내지 않음
  }

  var todayString = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (var j = 0; j < sendableCount; j++) {
    var rowIndex = targetIndexes[j];
    var rowValues = data[rowIndex];

    var sender = rowValues[COL.SENDER - 1];
    var name = rowValues[COL.NAME - 1];
    var company = rowValues[COL.COMPANY - 1];
    var title = rowValues[COL.TITLE - 1];
    var email = rowValues[COL.EMAIL - 1];

    var mail = buildMail(sender, name, company, title);
    MailApp.sendEmail(email, mail.subject, mail.body);

    var sheetRowNum = rowIndex + 2; // data 인덱스를 실제 시트 행 번호로 변환
    sheet.getRange(sheetRowNum, COL.STATUS).setValue('발송완료 ' + todayString);

    Utilities.sleep(2000); // 스팸 인식을 줄이기 위해 발송 간 2초 대기
  }
}

/**
 * 첫 메일을 보낸 지 3일이 지났고 아직 답장/후속 처리가 안 된 행을 확인해서,
 * 답장이 있으면 '답장옴'으로, 없으면 후속 메일을 보내고 '후속완료'로 표시한다.
 * 자동 후속 메일 트리거에서 매일 실행된다.
 */
function sendFollowUps() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  var data = dataRange.getValues();
  var now = new Date();

  for (var i = 0; i < data.length; i++) {
    var email = data[i][COL.EMAIL - 1];
    var status = String(data[i][COL.STATUS - 1]);

    if (!email) continue;
    if (status.indexOf('답장옴') === 0 || status.indexOf('후속완료') === 0) continue; // 이미 처리된 행은 건너뜀
    if (status.indexOf('발송완료') !== 0) continue; // 첫 메일을 아직 안 보낸 행은 건너뜀

    var sentDate = new Date(status.replace('발송완료', '').trim());
    if (isNaN(sentDate.getTime())) continue;

    var daysPassed = Math.floor((now.getTime() - sentDate.getTime()) / (1000 * 60 * 60 * 24));
    if (daysPassed < 3) continue; // 3일이 지나지 않은 행은 건너뜀

    var sheetRowNum = i + 2;

    // 해당 이메일 주소에서 나에게 보낸 메일(답장)이 있는지 Gmail에서 검색
    var threads = GmailApp.search('from:' + email);

    if (threads.length > 0) {
      sheet.getRange(sheetRowNum, COL.STATUS).setValue('답장옴');
    } else {
      var sender = data[i][COL.SENDER - 1];
      var name = data[i][COL.NAME - 1];
      var company = data[i][COL.COMPANY - 1];
      var title = data[i][COL.TITLE - 1];

      var mail = buildMail(sender, name, company, title);
      MailApp.sendEmail(email, mail.subject, mail.body);
      sheet.getRange(sheetRowNum, COL.STATUS).setValue('후속완료');

      Utilities.sleep(2000); // 스팸 인식을 줄이기 위해 발송 간 2초 대기
    }
  }
}

/**
 * '자동 후속 메일 켜기': sendFollowUps가 매일 오전 9시에 실행되도록
 * 시간 기반 트리거를 설치한다. 기존 트리거는 먼저 삭제해 중복을 막는다.
 */
function turnOnAutoFollowUp() {
  removeFollowUpTriggers_();

  ScriptApp.newTrigger('sendFollowUps')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('자동 후속 메일이 켜졌습니다. (매일 오전 9시)', '자동 후속 메일 켜기', 5);
}

/**
 * '자동 후속 메일 끄기': sendFollowUps에 연결된 트리거를 모두 삭제한다.
 */
function turnOffAutoFollowUp() {
  removeFollowUpTriggers_();
  SpreadsheetApp.getActiveSpreadsheet().toast('자동 후속 메일이 꺼졌습니다.', '자동 후속 메일 끄기', 5);
}

/**
 * sendFollowUps 함수에 연결된 트리거를 모두 찾아 삭제하는 내부 함수.
 */
function removeFollowUpTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendFollowUps') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
