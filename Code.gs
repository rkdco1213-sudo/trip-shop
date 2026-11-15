/**
 * 여행사 장보기 – 주문 받기 (구글 시트 Apps Script)
 * 이 코드를 구글 시트의 [확장 프로그램] > [Apps Script]에 그대로 붙여 넣고
 * 웹 앱으로 배포하세요. (설정방법.md 참고)
 *
 * 시트 두 개가 자동으로 만들어져요.
 *  - 주문 : 여행사가 보낼 때마다 한 줄 (보낸 시각, 여행사, 합계, 까닭, 물건 목록)
 *  - 물건 : 담은 물건 하나마다 한 줄 (살 때 이 시트를 보면 돼요)
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();

    var orders = getSheet_(ss, '주문', ['보낸 시각', '여행사', '합계(원)', '예산(원)', '고른 까닭', '물건 목록']);
    var items = getSheet_(ss, '물건', ['보낸 시각', '여행사', '상품 이름', '가격(원)', '개수', '금액(원)', '상품 바로가기']);

    var list = (data.items || []).map(function (it) {
      return '- ' + clean_(it.name) + ' × ' + num_(it.qty) + ' (' + num_(it.price) * num_(it.qty) + '원)';
    }).join('\n');

    orders.appendRow([now, clean_(data.agency), num_(data.total), num_(data.budget), clean_(data.reason), list]);

    (data.items || []).forEach(function (it) {
      var url = safeUrl_(it.link) || safeUrl_(it.search);
      var label = it.link ? '상품 보기' : '아이스크림몰에서 찾기';
      var row = [now, clean_(data.agency), clean_(it.name), num_(it.price), num_(it.qty), num_(it.price) * num_(it.qty), ''];
      items.appendRow(row);
      if (url) {
        items.getRange(items.getLastRow(), 7).setFormula('=HYPERLINK("' + url.replace(/"/g, '%22') + '","' + label + '")');
      }
    });

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

// 주소창에 웹 앱 주소를 열었을 때 연결이 잘 됐는지 확인하는 용도
function doGet() {
  return ContentService.createTextOutput('여행사 장보기 주문 받기가 연결되어 있어요.');
}

function getSheet_(ss, name, header) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

// 수식으로 해석되지 않도록 앞에 '를 붙이고 길이를 제한해요
function clean_(v) {
  var s = String(v == null ? '' : v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function num_(v) {
  var n = Math.round(Number(v) || 0);
  return n < 0 ? 0 : n;
}

function safeUrl_(u) {
  u = String(u || '');
  return /^https:\/\/[^\s"]+$/.test(u) ? u.slice(0, 500) : '';
}
