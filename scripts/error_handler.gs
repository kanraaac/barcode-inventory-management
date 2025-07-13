// 예외 및 안내 메시지 처리

function logError(context, err) {
  var msg = '[' + context + '] ' + (err && err.stack ? err.stack : err);
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('오류가 발생했습니다.\n' + msg);
}
