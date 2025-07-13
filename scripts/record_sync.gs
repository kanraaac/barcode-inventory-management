// 시트간 데이터 연동/동기화 함수 (예시)
// 제품모델별 재고 집계 등 활용 가능

function syncRecords() {
  try {
    // 실제 구현 시: 현재고 정보, 출고 이력, 입출고 기록 등 동기화
    // 필요시 countStockByModel 등 활용
  } catch (err) {
    logError('syncRecords', err);
  }
}

/**
 * 입출고기록 시트 전체를 분석하여 제조사별 재고표 시트를 생성/업데이트
 * - 사이즈별 재고수량, 가장 짧은 유효기간 표시
 * - 입고는 +1, 출고/사용은 -1로 집계
 */
function updateAllManufacturerStockSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stockSheet = ss.getSheetByName('현재고 정보');
  var data = stockSheet.getDataRange().getValues();
  if (data.length < 2) return;

  // 1. 제조사별, 사이즈별로 재고 및 유효기간 집계
  // 0:입고날짜, 1:바코드, 2:모델타입, 3:사이즈, 4:제조번호, 5:제작사, 6:제조일자, 7:유효기간
  var stockMap = {}; // { 제조사: { 사이즈: { modelType, count, expDates: [] } } }
  for (var i = 1; i < data.length; i++) { // 1행부터 데이터 시작(헤더 제외)
    var row = data[i];
    var modelType = row[2];
    var size = row[3];
    var maker = row[5];
    var expDate = row[7];
    if (!size || !maker) continue;

    if (!stockMap[maker]) stockMap[maker] = {};
    if (!stockMap[maker][size]) stockMap[maker][size] = { modelType: modelType, count: 0, expDates: [] };
    stockMap[maker][size].modelType = modelType;
    stockMap[maker][size].count += 1;
    if (expDate) stockMap[maker][size].expDates.push(expDate);
  }

  // 2. 제조사별 재고표 시트 생성/업데이트
  Object.keys(stockMap).forEach(function(maker) {
    var sheetName = '(' + maker + ')재고표';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['모델타입', '사이즈', '재고수량', '가장 짧은 유효기간']);
    } else {
      sheet.clearContents();
      sheet.appendRow(['모델타입', '사이즈', '재고수량', '가장 짧은 유효기간']);
    }
    var sizeMap = stockMap[maker];
    Object.keys(sizeMap).forEach(function(size) {
      var modelType = sizeMap[size].modelType;
      var count = sizeMap[size].count;
      var expDates = sizeMap[size].expDates;
      var minExp = '';
      if (expDates.length > 0) {
        var dateObjs = expDates
          .map(function(d) {
            if (typeof d !== 'string') return null;
            if (d.length === 6) {
              return new Date('20' + d.substr(0,2) + '-' + d.substr(2,2) + '-' + d.substr(4,2));
            }
            return new Date(d.replace(/\./g, '-'));
          })
          .filter(function(dt) { return dt && !isNaN(dt.getTime()); });
        if (dateObjs.length > 0) {
          var minDate = new Date(Math.min.apply(null, dateObjs));
          minExp = expDates[dateObjs.findIndex(function(dt) {
            return dt && dt.getTime() === minDate.getTime();
          })];
        } else {
          minExp = expDates[0];
        }
      }
      sheet.appendRow([modelType, size, count, minExp]);
    });
  });
}
