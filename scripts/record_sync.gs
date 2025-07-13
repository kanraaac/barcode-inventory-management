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
  var recordSheet = ss.getSheetByName('입/출고 기록');
  var data = recordSheet.getDataRange().getValues();
  if (data.length < 2) return;

  // 1. 제조사별, 사이즈별로 재고 및 유효기간 집계
  // (컬럼 인덱스는 실제 시트 구조에 맞게 조정 필요)
  // 0:입고날짜, 1:출고날짜, 2:바코드, 3:제품명, 4:사이즈, 5:모델타입, 6:제조사, 7:제조일자, 8:유효기간
  var stockMap = {}; // { 제조사: { 사이즈: { modelType, count, expDates: [] } } }
  for (var i = 2; i < data.length; i++) { // 2행부터 데이터 시작
    var row = data[i];
    var inDate = row[0];
    var outDate = row[1];
    var barcode = row[2];
    var size = row[4];
    var modelType = row[5];
    var maker = row[6];
    var expDate = row[8];
    if (!size || !maker) continue;

    // 입고/출고/사용 구분
    var isIn = inDate && !outDate;
    var isOut = outDate && !inDate;

    if (!stockMap[maker]) stockMap[maker] = {};
    if (!stockMap[maker][size]) stockMap[maker][size] = { modelType: modelType, count: 0, expDates: [] };
    // 모델타입이 여러개면 가장 마지막 값으로 덮어씀(동일 사이즈에 여러 모델타입이 있을 경우)
    stockMap[maker][size].modelType = modelType;

    if (isIn) {
      stockMap[maker][size].count += 1;
      if (expDate) stockMap[maker][size].expDates.push(expDate);
    } else if (isOut) {
      stockMap[maker][size].count -= 1;
      // 출고된 유효기간은 제외(가장 짧은 유효기간 계산에서)
      if (expDate) {
        var idx = stockMap[maker][size].expDates.indexOf(expDate);
        if (idx !== -1) stockMap[maker][size].expDates.splice(idx, 1);
      }
    }
  }

  // 2. 제조사별 재고표 시트 생성/업데이트
  Object.keys(stockMap).forEach(function(maker) {
    var sheetName = maker + '재고표';
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
        minExp = expDates.sort()[0];
      }
      sheet.appendRow([modelType, size, count, minExp]);
    });
  });
}
