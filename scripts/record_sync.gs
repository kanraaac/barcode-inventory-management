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

// HSL → HEX 변환 함수
function hslToHex(h, s, l) {
  s /= 100;
  l /= 100;
  let c = (1 - Math.abs(2 * l - 1)) * s;
  let x = c * (1 - Math.abs((h / 60) % 2 - 1));
  let m = l - c/2;
  let r = 0, g = 0, b = 0;
  if (0 <= h && h < 60) { r = c; g = x; b = 0; }
  else if (60 <= h && h < 120) { r = x; g = c; b = 0; }
  else if (120 <= h && h < 180) { r = 0; g = c; b = x; }
  else if (180 <= h && h < 240) { r = 0; g = x; b = c; }
  else if (240 <= h && h < 300) { r = x; g = 0; b = c; }
  else if (300 <= h && h < 360) { r = c; g = 0; b = x; }
  r = Math.round((r + m) * 255);
  g = Math.round((g + m) * 255);
  b = Math.round((b + m) * 255);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}
// 모델타입별 고유 색상 생성 함수 (HEX 반환)
function getColorByModelType(modelType) {
  var hash = 0;
  for (var i = 0; i < modelType.length; i++) {
    hash = (hash * 31 + modelType.charCodeAt(i)) % 360;
  }
  return hslToHex(hash, 60, 85); // 밝고 부드러운 색상
}

// 유효기간 배열에서 가장 임박한(가장 빠른) 날짜를 찾는 함수
function getEarliestExpDate(expDates) {
  var minDate = null;
  var minRaw = '';
  expDates.forEach(function(d) {
    var dt = null;
    if (typeof d !== 'string') return;
    if (d.length === 6 && /^\d{6}$/.test(d)) {
      // YYMMDD → YYYY-MM-DD
      dt = new Date('20' + d.substr(0,2) + '-' + d.substr(2,2) + '-' + d.substr(4,2));
    } else if (d.length === 10 && /^\d{4}-\d{2}-\d{2}$/.test(d)) {
      dt = new Date(d);
    } else if (d.length === 10 && /^\d{4}\.\d{2}\.\d{2}$/.test(d)) {
      dt = new Date(d.replace(/\./g, '-'));
    }
    if (dt && !isNaN(dt.getTime())) {
      if (!minDate || dt < minDate) {
        minDate = dt;
        minRaw = d;
      }
    }
  });
  // 변환 실패시 첫 번째 값 fallback
  return minRaw || (expDates.length > 0 ? expDates[0] : '');
}

// 날짜 문자열을 Date 객체로 변환 (YYMMDD, YYYY-MM-DD, YYYY.MM.DD 지원)
function parseDate(str) {
  if (!str) return null;
  if (str.length === 6 && /^\d{6}$/.test(str)) {
    // YYMMDD → Date
    return new Date('20' + str.substr(0,2) + '-' + str.substr(2,2) + '-' + str.substr(4,2));
  }
  if (str.length === 10 && /^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return new Date(str);
  }
  if (str.length === 10 && /^\d{4}\.\d{2}\.\d{2}$/.test(str)) {
    return new Date(str.replace(/\./g, '-'));
  }
  return null;
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

  // 1. 제조사별, 모델타입별, 사이즈별로 재고 및 유효기간 집계
  // 0:입고날짜, 1:바코드, 2:모델타입, 3:사이즈, 4:제조번호, 5:제작사, 6:제조일자, 7:유효기간
  var stockMap = {}; // { 제조사: { 모델타입: { 사이즈: { count, expDates: [] } } } }
  for (var i = 1; i < data.length; i++) { // 1행부터 데이터 시작(헤더 제외)
    var row = data[i];
    var modelType = row[2];
    var size = row[3];
    var maker = row[5];
    var expDate = row[7];
    if (!size || !maker || !modelType) continue;

    if (!stockMap[maker]) stockMap[maker] = {};
    if (!stockMap[maker][modelType]) stockMap[maker][modelType] = {};
    if (!stockMap[maker][modelType][size]) stockMap[maker][modelType][size] = { count: 0, expDates: [] };
    stockMap[maker][modelType][size].count += 1;
    if (expDate) stockMap[maker][modelType][size].expDates.push(expDate);
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
    var typeMap = stockMap[maker];
    Object.keys(typeMap).forEach(function(modelType) {
      var sizeMap = typeMap[modelType];
      Object.keys(sizeMap).forEach(function(size) {
        var count = sizeMap[size].count;
        var expDates = sizeMap[size].expDates;
        var minExp = getEarliestExpDate(expDates);
        var minExpDate = parseDate(minExp);
        sheet.appendRow([modelType, size, count, minExpDate || minExp || '']);
        // 모델타입별 셀 배경색 지정 (동적 색상)
        var rowIdx = sheet.getLastRow();
        var color = getColorByModelType(modelType);
        sheet.getRange(rowIdx, 1, 1, 4).setBackground(color);
      });
    });
  });
}