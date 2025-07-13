// 입고 처리 및 바코드 기반 자동 정보 추출, 3가지 정보(제품모델+유효기간+제조사) 매칭, 예외처리, 재고 집계까지 포함한 전체 코드
// 주요 함수: parseBarcode, findBarcodeRowByInfo, handleStockIn, countStockByModel, clearStockInList

/**
 * 3가지 정보(제품모델+유효기간+제조사)로 바코드정보 시트에서 행 찾기
 * @param {Sheet} sheet
 * @param {string} model
 * @param {string} expDate
 * @param {string} maker
 * @returns {number|null} 실제 row 번호
 */
function findBarcodeRowByInfo(sheet, model, expDate, maker) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var barcode = data[i][0];
    var info = parseBarcode(barcode);
    if (info.model === model && info.expDate === expDate && info.maker === maker) {
      return i + 1; // 시트의 실제 row 번호
    }
  }
  return null;
}

/**
 * 바코드 정보 시트에 해당 바코드(모델+유효기간+제조사)가 없으면 자동으로 신규 행을 추가
 * @param {Sheet} sheetBarcode
 * @param {string} barcode
 * @param {object} info (parseBarcode 결과)
 * @returns {number} 추가된 행의 row 번호
 */
function addBarcodeInfoIfNotExist(sheetBarcode, barcode, info) {
  try {
    var lastRow = sheetBarcode.getLastRow();
    // 바코드 정보 시트의 헤더는 2행까지라고 가정(3행부터 데이터)
    var newRow = lastRow + 1;
    // 사이즈 정보 추출(신규 모델이면 getSizeByModel에서 prompt 발생)
    var size = getSizeByModel(info.model, info.serialNo);
    // 바코드 정보 시트 컬럼: 바코드, (B:제품명/타입), C:사이즈, D:?, E:제작사, F:제조일자, G:유효기간
    // B열(제품명/타입)은 일단 빈 값, C열:사이즈, E열:제작사, F열:제조일자, G열:유효기간
    sheetBarcode.getRange(newRow, 1).setValue(barcode); // A: 바코드
    // B열(제품명/타입)은 드롭다운 필요시 main.gs의 setProductTypeDropdown 활용
    if (info.typeList && info.typeList.length > 0) {
      setProductTypeDropdown(newRow, info.typeList);
    }
    sheetBarcode.getRange(newRow, 3).setValue(size); // C: 사이즈
    sheetBarcode.getRange(newRow, 4).setValue(info.serialNo || ''); // D: 제조번호
    sheetBarcode.getRange(newRow, 5).setValue(info.maker); // E: 제작사
    sheetBarcode.getRange(newRow, 6).setValue(formatDate(info.mfgDate)); // F: 제조일자
    sheetBarcode.getRange(newRow, 7).setValue(formatDate(info.expDate)); // G: 유효기간
    return newRow;
  } catch (err) {
    logError('addBarcodeInfoIfNotExist', err);
    throw err;
  }
}

/**
 * 입고입력 시트에서 바코드 입력 후 입고 처리
 * - 3가지 정보(제품모델+유효기간+제조사)만 일치하면 입고 가능
 * - 예외처리 및 안내 메시지 포함
 */
function handleStockIn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetInput = ss.getSheetByName('입고입력');
  var sheetBarcode = ss.getSheetByName('바코드 정보');
  var sheetStock = ss.getSheetByName('현재고 정보');
  var sheetRecord = ss.getSheetByName('입/출고 기록');
  var ui = SpreadsheetApp.getUi();

  var lastRow = sheetInput.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var barcode = sheetInput.getRange('A' + i).getValue();
    if (!barcode) continue;
    var info = parseBarcode(barcode);
    var row = findBarcodeRowByInfo(sheetBarcode, info.model, info.expDate, info.maker);
    // 바코드 정보가 없으면 자동 등록 후 재조회
    if (!row) {
      try {
        row = addBarcodeInfoIfNotExist(sheetBarcode, barcode, info);
        // 안내 메시지: 신규 바코드 정보 자동 등록됨
        sheetInput.getRange('B' + i).setValue('신규 박스정보 자동등록');
      } catch (err) {
        sheetInput.getRange('B' + i).setValue('바코드 정보 등록 오류');
        continue;
      }
      // 등록 후, 다시 row를 찾음(정상적으로 등록되었는지 확인)
      row = findBarcodeRowByInfo(sheetBarcode, info.model, info.expDate, info.maker);
      if (!row) {
        sheetInput.getRange('B' + i).setValue('바코드 정보 등록 실패');
        continue;
      }
    }
    // 현재고 정보에 이미 있으면 중복 안내
    var stockData = sheetStock.getDataRange().getValues();
    var isExist = false;
    for (var j = 1; j < stockData.length; j++) {
      var stockBarcode = stockData[j][1];
      var stockInfo = parseBarcode(stockBarcode);
      if (stockInfo.model === info.model && stockInfo.expDate === info.expDate && stockInfo.maker === info.maker) {
        isExist = true;
        break;
      }
    }
    if (isExist) {
      sheetInput.getRange('B' + i).setValue('이미 입고됨');
      continue;
    }
    // 입고 처리: 현재고 정보 시트에 추가
    // 바코드 정보 시트의 해당 행 전체 데이터 읽기
    var barcodeRow = sheetBarcode.getRange(row, 1, 1, sheetBarcode.getLastColumn()).getValues()[0];
    // 입고날짜(오늘 날짜) + 바코드 정보 전체 컬럼을 합쳐서 현재고 정보에 입력
    var newRow = [new Date()].concat(barcodeRow);
    sheetStock.insertRowsBefore(3, 1);
    sheetStock.getRange(3, 1, 1, newRow.length).setValues([newRow]);
    sheetInput.getRange('B' + i).setValue('입고처리 완료');
    // 입/출고 기록 시트에도 추가
    sheetRecord.insertRowsBefore(4, 1);
    sheetRecord.getRange(4, 1).setValue(new Date()); // 입고날짜
    sheetRecord.getRange(4, 2).setValue(''); // 출고날짜는 빈 값
    sheetRecord.getRange(4, 3, 1, barcodeRow.length).setValues([barcodeRow]);
  }
}

/**
 * 호출한(현재 활성화된) 시트의 A, B열(2행~마지막행)만 초기화
 * - 안내 메시지 없이 동작
 */
function clearStockInList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); // 현재 활성화된 시트
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 1, lastRow - 1, 2).clearContent(); // 2행~마지막행, A,B열
  sheet.setActiveRange(sheet.getRange('A1')); // 초기화 후 A2 셀로 포커스 이동
}

/**
 * 제품모델별 총재고 집계 (현재고 정보 시트 기준)
 * @returns {object} {모델: 수량}
 */
function countStockByModel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetStock = ss.getSheetByName('현재고 정보');
  var data = sheetStock.getDataRange().getValues();
  var countMap = {};
  for (var i = 1; i < data.length; i++) {
    var barcode = data[i][1];
    var info = parseBarcode(barcode);
    if (!info.model) continue;
    countMap[info.model] = (countMap[info.model] || 0) + 1;
  }
  return countMap;
}
