  // 출고 처리 및 바코드 기반 자동 정보 추출, 3가지 정보(제품모델+유효기간+제조사) 매칭, 예외처리 포함 전체 코드
  // 주요 함수: handleStockOut

  /**
  * 출고입력 시트에서 바코드 입력 후 출고 처리
  * - 3가지 정보(제품모델+유효기간+제조사)만 일치하면 출고 가능
  * - 예외처리 및 안내 메시지 포함
  */
  function handleStockOut() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOut = ss.getSheetByName('출고입력');
    var sheetStock = ss.getSheetByName('현재고 정보');
    var sheetHistory = ss.getSheetByName('출고 이력');
    var sheetRecord = ss.getSheetByName('입/출고 기록');
    var ui = SpreadsheetApp.getUi();

    var lastRow = sheetOut.getLastRow();
    for (var i = 2; i <= lastRow; i++) {
      var barcode = sheetOut.getRange('A' + i).getValue();
      if (!barcode) continue;
      var info = parseBarcode(barcode);
      // 현재고 정보에서 3가지 정보로 매칭
      var stockData = sheetStock.getDataRange().getValues();
      var foundRow = null;
      for (var j = 1; j < stockData.length; j++) {
        var stockBarcode = stockData[j][1];
        var stockInfo = parseBarcode(stockBarcode);
        if (stockInfo.model === info.model && stockInfo.expDate === info.expDate && stockInfo.maker === info.maker) {
          foundRow = j + 1;
          break;
        }
      }
      if (!foundRow) {
        sheetOut.getRange('B' + i).setValue('현재고 정보에 없는 번호');
        continue;
      }
      // 현재고 정보 시트의 해당 행 전체 데이터 읽기
      var stockRowAll = sheetStock.getRange(foundRow, 1, 1, sheetStock.getLastColumn()).getValues()[0];
      var stockRow = stockRowAll.slice(2); // 3열(모델타입)부터 끝까지
      // 출고 이력 시트에 입력
      sheetHistory.insertRowsBefore(2, 1); // 2행(헤더 아래)에 삽입
      sheetHistory.getRange(2, 1).setValue(new Date()); // 출고날짜
      sheetHistory.getRange(2, 2).setValue(barcode);    // 바코드번호
      sheetHistory.getRange(2, 3, 1, stockRow.length).setValues([stockRow]);
      // 셀 배경을 명시적으로 흰색으로 지정
      sheetHistory.getRange(2, 1, 1, stockRow.length + 2).setBackground('#FFFFFF');
      // 입/출고 기록 시트에 입력
      sheetRecord.insertRowsBefore(2, 1); // 2행(헤더 아래)에 삽입
      sheetRecord.getRange(2, 1).setValue('');          // 입고날짜 없음
      sheetRecord.getRange(2, 2).setValue(new Date());  // 출고날짜
      sheetRecord.getRange(2, 3, 1, stockRow.length + 1).setValues([[barcode].concat(stockRow)]);
      // 셀 배경을 명시적으로 하늘색으로 지정 (A~L열 전체)
      sheetRecord.getRange(2, 1, 1, 12).setBackground('#BFEFFF');
      // 현재고 정보에서 삭제
      sheetStock.deleteRow(foundRow);
      sheetOut.getRange('B' + i).setValue('출고처리 완료');
    }
    // 출고처리 후 제조사 재고표 자동 업데이트
    updateAllManufacturerStockSheets();
  }

  /**
  * 출고입력 시트에서 '사용 입력' 버튼 클릭 시 실행되는 함수
  * - handleStockOut과 동일하게 동작하되,
  * - 출고 이력 시트와 입/출고 기록 시트에 입력되는 셀의 배경색을 노란색(#FFFF00)으로 지정
  */
  function handleStockUse() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetOut = ss.getSheetByName('출고입력');
    var sheetStock = ss.getSheetByName('현재고 정보');
    var sheetHistory = ss.getSheetByName('출고 이력');
    var sheetRecord = ss.getSheetByName('입/출고 기록');
    var ui = SpreadsheetApp.getUi();

    var lastRow = sheetOut.getLastRow();
    for (var i = 2; i <= lastRow; i++) {
      var barcode = sheetOut.getRange('A' + i).getValue();
      if (!barcode) continue;
      var info = parseBarcode(barcode);
      // 현재고 정보에서 3가지 정보로 매칭
      var stockData = sheetStock.getDataRange().getValues();
      var foundRow = null;
      for (var j = 1; j < stockData.length; j++) {
        var stockBarcode = stockData[j][1];
        var stockInfo = parseBarcode(stockBarcode);
        if (stockInfo.model === info.model && stockInfo.expDate === info.expDate && stockInfo.maker === info.maker) {
          foundRow = j + 1;
          break;
        }
      }
      if (!foundRow) {
        sheetOut.getRange('B' + i).setValue('현재고 정보에 없는 번호');
        continue;
      }
      // 현재고 정보 시트의 해당 행 전체 데이터 읽기
      var stockRowAll = sheetStock.getRange(foundRow, 1, 1, sheetStock.getLastColumn()).getValues()[0];
      var stockRow = stockRowAll.slice(2); // 3열(모델타입)부터 끝까지
      // 출고 이력 시트에 입력 (노란색 배경)
      sheetHistory.insertRowsBefore(2, 1); // 2행(헤더 아래)에 삽입
      sheetHistory.getRange(2, 1).setValue(new Date()); // 출고날짜
      sheetHistory.getRange(2, 2).setValue(barcode);    // 바코드번호
      sheetHistory.getRange(2, 3, 1, stockRow.length).setValues([stockRow]);
      // 노란색 배경 적용 (A~L열 전체)
      sheetHistory.getRange(2, 1, 1, 12).setBackground('#FFFF00');
      // 입/출고 기록 시트에 입력 (노란색 배경)
      sheetRecord.insertRowsBefore(2, 1); // 2행(헤더 아래)에 삽입
      sheetRecord.getRange(2, 1).setValue('');          // 입고날짜 없음
      sheetRecord.getRange(2, 2).setValue(new Date());  // 출고날짜
      sheetRecord.getRange(2, 3, 1, stockRow.length + 1).setValues([[barcode].concat(stockRow)]);
      // 노란색 배경 적용 (A~L열 전체)
      sheetRecord.getRange(2, 1, 1, 12).setBackground('#FFFF00');
      // 현재고 정보에서 삭제
      sheetStock.deleteRow(foundRow);
      sheetOut.getRange('B' + i).setValue('사용입력 완료');
    }
    // 사용입력 후 제조사 재고표 자동 업데이트
    updateAllManufacturerStockSheets();
  }
