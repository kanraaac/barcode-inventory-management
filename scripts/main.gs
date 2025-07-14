// 시스템 전체 진입점 및 메뉴 생성, onEdit 트리거
// 입고/출고 처리 메뉴, 한글 안내 포함

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('재고관리')
    .addItem('재고표갱신', 'updateAllManufacturerStockSheets') // 재고표 갱신 메뉴 추가
    .addToUi();
}


function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();
    var col = e.range.getColumn();

    // 바코드 입력 셀(A열)에서 한글 입력 시 알림+값 삭제
    if (
      ((sheetName === '바코드 정보' && col === 1 && row >= 3) ||
      (sheetName === '입고입력' && col === 1 && row >= 2) ||
      (sheetName === '출고입력' && col === 1 && row >= 2))
    ) {
      var value = e.value;
      // 이미 빈 셀에서 발생한 onEdit이면 아무것도 하지 않음
      if (!value) return;
      if (/[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(value)) {
        SpreadsheetApp.getUi().alert('한글입력을 영문으로 바꿔주세요');
        sheet.getRange(row, 1).setValue('');
        // 위로 한 칸 포커스 이동 (row > 2일 때만)
        if (row > 2) {
          sheet.setActiveRange(sheet.getRange(row , 1));
        }
        return;
      }
    }

    // 모델사이즈맵 시트의 모델타입(B열) 변경 시 처리
    if (sheetName === '모델사이즈맵' && col === 2 && row > 1) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var barcodeSheet = ss.getSheetByName('바코드 정보');
      if (barcodeSheet) {
        // 변경된 제조번호와 모델타입 드롭다운을 바코드 정보 시트에 동기화
        var serialNo = sheet.getRange(row, 1).getValue(); // A열: 제조번호
        var dropdownList = getDropdownListBySerialNo(serialNo);
        var modelType = sheet.getRange(row, 2).getValue(); // B열: 모델타입 값
        var barcodeData = barcodeSheet.getDataRange().getValues();
        var firstTargetRow = null;
        for (var i = 2; i < barcodeData.length; i++) { // 3행부터
          if (barcodeData[i][3] === serialNo) { // D열(제조번호) 일치
            setProductTypeDropdown(i + 1, dropdownList, modelType);
            if (firstTargetRow === null) firstTargetRow = i + 1;
          }
        }
        // 바코드 정보 시트로 포커스 이동 (첫 번째 해당 행의 B열)
        if (firstTargetRow !== null) {
          barcodeSheet.activate();
          barcodeSheet.setActiveRange(barcodeSheet.getRange(firstTargetRow, 2));
        }
      }
      return;
    }
    autoFillBarcodeInfo(e); // 바코드 정보 시트 자동입력
    // 필요시 다른 자동화 로직 추가
  } catch (err) {
    logError('onEdit', err);
  }
}

/**
 * 바코드 정보 시트에서 바코드 입력 시 자동으로 제작사, 제조일자, 유효기간, 사이즈 입력
 * - 모델사이즈맵에 제조번호가 있으면, 해당 행의 모델타입(B열) 드롭다운 리스트와 값을 동기화
 * - 없으면 신규 사이즈 입력 및 모델타입 드롭다운은 모델사이즈맵에서 수동 지정 후 동기화
 */
function autoFillBarcodeInfo(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== '바코드 정보') return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (col !== 1 || row < 3) return; // A열, 3행 이상만 적용

  var barcode = e.value;
  if (!barcode) return;

  var info = parseBarcode(barcode); // barcode_parser.gs의 함수 사용
  Logger.log('[autoFillBarcodeInfo] info.serialNo: ' + info.serialNo + ', barcode: ' + barcode);

  // 메가젠 바코드일 경우 드롭다운/신규입력창 없이 바로 입력
  if (info.maker === '메가젠') {
    // 모델사이즈맵에 [제조번호, 모델타입, 사이즈] 자동 입력 (getSizeByModel에서 처리)
    var size = getSizeByModel(info.serialNo, info);
    // 바코드 정보 시트에 값 자동 입력
    sheet.getRange(row, 2).setValue(info.modelType || ''); // B열: 모델타입
    sheet.getRange(row, 3).setValue(info.size || '');      // C열: 사이즈
    sheet.getRange(row, 4).setValue(info.serialNo || '');  // D열: 제조번호
    sheet.getRange(row, 5).setValue(info.maker);           // E열: 제작사
    sheet.getRange(row, 6).setValue(parseDate(info.mfgDate)); // F열: 제조일자 (Date 타입)
    sheet.getRange(row, 7).setValue(parseDate(info.expDate)); // G열: 유효기간 (Date 타입)
    return;
  }

  // 디오/오스템 등 기존 로직
  var size = getSizeByModel(info.serialNo, info); // model_size_db.gs의 함수 사용
  sheet.getRange(row, 3).setValue(size);         // C열: 사이즈
  sheet.getRange(row, 4).setValue(info.serialNo || ''); // D열: 제조번호
  sheet.getRange(row, 5).setValue(info.maker);   // E열: 제작사
  sheet.getRange(row, 6).setValue(parseDate(info.mfgDate)); // F열: 제조일자 (Date 타입)
  sheet.getRange(row, 7).setValue(parseDate(info.expDate)); // G열: 유효기간 (Date 타입)

  // 모델사이즈맵에 제조번호가 있으면 드롭다운 및 값 동기화
  var dropdownList = getDropdownListBySerialNo(info.serialNo);
  var modelType = getTypeByModel(info.serialNo, info);
  if (dropdownList && dropdownList.length > 0) {
    setProductTypeDropdown(row, dropdownList, modelType);
  } else if (info.typeList && info.typeList.length > 0) {
    // 모델사이즈맵에 없고, 바코드 파싱에서 typeList가 있으면 fallback
    setProductTypeDropdown(row, info.typeList, modelType);
  }
}

/**
 * yymmdd 형식의 날짜를 yyyy-mm-dd로 변환
 */
function formatDate(yymmdd) {
  if (!yymmdd || yymmdd.length !== 6) return '';
  var year = '20' + yymmdd.substr(0, 2);
  var month = yymmdd.substr(2, 2);
  var day = yymmdd.substr(4, 2);
  return year + '-' + month + '-' + day;
}

/**
 * 날짜 문자열을 Date 객체로 변환 (YYMMDD, YYYY-MM-DD, YYYY.MM.DD 지원)
 */
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
 * 바코드 정보 시트의 B열(모델타입)에 드롭다운을 설정하고, 기본값을 지정
 * @param {number} row - 적용할 행 번호
 * @param {Array<string>} typeList - 드롭다운 리스트
 * @param {string} defaultType - 기본값(없으면 첫 번째 값)
 */
function setProductTypeDropdown(row, typeList, defaultType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('바코드 정보');
  var range = sheet.getRange(row, 2); // B열
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeList, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  // 드롭다운 기본값 지정
  if (typeList && typeList.length > 0) {
    range.setValue(defaultType || typeList[0]);
  }
}