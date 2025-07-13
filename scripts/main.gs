// 시스템 전체 진입점 및 메뉴 생성, onEdit 트리거
// 입고/출고 처리 메뉴, 한글 안내 포함

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('재고관리')
    .addItem('입고처리', 'handleStockIn')
    .addItem('출고처리', 'handleStockOut')
    .addItem('재고표갱신', 'updateAllManufacturerStockSheets') // 재고표 갱신 메뉴 추가
    .addToUi();
}

// 바코드 입력 시, 제조사가 '디오'이면 해당 행의 B열(제품명)에 UF3, UF2, UV3 드롭다운 생성
function setProductTypeDropdown(row, typeList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('바코드 정보');
  var range = sheet.getRange(row, 2); // B열
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(typeList, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  // 드롭다운 기본값을 첫 번째 값(UF3)으로 설정
  if (typeList && typeList.length > 0) {
    range.setValue(typeList[0]);
  }
}

// 한글 자모를 QWERTY 영문으로 변환
function hangulToQwerty(str) {
  // 한글 두벌식 키보드 자모 → QWERTY 영문 매핑 (모든 자모/복모음/쌍자음 포함)
  var map = {
    // 자음
    'ㄱ': 'r', 'ㄲ': 'R', 'ㄴ': 's', 'ㄷ': 'e', 'ㄸ': 'E', 'ㄹ': 'f', 'ㅁ': 'a', 'ㅂ': 'q', 'ㅃ': 'Q',
    'ㅅ': 't', 'ㅆ': 'T', 'ㅇ': 'd', 'ㅈ': 'w', 'ㅉ': 'W', 'ㅊ': 'c', 'ㅋ': 'z', 'ㅌ': 'x', 'ㅍ': 'v', 'ㅎ': 'g',
    // 모음
    'ㅏ': 'k', 'ㅐ': 'o', 'ㅑ': 'i', 'ㅒ': 'O', 'ㅓ': 'j', 'ㅔ': 'p', 'ㅕ': 'u', 'ㅖ': 'P',
    'ㅗ': 'h', 'ㅘ': 'hk', 'ㅙ': 'ho', 'ㅚ': 'hl',
    'ㅛ': 'y', 'ㅜ': 'n', 'ㅝ': 'nj', 'ㅞ': 'np', 'ㅟ': 'nl',
    'ㅠ': 'b', 'ㅡ': 'm', 'ㅢ': 'ml', 'ㅣ': 'l'
  };
  // 복모음(2글자) 우선 변환
  var twoCharMap = {
    'ㅘ': 'hk', 'ㅙ': 'ho', 'ㅚ': 'hl', 'ㅝ': 'nj', 'ㅞ': 'np', 'ㅟ': 'nl', 'ㅢ': 'ml', 'ㅒ': 'O', 'ㅖ': 'P'
  };
  // 복모음 우선 치환
  var converted = str;
  Object.keys(twoCharMap).forEach(function(k) {
    converted = converted.split(k).join(twoCharMap[k]);
  });
  // 단일 자모 치환
  converted = converted.replace(/[ㄱ-ㅎㅏ-ㅣ]/g, function(ch) {
    return map[ch] || ch;
  });
  return converted.toUpperCase(); // 모두 대문자로 변환
}

function onEdit(e) {
  try {
    // 바코드 정보, 입고입력, 출고입력 시트의 A열(1열) 입력 시 한글 자모를 QWERTY로 자동 변환
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();
    var col = e.range.getColumn();
    // 바코드 정보 시트는 3행부터, 입고입력/출고입력 시트는 2행부터 적용
    if ((sheetName === '바코드 정보' && col === 1 && row >= 3) ||
        ((sheetName === '입고입력' || sheetName === '출고입력') && col === 1 && row >= 2)) {
      var value = e.value;
      if (value && /[ㄱ-ㅎㅏ-ㅣ]/.test(value)) {
        var converted = hangulToQwerty(value);
        if (converted !== value) {
          sheet.getRange(row, 1).setValue(converted);
        }
      }
    }
    // 모델사이즈맵 시트의 모델타입(B열) 변경 시 처리
    if (sheetName === '모델사이즈맵' && col === 2 && row > 1) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var barcodeSheet = ss.getSheetByName('바코드 정보');
      if (barcodeSheet) {
        var lastRow = barcodeSheet.getLastRow();
        barcodeSheet.setActiveRange(barcodeSheet.getRange(lastRow + 1, 1));
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
  var size = getSizeByModel(info.model, info.serialNo); // model_size_db.gs의 함수 사용

  // C열(사이즈), D열(제조번호), E열(제작사), F열(제조일자), G열(유효기간) 자동 입력
  sheet.getRange(row, 3).setValue(size);         // C열: 사이즈
  sheet.getRange(row, 4).setValue(info.serialNo || ''); // D열: 제조번호
  sheet.getRange(row, 5).setValue(info.maker);   // E열: 제작사
  sheet.getRange(row, 6).setValue(formatDate(info.mfgDate)); // F열: 제조일자
  sheet.getRange(row, 7).setValue(formatDate(info.expDate)); // G열: 유효기간

  // typeList가 있으면 B열 드롭다운 생성
  if (info.typeList && info.typeList.length > 0) {
    setProductTypeDropdown(row, info.typeList);
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
