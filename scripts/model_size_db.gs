// 제품모델-사이즈 매핑 테이블 (시트 연동)
// "모델사이즈맵" 시트에 영구 저장, [모델코드, 사이즈, 타입] 3컬럼

/**
 * 제조번호로 사이즈와 타입을 반환. 없으면 다이얼로그로 입력받아 시트에 자동 추가
 * @param {string} serialNo
 * @returns {string} 사이즈
 */
function getSizeByModel(serialNo, info) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('모델사이즈맵');
  if (!sheet) {
    sheet = ss.insertSheet('모델사이즈맵');
    sheet.appendRow(['모델코드', '모델타입', '사이즈']);
  }
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[data[i][0]] = data[i][2]; // 사이즈는 3번째 컬럼
  }
  // 기존에 있으면 반환
  var key = (info && info.maker === '메가젠') ? info.model : serialNo;
  if (map[key]) {
    return map[key];
  }
  // 메가젠 바코드일 경우 appendRow 전에 중복 체크
  if (info && info.maker === '메가젠') {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === info.model) {
        return data[i][2]; // 이미 등록된 사이즈 반환
      }
    }
    sheet.appendRow([info.model || '', info.modelType || '', info.size || '']);
    return info.size || '';
  }
  // 기존 로직 (디오/오스템 등)
  var ui = SpreadsheetApp.getUi();
  var msg = '제조번호: ' + (serialNo || '-') + '\n해당 제조번호에 맞는 사이즈를 입력하세요.';
  var response = ui.prompt('신규 제조번호 발견', msg, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var size = response.getResponseText();
    if (size) {
      var lastRow = sheet.getLastRow() + 1;
      sheet.appendRow([serialNo, '', size]);
      // 제조사별 드롭다운 및 기본값 분기
      var typeRange = sheet.getRange(lastRow, 2);
      var dropdownList = [];
      var defaultType = '';
      if (serialNo && serialNo.startsWith('FTN')) { // 오스템
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TS4BA','TS3BA', 'TS3HOI', 'MS'], true)
        .setAllowInvalid(false)
        .build();
      } else { // 디오 등 기타
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['UF3', 'UF2', 'UV3'], true)
        .setAllowInvalid(false)
        .build();
      }

      typeRange.setDataValidation(rule);
      // 드롭다운 셀로 포커스 이동
      sheet.setActiveRange(typeRange);
      return size;
    }
  }
  return '';
}

/**
 * 제조번호로 모델타입(B열)을 반환
 * @param {string} serialNo
 * @returns {string} 모델타입
 */
function getTypeByModel(serialNo, info) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('모델사이즈맵');
  if (!sheet) return '';
  var data = sheet.getDataRange().getValues();
  var key = (info && info.maker === '메가젠') ? info.model : serialNo;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1] || '';
    }
  }
  // 메가젠 바코드일 경우 info에서 바로 반환
  if (info && info.maker === '메가젠') {
    return info.modelType || '';
  }
  return '';
}

/**
 * 제조번호로 모델타입(B열) 드롭다운 리스트를 반환
 * @param {string} serialNo
 * @returns {Array<string>} 드롭다운 리스트(없으면 빈 배열)
 */
function getDropdownListBySerialNo(serialNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('모델사이즈맵');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === serialNo) {
      var cell = sheet.getRange(i + 1, 2); // B열
      var rule = cell.getDataValidation();
      if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var args = rule.getCriteriaValues();
        if (args && args[0]) return args[0];
      }
      break;
    }
  }
  return [];
}