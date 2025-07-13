// 바코드 파싱 및 정보 추출 함수 (제조사별 규칙 분기)
// 1순위로 바코드 앞부분을 보고 제조사 판별 → 제조사별로 파싱 규칙 분기

/**
 * 바코드에서 제품모델, 제조일자, 유효기간, 제작사 추출 (제조사별 규칙 분기)
 * @param {string} barcode
 * @returns {object}
 */
function parseBarcode(barcode) {
  var maker = '';
  var model = '';
  var mfgDate = '';
  var expDate = '';

  // 1. 제조사 판별
  if (barcode && barcode.startsWith('01088000494')) {
    maker = '디오';
  } else if (barcode && barcode.startsWith('880905920')) {
    maker = '오스템'; // 예시: 오스템 바코드 시작 패턴(실제 규칙 필요)
  } else if (barcode && barcode.startsWith('880926')) {
    maker = '메가젠'; // 예시: 메가젠 바코드 시작 패턴(실제 규칙 필요)
  } else {
    maker = '기타';
  }

  // 2. 제조사별 바코드 파싱 규칙 분기
  if (maker === '디오') {
    // 디오: 제품모델 11~16(5자리), 기준위치 31 이후 11/17
    model = barcode.substring(11, 16); // 제품모델 코드 추출 (예: UF3A1)
    // 제조번호: 19~31(12자리, 예: 241025P04025)
    var serialNo = barcode.substring(18, 30); // 제조번호 추출 (예: 241025P04025)
    var mfgDate = barcode.substring(32, 38);  // 제조일자(6자리)
    var expDate = barcode.substring(40, 46);  // 유효기간(6자리)

    // 디오 타입 드롭다운 값
    var typeList = ['UF3', 'UF2', 'UV3'];
    
  } else if (maker === '오스템') {
    // 오스템: 추후 규칙 추가 예정
    // model = ...;
    // mfgDateRaw = ...;
    // expDateRaw = ...;
    var typeList = [];
  } else if (maker === '메가젠') {
    // 메가젠: 추후 규칙 추가 예정
    // model = ...;
    // mfgDateRaw = ...;
    // expDateRaw = ...;
    var typeList = [];
  } else {
    // 기타 제조사: 기본값
    var typeList = [];
  }

  return { model: model, mfgDate: mfgDate, expDate: expDate, maker: maker, typeList: typeList, serialNo: serialNo };
}
