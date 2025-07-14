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
  var serialNo = ''; // 제조번호를 저장할 변수

  // 1. 제조사 판별
  if (barcode && barcode.startsWith('01088000494')) {
    maker = '디오';
  } else if (barcode && barcode.startsWith('0108800')) {
    maker = '오스템'; // 예시: 오스템 바코드 시작 패턴(실제 규칙 필요)
  } else if (barcode && barcode.startsWith('0108809')) {
    maker = '메가젠';
  } else {
    maker = '기타';
  }

  // 2. 제조사별 바코드 파싱 규칙 분기
  if (maker === '디오') {
    // 디오: 제품모델 11~16(5자리), 기준위치 31 이후 11/17
    model = barcode.substring(11, 16); // 제품모델 코드 추출 (예: UF3A1)
    // 제조번호: 19~31(12자리, 예: 241025P04025)
    serialNo = barcode.substring(18, 30); // 제조번호 추출 (예: 241025P04025)
    var mfgDate = barcode.substring(32, 38);  // 제조일자(6자리)
    var expDate = barcode.substring(40, 46);  // 유효기간(6자리)
    var typeList = ['UF3', 'UF2', 'UV3'];   // 디오 타입 드롭다운 값
    
  } else if (maker === '오스템') {
    // 오스템 바코드 절대 위치 파싱
    // 1) 모델코드: 7~17 (11자리)
    if (barcode.length >= 18) {
      model = barcode.substring(7, 18);
    }
    // 2) 제조번호: 18~26 (9자리)
    serialNo = '';
    if (barcode.length > 18) {
      serialNo = barcode.substring(18, 27);
    }

    // 제조일자: 29~34 (6자리, YYMMDD)
    mfgDate = '';
    if (barcode.length >= 35) {
      mfgDate = barcode.substring(29, 35);
    }
    // 4) 유효기간: 35~42 (8자리)
    expDate = '';
    if (barcode.length >= 42) {
      expDate = "barcode.substring(43, 49);"
    }
    var typeList = ['TS4BA','TS3BA', 'TS3HOI', 'MS']; // 오스템 모델타입 드롭다운 자동 세팅
  } else if (maker === '메가젠') {
    // 메가젠 바코드 절대 위치 파싱 (최종)
    // 모델코드: 7~15 (9자리)
    // 제조일자: 18~23 (6자리)
    // 유효기간: 26~31 (6자리)
    // 제조번호: 34~47 (14자리)
    // 모델타입: 56~59 (4자리)
    // 사이즈: 60~65 (6자리)
    model = barcode.substring(7, 16);      // 7~15 (9자리)
    mfgDate = barcode.substring(18, 24);   // 18~23 (6자리)
    expDate = barcode.substring(26, 32);   // 26~31 (6자리)
    serialNo = barcode.substring(34, 48);  // 34~47 (14자리)
    // [메가젠 모델타입/사이즈 추출 - 바코드 뒤에서부터 기준]
    // 사이즈: 바코드 뒤에서 두 번째부터 6글자
    var size = barcode.substring(barcode.length - 7, barcode.length - 1);
    // 모델타입: 사이즈 앞 4글자(알파벳/숫자만, 3~4글자 모두 지원)
    var modelType = barcode.substring(barcode.length - 11, barcode.length - 6 ).replace(/[^A-Z]/g, '');
    var typeList = [modelType];
    return { model: model, mfgDate: mfgDate, expDate: expDate, maker: maker, typeList: typeList, serialNo: serialNo, modelType: modelType, size: size };
  } else {
    // 기타 제조사: 기본값
    var typeList = [];
  }

  return { model: model, mfgDate: mfgDate, expDate: expDate, maker: maker, typeList: typeList, serialNo: serialNo };
}
