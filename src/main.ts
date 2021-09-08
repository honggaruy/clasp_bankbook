/**
 * @onlyCurrentDoc
 *
 */

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡통장 업데이트📅')
        .addItem('엑셀 to 구글시트', 'convExl2Gsheet')
        .addItem('업데이트 실행', 'myFunction')
        .addItem('개발중 테스트', 'devTest')
        .addItem('초기화', 'afterCheck')
        .addToUi();
}

/**
 * 사용한 라이브러리 : BetterLog ( https://github.com/peterherrmann/BetterLog )
 *                엑셀파일에 로그를 저장하기 위해 채용
 */

 namespace Library {
    // 로그를 기록할 스프레드시트 지정
    export var Logger = BetterLog.useSpreadsheet('1qizXc_-X4iWYMUdcR9_7JetGKIV_frjeUPESSWQnAAU');
    export var moment = Moment.moment;   // Momentjs를 사용하기 위해 글로벌 객체 지정
    export var QUnitGS2 = QUnitGS2;     // 유닛 테스트 플랫폼 추가
 }

 namespace SsConfig {
     export const excludeSheets = ['대시보드'] // 먼저 통장계좌 처리에서 제외할 시트 등록 
 }

var convExl2Gsheet = Testexceltogsheet.convertExcelToGoogleSheets;

/* Qunit 결과를 보기위한 웹 앱 (테스트 배포로는 현재 확인이 안됨. 그냥 배포로 해야함 )
 * 참조 링크 : http://qunitgs2.com/examples/step-by-step-tutorial
 */
function doGet() {
    Library.QUnitGS2.init();   // initialize the library
    QunitTests.testsForQunit()  // Test는 별도 파일로 구현 
    return Library.QUnitGS2.getHtml();  //  HTML 결과로 반환
}

function getResultsFromServer() {
    return Library.QUnitGS2.getResultsFromServer();
}

function myFunction() {
    // Trigger로 실행시 Active 상태가 아니므로 강제 모드 변환 필요 
    // 그런데.. 한번도 Active가 아닌적이 없다. 언제 비활성화 되는거지? 아직 테스트 못함.
    // 참조 : https://stackoverflow.com/a/48045857/9457247
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
        Library.Logger.log('Spreadsheet이 활성화 되어 있음');
    }
    else {
        ss = SpreadsheetApp.openById('1-JyrBAU6F-74z7h3_km6ojdxiclSo-OGe-oLUGQNgac');
        SpreadsheetApp.setActiveSpreadsheet(ss);
        Library.Logger.log('Spreadsheet를 강제로 활성화 함');
    }
    for (let currentSheet of ss.getSheets()) {
        // 먼저 통장계좌 처리에서 제외할 시트 처리
        if (SsConfig.excludeSheets.includes(currentSheet.getName()))
            continue;
        const process = new sheetNamespace.BankProcessor(currentSheet);
        //var oldData = new sheetNamespace.LegacyIBKAccount(currentSheet);
        // 2. 지정된 통장시트와 관련된 파일을 검색하여 시트복사해옴
        // 시트 이름이 형식(통장이름_캡춰날짜)에 맞는지 정규식으로 검토하는 과정이 있으면 좋겠음
        // 반환된 배열에는 시트객체가 시트이름의 시간순으로 배열되어야 함
        const relatedSheets = fileManager.findRelatedFilesWith(ss, currentSheet);
        if (relatedSheets.length) {
            // 관련시트가 여러개일 경우 반복수행
            for (let newSheet of relatedSheets) {
                try {
                    Library.Logger.log("시트 : '%s', 시작", newSheet.getName());
                    process.updateProcess(newSheet);
                    //oldData.newDataSetup(newSheet)
                    //oldData.updateNewData()
                }
                catch (err) {
                    Library.Logger.severe(err.stack);
                }
                finally {
                    Library.Logger.log("시트(%s) 제거함", newSheet.getName());
                    ss.deleteSheet(newSheet);
                }
            }
        }
        else {
            Library.Logger.log("'%s' 통장 관련 신규 시트가 없습니다.", currentSheet.getName());
        }
    }
}

/**
 * 개발중 빠른 실행을 위함
 */
function devTest() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("경조사계좌");
    const handle = new sheetNamespace.LegacyKKOAccount(sheet);
    const relatedSheets = fileManager.findRelatedFilesWith(ss, sheet);
    handle.newDataSetup(ss.getSheetByName("카카오뱅크 거래내역의 사본"));
    handle.updateMetadata(handle.newData.metaData);
    handle.updateMaindata(handle.newData.getNewBankingRange(handle.lastBanking.getValue()));
    //    handle.newData.getNewBankingRange(handle.lastBanking.getValue())
    //handle.updateMetadata(newSheet.metaData)
    //handle.updateMaindata(newSheet.getNewBankingRange(handle.lastBanking.getValue()))
}
/* 개발중 결과 확인 뒤 스프레드시트 상태 초기화하는 함수
   개발 완료후 지워도 됨.
*/
function afterCheck() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet;
    for (sheet of ss.getSheets()) {
        if (sheet.getName().endsWith('의 사본')) {
            Library.Logger.log("'%s' 시트를 제거 했습니다.", sheet.getName());
            ss.deleteSheet(sheet);
        }
    }
    Library.Logger.log("파일이 초기화 되었습니다.");
}
