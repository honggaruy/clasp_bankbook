namespace QunitTests {

    var QUnit = Library.QUnitGS2.QUnit;
    const {test} = QUnit    // Object destructuring, Nested Scope 를 이용하기 위해 필요

    // 간단한 테스트 대상 함수. 나중에 필요없어지면 제거할 것
    function divideThenRound(numberator, denominator) {
        return Math.round(numberator / denominator);
    }

    // QUnit Doc 을 참고 하여 테스트 작성할 것, https://api.qunitjs.com/
    // https://api.qunitjs.com/QUnit/module/#example-hooks-on-nested-modules
    export function testsForQunit() {
        const ss = SpreadsheetApp.getActiveSpreadsheet()

        QUnit.module("기존시트 정보출력 클래스 테스트", hooks => {
            hooks.before( function(assert) { 
                console.log('LegacySheetInfo 부르기 전')
                this.info = new sheetNamespace.LegacySheetInfo()
                console.log('LegacySheetInfo 부른 후')
                this.allNames = this.info.kakao.nameList.concat(this.info.ibk.nameList, ['난계좌아님'])
                console.log(this.allNames)
                const isCorrectEnd = (name: string) => name.endsWith('계좌')
                console.log(this.allNames.every(isCorrectEnd))
                assert.ok(this.allNames.every(isCorrectEnd), "등록된 이름이 모두 ~계좌로 끝나는 이름인지 체크");
            });

            console.log('simple number 이전')
            test("simple numbers", function(assert) {
                const mapResult = ss.getSheets().map(sheet => sheet.getName())
                const isCorrectEnd = (name: string) => name.endsWith('계좌')
                assert.ok(this.allNames.every(isCorrectEnd), "이건 이상한데?");
                assert.equal(divideThenRound(10, 2), 5, "~계좌로 끝나는 이름인지 체크");
                assert.equal(divideThenRound(10, 4), 3, "decimal numbers");
            });

        });

        QUnit.start();
    }
}