namespace sheetNamespace {

    /**
     * 기존 은행 관리 시트 공통 클래스
     * 현재 관리 은행 종류 : 기업 은행 , 카카오 뱅크
     *
     * 현재 시트 한개당 통장 한 개의 거래 정보를 모두 기록한다.
     * 해당 시트의 참조를 포함하여 주요 정보를 추출하여 멤버 변수로 지정하는 클래스
     */
    class LegacyAccount {

        /**
         * @param {Types.Sheet} legacySheet - 현재 검토대상 스프레드시트
         */
        sheet: Types.Sheet          // 관련 시트 객체 
        accountNum: Types.Range     // '계좌번호' 텍스트를 가진 단일 셀 Range 
        lastBanking: Types.Range    // '거래일시' Range 바로 아랫행 단일 셀 Range
        metaData: Types.Range       // '계좌번호' Range를 포함하는 데이타리전 Range
        mainDataTitle: Types.Range  // 
        newData: any
        constructor(legacySheet) {
            this.sheet = legacySheet;
            this.accountNum = this.sheet.createTextFinder('계좌번호').findNext();
            this.lastBanking = this.sheet.createTextFinder('거래일시').findNext().offset(1, 0);
            this.metaData = this.accountNum.getDataRegion();
        }

        /**
         * 메타데이터 영역 값만 업데이트
         * 입력받는 메타데이터 영역이 현재 영역과 형식이 동일해야 함.
         * @param {Types.Range} newMetaData - 신규메타데이터 Range
         */
        updateMetadata(newMetaData) {
            // 메타데이터 영역 복사 
            if (newMetaData) {
                Library.Logger.info("메타데이터 영역 복사");
                this.metaData.setValues(newMetaData.getValues());
            }
            else {
                Library.Logger.log("makeNewMetaData() 실패");
            }
        }

        /**
         * 메인데이터 영역 업데이트
         *
         * @param {Types.Range} newMainData 신규시트에서 선택된 업데이트할 영역
         */
        updateMaindata(newMainData) {
            if (newMainData) {
                Library.Logger.log("업데이트할 거래 횟수 : %d", newMainData.getNumRows());
                // 신규 데이타 영역만큼 빈 Row 삽입 
                // 중요 사항: insertRowsAfter는 지정된 영역의 형식을 밑으로 전파 하지만, 나중에 메인데이타 형식으로 덮어쓸것임
                this.sheet.insertRowsAfter(this.mainDataTitle.getRow(), newMainData.getNumRows());
                /**
                 * 메인데이타 영역의 첫줄을 삽입할 데이타영역의 Row 크기만큼 복사함
                 * 메인 데이타 영역의 테두리등 서식을 미리 복사해놓는 효과를 꾀함 ( 이후 작업으로 값만 복사하면 됨)
                 * 오른쪽 사이드에 별도 정리 영역도 같이 복사됨 ( 줄단위로 복사하므로 )
                 */
                const mainDataUpperLeft = this.mainDataTitle.offset(newMainData.getNumRows() + 1, 0);
                const mainDatafirstRow = this.sheet.getRange(mainDataUpperLeft.getRow(), mainDataUpperLeft.getColumn(), 1, this.sheet.getLastColumn());
                mainDatafirstRow.copyTo(this.mainDataTitle.offset(1, 0, newMainData.getNumRows()));
                // 값만 복사함
                newMainData.copyTo(this.mainDataTitle.offset(1, 0), { contentsOnly: true });
                Library.Logger.log("업데이트 완료");
            }
            else {
                Library.Logger.warning('신규 데이터에 문제가 있어 업데이트 하지 않음');
            }
        }

        /**
         * 신규 통장 멤버의 게으른 초기화 이후에 호출되는 업데이트 함수
         */
        updateNewData() {
            // 메타데이터 영역을 신규 데이터로 업데이트
            this.updateMetadata(this.newData.metaData);
            // 신규 메인데이터에서 업데이트 할 부분 선택
            this.updateMaindata(this.newData.getNewBankingRange(this.lastBanking.getValue()));
        }

        /**
         * 기존 통장 클래스의 하위 멤버인 신규 통장 클래스의 게으른 초기화를 호출하기 위한 wrapper함수
         *
         * @param newSheet - 신규 통장 시트
         */
        newDataSetup(newSheet) {
            this.newData.lazyInit(this.sheet, newSheet, this.accountNum.getValue());
        }
    }

    /**
     * 기업 은행 기존 시트 관리 클래스
     */
    class LegacyIBKAccount extends LegacyAccount {
        constructor(legacySheet) {
            super(legacySheet);
            this.mainDataTitle = this.sheet.createTextFinder('No').findNext();
            this.newData = new NewIBKdata();
        }
    }

    /**
     * 카카오 뱅크 기존 시트 핸들링 클래스
     */
    export class LegacyKKOAccount extends LegacyAccount {
        constructor(legacySheet) {
            super(legacySheet);
            this.mainDataTitle = this.sheet.createTextFinder('거래일시').findNext();
            this.newData = new NewKakao();
        }
    }

    /**
     * 신규 데이터 기본 클래스
     *
     * ---추상 클래스로 만든 이유
     * 은행별로 처리 프로세스는 다르지 않아 하나의 클래스로 통일하기 위함
     * 달라지는 초기화 정보만 추상 메소드로 다르게 구분하도록 함
     * 추상 메소드를 쓰기 위해선 추상 클래스로 선언되어야 함
     */
    abstract class NewBase {
        sheet: Types.Sheet
        timeFormat: string
        accountNum: Types.Range
        lastBankingNum: Types.Range 
        mainData: Types.Range
        metaData: Types.Range
        abstract initMainData()
        abstract initMetaData()
        /**
         * 신규 통장 클래스 객체 생성시에 초기화에 필요한 정보가 없으므로 나중에 초기화 하는 함수
         *
         * --- 필요 사유
         * 기존 통장 클래스 객체 멤버로 신규 통장 클래스 객체도 포함하도록 구조변경
         * 기존 통장 객체 정보로 드라이브에서 신규 통장 정보를 긁어오므로 생성시에는 신규 통장 초기화 불가
         *
         * @param {Types.Sheet} legacySheet  - 기존 통장 시트
         * @param {Types.Sheet} newSheet  - 신규로 읽어온 통장 시트
         * @param {string} legacyAccountNum  - 매칭되는 기존 통장 시트의 계좌번호 정보
         */
        lazyInit(legacySheet, newSheet, legacyAccountNum) {
            if (newSheet) {
                this.sheet = newSheet;
                const info = new LegacySheetInfo();
                this.timeFormat = info.getInfo(legacySheet.getName(), "timeFormat"); // child class에서 정의 
                Library.Logger.log(`타임포맷은 ${this.timeFormat}`);
                this.accountNum = this.sheet.createTextFinder('계좌번호').findNext();
                Library.Logger.log(`계좌번호: ${this.accountNum.getValue()}`);
                this.checkMetaData(legacyAccountNum);
                this.initMainData();
                // 메인데이터 내림차순 정렬이후에 최근거래일 가져와야 함
                this.lastBankingNum = this.sheet.createTextFinder('거래일시').findNext().offset(1, 0).getValue();
                Library.Logger.log(`마지막거래일시: ${this.lastBankingNum}`);
            }
            else {
                Library.Logger.severe(`찾는 신규 데이타 시트가 없습니다`);
            }
        }
        /**
        * 신규로 들어온 정보가 전 처리가 필요한 경우 정의함
        *
        * @param {string} legacyAccountNum - legacy 통장시트의 계좌번호 셀의 문자열 값
        */
        checkMetaData(legacyAccountNum) {
            this.initMetaData();
            if (legacyAccountNum == this.accountNum.getValue()) {
                Library.Logger.info(`계좌번호 일치 : ${legacyAccountNum}`);
            }
            else {
                const errorStr = `계좌번호 불일치 : ${legacyAccountNum} =/= ${this.accountNum.getValue()}`;
                Library.Logger.severe(errorStr);
                throw new Error(errorStr);
            }
        }
        /**
         *  신규시트에서, "거래 일시" 컬럼에서 입력받은 날짜를 검색한다.
         *
         * newData.maindata는 제목줄을 포함하므로 제목줄을 제외하고
         * 기존 최종거래일도 제외하는 메인데이타 영역을 리턴한다
         *
         * @param {Object} legacyLastDate 검색할 날짜를 입력받는다. Legacy 시트의 최종거래시간이 입력된다.
         * @return {Types.Range} 추가할 신규 데이타 영역을 반환한다.
         */
        getNewBankingRange(legacyLastDate) {
            const strLld = Library.moment(legacyLastDate, this.timeFormat).format("YYYY-MM-DD");
            Library.Logger.log("기존 시트, 최종 업데이트 날짜: %s", strLld);
            // 기존 마지막 거래날짜 위치를 신규 데이타 영역에서 검색하여 offsetRowIndex를 찾음.
            const offsetRowIndex = Utils.indexOfinDate(this.mainData.getValues(), '거래일시', legacyLastDate, (src, target) => Library.moment(src, this.timeFormat).isSame(target));
            let result = null;
            if (offsetRowIndex <= 1) {
                // 객체 초기화시에 계산된 속성명을 쓸 수 있음 : 
                // https://wiki.developer.mozilla.org/ko/docs/Web/JavaScript/Reference/Operators/Object_initializer#%EA%B3%84%EC%82%B0%EB%90%9C_%EC%86%8D%EC%84%B1%EB%AA%85
                Library.Logger.severe({
                    [1]: `업데이트 불필요, 기존 시트와 신규 시트의  최종 거래가 동일함 : ${strLld}`,
                    [0]: `제목줄에서 찾았다면 이상함. 확인필요`,
                    [-1]: `찾는 날짜가 없음, 찾는 날짜: ${strLld}, 신규 시트 최종: ${Library.moment(this.lastBankingNum, this.timeFormat).format("YYYY-MM-DD")}`
                }[offsetRowIndex]);
            }
            else {
                // 기존 마지막 거래일 이후의 신규 데이타 영역을 반환한다. 
                result = this.mainData.offset(1, 0, offsetRowIndex - 1);
                result.setBackground('yellow');
            }
            return result;
        }
    }

    /**
     * 신규로 읽어온 기업은행 정보 저장용 클래스
     *
     */
    class NewIBKdata extends NewBase {
        /**
         * 신규 데이타의 통장 시트의 메타 데이터가 기존 계좌번화와 같은지 확인하고 형식을 맞춘다
         *
         * IBK의 경우 메터데이터 4줄이 한 셀로 들어오는데 (4행 6열로) 분리가 필요함
         * 최종 목적은 this.metaData를 초기화 하는것
         */
        initMetaData() {
            // 1. 수정할 영역 찾기 
            const checkDate = this.sheet.createTextFinder('조회기준일').findNext();
            // 2. 병합된 영역 컬럼으로 분리하기
            const firstMergedRange = this.accountNum.getMergedRanges()[0]; // 계좌번호로 시작하는 병합 영역
            const secondMergedRange = checkDate.getMergedRanges()[0]; // 조회기준일로 시작하는 병합 영역
            firstMergedRange.breakApart(); // 1 row 6 columns으로 병합 찢기 
            secondMergedRange.breakApart();
            this.accountNum.splitTextToColumns('\n'); // 4 rows를 컬럼으로 배열 
            checkDate.splitTextToColumns('\n'); // 즉, 1 row 4 columns로 배열되며 2 column은 빈 칸으로 남음 
            // 3. 신규시트의 메타데이터아래에 4줄 삽입: transpose전 공간 확보
            // transpose 된 이후에 빈 줄이 1줄 포함되도록 한다. 
            // range.getDataRegion으로 메타데이터만 선택될 수 있도록 빈 줄을 삽입한다.
            this.sheet.insertRowsAfter(this.accountNum.getRow(), 4);
            // 4. 컬럼으로 분리된 영역을 transpose하고 다시 병합하기 
            Utils.transpose_range(this.sheet.getRange(this.accountNum.getRow(), this.accountNum.getColumn(), 1, 4));
            Utils.transpose_range(this.sheet.getRange(checkDate.getRow(), checkDate.getColumn(), 1, 4));
            for (let i = 0; i < 4; i++) {
                firstMergedRange.offset(i, 0).merge();
                secondMergedRange.offset(i, 0).merge();
            }
            this.metaData = this.accountNum.getDataRegion();
        }
        /**
         * 신규로 가져온 시트의 메인데이타 정보를 기존 시트와 형식을 맞춤
         *
         * 최종 결과물은 this.mainData의 초기화
         * IBK는 신규 데이터의 메인데이타가 내림차순이므로 정렬 필요없음
         */
        initMainData() {
            // 데이타 영역의 upperLeft 코너의 셀을 기준으로 영역을 정의한다. 
            const upperLeftCell = this.sheet.createTextFinder('No').findNext();
            this.mainData = upperLeftCell.getDataRegion();
        }
    }

    /**
     * 카카오 뱅크 신규 데이터 클래스
     *
     * 특징
     * 메인데이터 영역이 날짜가 오름차순 (최신이 아래)으로 되어있는데 내림차순 ( 최신이 위)로 바꿔야 함
     *
     */
    class NewKakao extends NewBase {
        /**
         * 신규로 들어온 정보가 전 처리가 필요한 경우 정의함
         *
         * @param {string} legacyAccountNum - legacy 통장시트의 계좌번호 셀의 문자열 값
         */
        initMetaData() {
            //카카오 메타데이타는 단순하게 getDataRegion() 하면 "카카오 거래내역" 제목까지 포함하게 되어 정확한 좌표설정 필요함 
            this.metaData = this.sheet.createTextFinder('성명').findNext().offset(0, 0, 2, 6);
        }
        /**
         * 신규로 가져온 시트의 메인데이타 정보를 추출함
         *
         * 원하는 값을 검색하기 위해 제목줄을 포함함
         * 내림차순일 경우 오름차순으로 정렬
         * 금액 부분이 문자열일 경우 금액이 계산가능하도록 포맷을 숫자로 변경함
         *
         * 맨 마지막에 lastBankingNum을 설정하는 이유
         * 카카오 뱅크 신규데이터는 초기에 내림차순으로 데이터가 되어있어 오름차순 정렬이후 가져와야 함
         */
        initMainData() {
            const mainDataTotal = this.sheet.createTextFinder('거래일시').findNext().getDataRegion();
            const onlyData = mainDataTotal.offset(1, 0, mainDataTotal.getNumRows() - 1);
            // 메인데이타 오름차순으로 sorting
            onlyData.sort({ column: onlyData.getColumn(), ascending: false });
            // 메인데이타 중 금액부분을 숫자로 설정
            onlyData.offset(0, 2, onlyData.getNumRows(), 2).setNumberFormat("0,#");
            // 메인데이터로 확정 (제목줄까지 포함)
            this.mainData = onlyData.getDataRegion();
            // 메인데이터중 가장 최근거래일, "거래일시"가 첫번째 컬럼(0)
            this.lastBankingNum = this.mainData.offset(1, 0).getValue();
        }
    }

    /**
     * 은행별 엑셀파일 출력형식 정보
     */
    interface BankInfo {
        nameList: string[];     // 카카오뱅크 관련 시트 이름 배열
        excelExt: string;       // 엑셀파일 가져올 때 확장자
        timeFormat: string;     // 엑셀파일에서 사용된 비표준 time 형식 (moment Library에서 인식못함)
    }

    /**
     * 기존에 있던 시트의 엑셀출력형식 정보  
     */
    class LegacySheetInfo {
        kakao: BankInfo         // 카카오뱅크 다운로드하는 엑셀파일 출력형식 정보 
        ibk: BankInfo           // 기업은행에서 다운로드하는 엑셀파일 출력형식 정보
        allmember: BankInfo[]   // 전체 은행시트 엑셀출력 형식 목록
        constructor() {
            this.kakao = {
                nameList: ['경조사계좌'],
                excelExt: '.xlsx',
                timeFormat: 'YYYY.MM.DD HH:mm:ss',
            };
            this.ibk = {
                nameList: ['개인계좌', '화실계좌', '메인계좌', '집세계좌'],
                excelExt: '.xls',
                timeFormat: 'YYYY. M. D a H:mm:ss',
            };
            this.allmember = [this.kakao, this.ibk];
        }
        /**
         * 기존 시트 이름을 입력하고 원하는 정보를 지정하면 해당 정보를 출력
         *
         * 디폴트 값은 IBK 인포로 지정
         *
         * @param {string} sheetName - 기존시트 이름
         * @returns {T}
         */
        getInfo(sheetName, T) {
            return this.allmember.reduce((acc, cur) => {
                if (cur.nameList.includes(sheetName))
                    acc = cur[T];
                return acc;
            }, this.ibk[T]);
        }
    }

    /**
     * 시트별로 처리 클래스가 다른데 메인에서 하나의 클래스로 처리과정을 통합하기 위해 만듬
     */
    export class BankProcessor {
        legacyAccount: LegacyKKOAccount
        constructor(legacySheet) {
            const info = new LegacySheetInfo();
            if (info.kakao.nameList.includes(legacySheet.getName())) {
                // 카카오 뱅크 계좌는 위의 리스트에 추가하면 됨
                this.legacyAccount = new LegacyKKOAccount(legacySheet);
            }
            else {
                // 그외는 모두 기업은행 계좌임
                this.legacyAccount = new LegacyIBKAccount(legacySheet);
            }
        }
        /**
         * 기존 계좌 시트에 신규 계좌 정보를 업데이트하는 처리과정 수행
         *
         * @param {Types.Sheet} newSheet  - 신규 데이터 시트
         */
        updateProcess(newSheet) {
            this.legacyAccount.newDataSetup(newSheet);
            this.legacyAccount.updateNewData();
        }
    }

}
