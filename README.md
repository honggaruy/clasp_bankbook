# about clasp_bankbook

통장 사용내역을 수시로 읽어와 구글시트에 누적하는 작업을 **반자동화** 한다.

# Limitations

* 은행에 자동으로 접속하여 데이타를 긁어오진 못함.
    * 대부분의 은행이 엑셀 다운로드를 지원함.
    * 엑셀 다운로드 받는것 까지는  수동으로 진행해야 함
* 아직 내가 사용하는 "기업은행"과 "카카오통장"의 엑셀파일 형식만 지원

# Dependencies

## 작업하던 시트

* 처음에 수작업으로 진행하다가 반복되는 작업만 자동화하는게 목표라서 기존에 작업하던 구글시트가 필요하다.
* 구글시트의 구성
    * 대시보드 시트 - 각 계좌시트의 `최종 거래일`을 한 시트에서 보여준다
    * 각 계좌별 시트 
* 따라서 이 프로젝트는 [standalone project](https://developers.google.com/apps-script/guides/standalone) 타입이 아닌 [Bound to G Suite Documents](https://developers.google.com/apps-script/guides/bound) 타입이다.

## google-apps-script 라이브러리

* Test excel to google sheet - [요기](https://stackoverflow.com/a/49265306/9457247)서 가져온 코드를 라이브러리화
* moment.js 라이브러리 - [momentjs.com](https://momentjs.com/) 에서 다운로드 받아 라이브러리화
* BetterLog 라이브러리 - [momentjs.com](https://github.com/peterherrmann/BetterLog) 에 사용법이 나옴. 구글시트에 로그를 기록하는 모듈이다. 

# 작업흐름

* [엑셀 파일을 구글 시트로 바꾸는 흐름](https://honggaruy.github.io/wiki/excel2gsheet/#2-전개)을 따라간다.

1. 은행 사이트에 로그인
1. 거래내역 메뉴에서 최종 거래일 하루전부터 오늘까지 기간설정
1. 조회
1. 엑셀로 저장, 저장시에 `계좌이름_오늘날짜.xls`로 지정
    * 기업은행의 경우 `입출금_...._오늘날짜.xls`형식이므로 앞부분만 `계좌이름`으로 변경하면 됨
    * 카카오뱅크의 경우 다운로드받은 엑셀파일에 암호가 걸려있는데 풀어서 저장해야함 ( [엑셀 암호설정 메뉴에 접근하는 방법](https://support.microsoft.com/ko-kr/office/excel-파일-보호-7359d4ae-7213-4ac2-b058-f75e9311b599)으로 암호설정메뉴로 가서 빈 암호를 입력하고 저장)
    * 로컬에 백업 폴더로 설정한 `구글 드라이브` 폴더아래 지정된 `엑셀파일`폴더에 저장 
1. 필요한 엑셀파일이 모두 준비됨.
1. `통장 업데이트` 메뉴 상에서 `엑셀 to 구글시트` 클릭 
1. `업데이트 실행` 클릭
1. 제대로 업데이트 되면 `최종 거래일`에 마지막 거래일이 업데이트된다.

