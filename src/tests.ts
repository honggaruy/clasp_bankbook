/*
 * Google Apps Script에서 실행되는 Unit Testing Platform을 찾다가 추가함
 * 참고 링크 : https://github.com/kevincar/gast
 */

if ((typeof GasTap) === 'undefined') { // GasT Initialization. (only if not initialized yet.)
  let cs = CacheService.getScriptCache().get('gast');
  if(!cs) {
    cs = UrlFetchApp.fetch('https://raw.githubusercontent.com/kevincar/gast/master/index.js').getContentText();
    CacheService.getScriptCache().put('gast', cs, 21600);
  }
  eval(cs)
} // Class GasTap is ready for use now!


function gastTestRunner() {
  let tap: GasTap = new GasTap();

  tap.test('do calculation right', function (t: test) {
      let i: number = 3 + 4;
      t.equal(i, 7, 'calc 3 + 4 = 7 right');
  });

  tap.test('Spreadsheet exist', function (t: test) {
      let url: string = 'https://docs.google.com/spreadsheets/d/1-JyrBAU6F-74z7h3_km6ojdxiclSo-OGe-oLUGQNgac/edit#gid=0'
      let ss: Types.Ss = SpreadsheetApp.openByUrl(url)
      t.ok(ss, 'Spread 시트 열기 성공')
  });

  tap.finish();
}