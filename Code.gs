/**
 * ============================================
 * Code.gs — Main Entry Point
 * WWTP Operations Data Entry System
 * ============================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WWTP Operations System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testReadRecords() {
  var result = readRecords({page: 1, limit: 5});
  Logger.log(JSON.stringify(result).substring(0, 500));
  return result;
}
