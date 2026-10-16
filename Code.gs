/**
 * ============================================
 * Code.gs — Main Entry Point
 * WWTP Operations Data Entry System v6
 * ============================================
 */

function doGet(e) {
  var t = HtmlService.createTemplateFromFile('Index');
  // Tanggal & shift default diambil dari waktu lokal server (zona waktu script)
  t.serverInfoJson = JSON.stringify(getServerInfo()).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('WWTP Operations System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
