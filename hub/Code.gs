function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cascadas Hotel — Sistema Integrado')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
