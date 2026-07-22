function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cascadas Hotel — Restaurant')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
