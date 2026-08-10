function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Headset Inventory System")
    .setFaviconUrl("https://ssl.gstatic.com/docs/script/images/favicon.ico")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename){
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}