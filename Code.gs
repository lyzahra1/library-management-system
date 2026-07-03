function doGet() {

  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Perpustakaan")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}

function include(file){

  return HtmlService
    .createHtmlOutputFromFile(file)
    .getContent();

}
