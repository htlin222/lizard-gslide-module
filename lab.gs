// Laboratory data handling functions for Google Slides module
function createCenteredGrayTextbox() {
  // 取得目前簡報的第一張投影片
  const slide = SlidesApp.getActivePresentation().getSlides()[0];

  // 插入一個文字方塊
  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 100, 100, 300, 100);

  // 設定文字內容
  const textRange = shape.getText();
  textRange.setText("Centered Text");

  // 設定背景為灰色
  shape.getFill().setSolidFill("#CCCCCC");

  // 水平置中對齊
  const paragraphs = textRange.getParagraphs();
  for (let i = 0; i < paragraphs.length; i++) {
    paragraphs[i].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  // 垂直置中對齊
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
}

function addLineToSlide() {
  const slide = SlidesApp.getActivePresentation().getSlides()[0];
  slide.insertLine(SlidesApp.LineCategory.STRAIGHT, 100, 100, 300, 100); // x1, y1, x2, y2
}