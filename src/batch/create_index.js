function generateIndexSlide() {
  const presentation = SlidesApp.getActivePresentation();
  let slides = presentation.getSlides();

  // Step 0: 刪除所有標題為 "Index" 的 TITLE_ONLY 投影片
  for (let i = slides.length - 1; i >= 0; i--) {
    const slide = slides[i];
    const layout = slide.getLayout();
    const layoutName = layout ? layout.getLayoutName() : "";

    if (layoutName === "TITLE_ONLY") {
      const shapes = slide.getShapes();
      if (shapes.length > 0) {
        const titleText = shapes[0].getText().asString().trim();
        if (titleText === "Index") {
          slide.remove();
        }
      }
    }
  }

  slides = presentation.getSlides(); // 重新獲取投影片列表
  const indexItems = [];

  // Step 1: 收集非 SECTION_HEADER 版面投影片的標題與頁碼
  slides.forEach((slide, index) => {
    const layout = slide.getLayout();
    const layoutName = layout ? layout.getLayoutName() : "";
    if (layoutName === "SECTION_HEADER") return;

    const pageElements = slide.getPageElements();
    let titleText = "";

    for (const element of pageElements) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const text = shape.getText().asString().trim();
        if (text.length > 0) {
          titleText = text;
          break;
        }
      }
    }

    indexItems.push({
      text: `${titleText || "[No title]"}, p${index + 1}`,
      slide: slide,
    });
  });

  // Step 2: 新增 Index 投影片
  const newSlide = presentation.appendSlide(
    SlidesApp.PredefinedLayout.TITLE_ONLY
  );
  newSlide.getShapes()[0].getText().setText("Index");

  // Step 3: 插入指定寬度的文字框並添加索引項目（每10項換欄）
  const initialLeft = 40;
  const initialTop = 90;
  const width = 230;
  const height = 20;
  const spacing = 0;
  const itemsPerColumn = 12;

  indexItems.forEach((item, i) => {
    const column = Math.floor(i / itemsPerColumn);
    const row = i % itemsPerColumn;
    const left = initialLeft + column * width;
    const top = initialTop + row * (height + spacing);

    const shape = newSlide.insertShape(
      SlidesApp.ShapeType.TEXT_BOX,
      left,
      top,
      width,
      height
    );
    shape.getText().setText(item.text);
    shape.getText().getTextStyle().setFontSize(9);
    shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); // 垂直置中對齊
    shape.setLinkSlide(item.slide);
  });

  Logger.log("✅ Index slide with clickable links created.");
}
