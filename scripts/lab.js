function applyListToSelectedTextbox() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectedPageElements = selection.getPageElementRange().getPageElements();

  for (const element of selectedPageElements) {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      const textRange = shape.getText();
      const textString = textRange.asString();

      Logger.log("Text content: " + textString);

      const paragraphs = textRange.getParagraphs();
      Logger.log("Paragraph count: " + paragraphs.length);

      // 嘗試套用清單樣式
      try {
        shape.getText().getListStyle()
          .applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
        Logger.log("✅ 成功套用清單樣式");
      } catch (e) {
        Logger.log("❌ 套用失敗: " + e.message);
      }
    }
  }
}