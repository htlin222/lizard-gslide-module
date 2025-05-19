// Date utility functions for Google Slides module
function updateDateInFirstSlide() {
  const presentation = SlidesApp.getActivePresentation();
  const firstSlide = presentation.getSlides()[0];
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const regex = /^\d{4}[-\/]\d{2}[-\/]\d{2}$/; // Matches yyyy-mm-dd or yyyy/mm/dd

  const pageElements = firstSlide.getPageElements();

  pageElements.forEach(element => {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      const textRange = shape.getText();
      const text = textRange.asString().trim();

      if (regex.test(text)) {
        textRange.setText(today);
      }
    }
  });
} 