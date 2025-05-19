// Logging functionality for Google Slides module

function getSlideTitles() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const titles = [];

  slides.forEach((slide, index) => {
    const pageElements = slide.getPageElements();
    let titleText = "";

    for (let element of pageElements) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const text = shape.getText().asString().trim();

        if (text.length > 0) {
          titleText = text;
          break; // 假設第一個有文字的 shape 是 title
        }
      }
    }

    titles.push(`Slide ${index + 1}: ${titleText || '[No title found]'}`);
  });

  Logger.log(titles);
  return titles;
}   
 



function logSlideLayouts() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var layout = slide.getLayout();
    var layoutName = layout ? layout.getLayoutName() : 'No layout';
    console.log('Slide ' + (i + 1) + ' layout: ' + layoutName);
  }
} 

function logSectionTitleSlides() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var layout = slide.getLayout();
    var layoutName = layout ? layout.getLayoutName() : '';

    if (layoutName === 'SECTION_TITLE_AND_DESCRIPTION') {
      var shapes = slide.getShapes();
      var titleText = 'No non-empty text box found';

      for (var j = 0; j < shapes.length; j++) {
        if (shapes[j].getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
          var text = shapes[j].getText().asString().trim();
          if (text) {
            titleText = text;
            break;
          }
        }
      }

      console.log('Slide ' + (i + 1) + ' layout: ' + layoutName);
      console.log('Title text: ' + titleText);
    }
  }
}