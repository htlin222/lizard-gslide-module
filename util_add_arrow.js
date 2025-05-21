function drawArrowOnCurrentSlide() {
  const presentation = SlidesApp.getActivePresentation();
  const slide = presentation.getSelection().getCurrentPage();

  // Define the color blue
  const blueColor = main_color;

  // Box 1
  const box1 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 360, 200, 20, 100);
  box1.getFill().setSolidFill(blueColor);
  box1.getBorder().getLineFill().setSolidFill(blueColor);

  // Box 2
  const box2 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 360, 280, 100, 20);
  box2.getFill().setSolidFill(blueColor);
  box2.getBorder().getLineFill().setSolidFill(blueColor);

  // Box 3
  const box3 = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 404, 181, 20, 130);
  box3.setRotation(45);
  box3.getFill().setSolidFill(blueColor);
  box3.getBorder().getLineFill().setSolidFill(blueColor);
  
  // Group all three shapes together to form the arrow
  const pageElements = [box1, box2, box3];
  const groupedArrow = slide.group(pageElements);
  
  // Return the grouped arrow in case it needs to be used elsewhere
  return groupedArrow;
}

function logSelectedItemGeometry() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  const pageElements = selection.getPageElementRange().getPageElements();

  if (pageElements.length === 0) {
    Logger.log('No page element selected.');
    return;
  }

  const element = pageElements[0]; // Just taking the first selected element
  const transform = element.getTransform();

  const x = transform.getTranslateX();
  const y = transform.getTranslateY();
  const width = element.getWidth();
  const height = element.getHeight();

  Logger.log(`Selected Element Geometry:
    X: ${x}
    Y: ${y}
    Width: ${width}
    Height: ${height}`);
}