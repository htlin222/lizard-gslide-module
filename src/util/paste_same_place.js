function coverImageWithWhite() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentSlide = selection.getCurrentPage();

  const imageElement = validateImage(selection);
  if (!imageElement) return;

  const imgWidth = imageElement.getWidth();
  const imgHeight = imageElement.getHeight();
  const left = imageElement.getLeft();
  const top = imageElement.getTop();

  // Create a transparent white rectangle (middle layer)
  const shape = currentSlide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    left,
    top,
    imgWidth,
    imgHeight
  );

  // Set the fill to white with 50% transparency
  const fill = shape.getFill();
  fill.setSolidFill("#FFFFFF", 0.5); // White color with alpha 0.5
}

// Utility for pasting content in the same place in Google Slides
function duplicateImageInPlace() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentSlide = selection.getCurrentPage();

  const imageElement = validateImage(selection);
  if (!imageElement) return;

  const imgWidth = imageElement.getWidth();
  const imgHeight = imageElement.getHeight();
  const left = imageElement.getLeft();
  const top = imageElement.getTop();

  // Create a transparent white rectangle (middle layer)
  const shape = currentSlide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    left,
    top,
    imgWidth,
    imgHeight
  );

  // Set the fill to white with 50% transparency
  const fill = shape.getFill();
  fill.setSolidFill("#FFFFFF", 0.5); // White color with alpha 0.5

  // Set the border to the same color and transparency
  shape.getBorder().setWeight(1); // 1 point border weight
  shape.getBorder().getLineFill().setSolidFill("#FFFFFF", 0.5); // White border with alpha 0.5

  // Create a new image (top layer)
  const blob = imageElement.getBlob();
  const newImage = currentSlide.insertImage(blob);
  newImage.setLeft(left);
  newImage.setTop(top);
  newImage.setWidth(imgWidth);
  newImage.setHeight(imgHeight);
}

function validateImage(selection) {
  const pageElementRange = selection.getPageElementRange();

  if (!pageElementRange) {
    SlidesApp.getUi().alert("No selection found. Please select an image.");
    return null;
  }

  const selectedElements = pageElementRange.getPageElements();

  if (!selectedElements || selectedElements.length === 0) {
    SlidesApp.getUi().alert("Please select at least one image.");
    return null;
  }

  for (const element of selectedElements) {
    if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      return element.asImage();
    }
  }

  SlidesApp.getUi().alert("Selection does not contain any images.");
  return null;
}
