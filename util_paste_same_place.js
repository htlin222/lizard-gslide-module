// Utility for pasting content in the same place in Google Slides
function duplicateImageInPlace() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const slide = selection.getCurrentPage();

  const imageElement = validateImage(selection);
  if (!imageElement) return;

  const imgWidth = imageElement.getWidth();
  const imgHeight = imageElement.getHeight();
  const imgX = imageElement.getLeft();
  const imgY = imageElement.getTop();

  const blob = imageElement.getBlob();
  const newImage = slide.insertImage(blob);
  newImage.setLeft(imgX);
  newImage.setTop(imgY);
  newImage.setWidth(imgWidth);
  newImage.setHeight(imgHeight);

  try {
    newImage.getBorder().setWeight(2); // default border
  } catch (e) {
    // Silently ignore if unsupported
  }
} 


function validateImage(selection) {
  const pageElementRange = selection.getPageElementRange();

  if (!pageElementRange) {
    SlidesApp.getUi().alert('No selection found. Please select an image.');
    return null;
  }

  const selectedElements = pageElementRange.getPageElements();

  if (!selectedElements || selectedElements.length === 0) {
    SlidesApp.getUi().alert('Please select at least one image.');
    return null;
  }

  for (const element of selectedElements) {
    if (element.getPageElementType() === SlidesApp.PageElementType.IMAGE) {
      return element.asImage();
    }
  }

  SlidesApp.getUi().alert('Selection does not contain any images.');
  return null;
}