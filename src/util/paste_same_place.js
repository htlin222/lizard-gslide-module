/**
 * Covers the selected element with a semi-transparent white rectangle.
 * Works with any element type including images, shapes, and textboxes.
 * @return {boolean} True if successful, false otherwise
 */
function coverImageWithWhite() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    // Get the selected element
    const pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      SlidesApp.getUi().alert("No selection found. Please select an element.");
      return false;
    }

    const selectedElements = pageElementRange.getPageElements();

    if (!selectedElements || selectedElements.length === 0) {
      SlidesApp.getUi().alert("Please select at least one element.");
      return false;
    }

    // Use the first selected element
    const element = selectedElements[0];

    // Get element dimensions and position
    const elementWidth = element.getWidth();
    const elementHeight = element.getHeight();
    const left = element.getLeft();
    const top = element.getTop();

    // Create a transparent white rectangle
    const shape = currentSlide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      left,
      top,
      elementWidth,
      elementHeight
    );

    // Set the fill to white with 50% transparency
    const fill = shape.getFill();
    fill.setSolidFill("#FFFFFF", 0.5); // White color with alpha 0.5

    // Set the border to a minimal weight and make it transparent
    shape.getBorder().setWeight(0.1); // Minimum valid weight
    shape.getBorder().getLineFill().setSolidFill("#FFFFFF", 0.5); // Transparent border

    // Send the shape to back so it's behind the element
    shape.sendBackward();

    return true;
  } catch (error) {
    SlidesApp.getUi().alert("Error: " + error.message);
    return false;
  }
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

/**
 * Validates if the selection contains an image and returns it.
 * @param {Selection} selection - The current selection in the presentation
 * @return {Image|null} The selected image element or null if not found
 */
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
