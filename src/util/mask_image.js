/**
 * Masks an image by highlighting a specific area defined by a shape.
 * Creates four semi-transparent white rectangles that cover everything except the shape area.
 * Requires selecting both a shape and an image.
 * @return {boolean} True if successful, false otherwise
 */
function maskImage() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentSlide = selection.getCurrentPage();

    // Get the selected elements
    const pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      SlidesApp.getUi().alert(
        "No selection found. Please select a shape and an image."
      );
      return false;
    }

    const selectedElements = pageElementRange.getPageElements();

    if (!selectedElements || selectedElements.length !== 2) {
      SlidesApp.getUi().alert(
        "Please select exactly two elements: one shape and one image."
      );
      return false;
    }

    // Find which element is the shape and which is the image
    let shapeElement = null;
    let imageElement = null;

    for (const element of selectedElements) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        shapeElement = element.asShape();
      } else if (
        element.getPageElementType() === SlidesApp.PageElementType.IMAGE
      ) {
        imageElement = element.asImage();
      }
    }

    if (!shapeElement || !imageElement) {
      SlidesApp.getUi().alert("Please select exactly one shape and one image.");
      return false;
    }

    // Get shape and image dimensions and positions
    const shapeLeft = shapeElement.getLeft();
    const shapeTop = shapeElement.getTop();
    const shapeWidth = shapeElement.getWidth();
    const shapeHeight = shapeElement.getHeight();

    const imageLeft = imageElement.getLeft();
    const imageTop = imageElement.getTop();
    const imageWidth = imageElement.getWidth();
    const imageHeight = imageElement.getHeight();

    // Check if the shape is at least partially within the image bounds
    if (
      shapeLeft > imageLeft + imageWidth ||
      shapeLeft + shapeWidth < imageLeft ||
      shapeTop > imageTop + imageHeight ||
      shapeTop + shapeHeight < imageTop
    ) {
      SlidesApp.getUi().alert(
        "The shape and image don't overlap. Please adjust their positions."
      );
      return false;
    }

    // Create four rectangles to cover the areas outside the shape
    const overlays = [];

    // 1. Top rectangle (covers area above the shape)
    if (shapeTop > imageTop) {
      const topOverlay = currentSlide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        imageLeft,
        imageTop,
        imageWidth,
        shapeTop - imageTop
      );
      topOverlay.getFill().setSolidFill("#FFFFFF", 0.5);
      topOverlay.getBorder().setWeight(0.1); // Minimum valid weight
      topOverlay.getBorder().getLineFill().setSolidFill("#FFFFFF", 0); // Transparent border
      overlays.push(topOverlay);
    }

    // 2. Bottom rectangle (covers area below the shape)
    if (shapeTop + shapeHeight < imageTop + imageHeight) {
      const bottomOverlay = currentSlide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        imageLeft,
        shapeTop + shapeHeight,
        imageWidth,
        imageTop + imageHeight - (shapeTop + shapeHeight)
      );
      bottomOverlay.getFill().setSolidFill("#FFFFFF", 0.5);
      bottomOverlay.getBorder().setWeight(0.1); // Minimum valid weight
      bottomOverlay.getBorder().getLineFill().setSolidFill("#FFFFFF", 0); // Transparent border
      overlays.push(bottomOverlay);
    }

    // 3. Left rectangle (covers area to the left of the shape)
    if (shapeLeft > imageLeft) {
      const leftOverlay = currentSlide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        imageLeft,
        Math.max(imageTop, shapeTop),
        shapeLeft - imageLeft,
        Math.min(imageTop + imageHeight, shapeTop + shapeHeight) -
          Math.max(imageTop, shapeTop)
      );
      leftOverlay.getFill().setSolidFill("#FFFFFF", 0.5);
      leftOverlay.getBorder().setWeight(0.1); // Minimum valid weight
      leftOverlay.getBorder().getLineFill().setSolidFill("#FFFFFF", 0); // Transparent border
      overlays.push(leftOverlay);
    }

    // 4. Right rectangle (covers area to the right of the shape)
    if (shapeLeft + shapeWidth < imageLeft + imageWidth) {
      const rightOverlay = currentSlide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        shapeLeft + shapeWidth,
        Math.max(imageTop, shapeTop),
        imageLeft + imageWidth - (shapeLeft + shapeWidth),
        Math.min(imageTop + imageHeight, shapeTop + shapeHeight) -
          Math.max(imageTop, shapeTop)
      );
      rightOverlay.getFill().setSolidFill("#FFFFFF", 0.3);
      rightOverlay.getBorder().setWeight(0.1); // Minimum valid weight
      rightOverlay.getBorder().getLineFill().setSolidFill("#FFFFFF", 0); // Transparent border
      overlays.push(rightOverlay);
    }

    // Create an array of all elements to group (overlays and original image, but not the shape)
    const pageElements = [imageElement, ...overlays].filter(Boolean);

    // Group all elements together
    if (pageElements.length > 1) {
      currentSlide.group(pageElements);
    }

    // Delete the shape as it's no longer needed
    shapeElement.remove();

    return true;
  } catch (error) {
    SlidesApp.getUi().alert("Error: " + error.message);
    console.log("Error in maskImage: " + error.message);
    console.log(error.stack);
    return false;
  }
}
