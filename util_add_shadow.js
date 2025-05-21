function createOffsetBlueShape() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const currentPage = selection.getCurrentPage();
  const ui = SlidesApp.getUi();

  const selectedElement = getSingleSelectedElement(selection, ui);
  if (!selectedElement) return;

  const offset = promptForOffset(ui);
  if (offset === null) return;

  const { x, y, width, height } = getElementDimensions(selectedElement);

  const blueShape = drawOffsetShape(currentPage, x, y, width, height, offset, main_color);

  const triangles = drawDecorativeTriangles(currentPage, x, y, width, height, offset, main_color);

  const overlay = applyBorderOrOverlay(currentPage, selectedElement, x, y, width, height, main_color);

  selectedElement.bringToFront();
  
  // Group all four shapes together
  const pageElements = [selectedElement, blueShape, ...triangles, overlay].filter(Boolean);
  currentPage.group(pageElements);
}

// Utility Functions

function getSingleSelectedElement(selection, ui) {
  const pageElementRange = selection.getPageElementRange();
  if (!pageElementRange) {
    ui.alert("Please select exactly one item.");
    return null;
  }

  const selectedElements = pageElementRange.getPageElements();
  if (selectedElements.length !== 1) {
    ui.alert("Please select exactly one item.");
    return null;
  }

  return selectedElements[0];
}

function promptForOffset(ui) {
  const response = ui.prompt('Enter offset (number of pixels to shift X and Y)', 'e.g. 10', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const offset = parseFloat(response.getResponseText().trim());
  if (isNaN(offset)) {
    ui.alert("Invalid number. Please enter a numeric value.");
    return null;
  }

  return offset;
}

function getElementDimensions(element) {
  return {
    x: element.getLeft(),
    y: element.getTop(),
    width: element.getWidth(),
    height: element.getHeight()
  };
}

function drawOffsetShape(page, x, y, width, height, offset, color) {
  const shape = page.insertShape(SlidesApp.ShapeType.RECTANGLE, x + offset, y + offset, width, height);
  shape.getFill().setSolidFill(color);
  shape.getBorder().getLineFill().setSolidFill(color);
  return shape;
}

function drawDecorativeTriangles(page, x, y, width, height, offset, color) {
  const triangle1 = insertTriangle(page, x + width, y, offset, color, false);       // Right side triangle
  const triangle2 = insertTriangle(page, x + offset, y + height, offset, color, true); // Bottom-left mirrored triangle
  return [triangle1, triangle2];
}

function insertTriangle(page, x, y, size, color, isLeft) {
  const triangle = page.insertShape(SlidesApp.ShapeType.RIGHT_TRIANGLE, x, y, size, size);

  if (isLeft) {
    triangle.scaleWidth(-1);
    triangle.setRotation(-90);
  }

  triangle.getFill().setSolidFill(color);
  triangle.getBorder().setWeight(1);
  triangle.getBorder().getLineFill().setSolidFill(color);
  
  return triangle;
}

function applyBorderOrOverlay(page, element, x, y, width, height, color) {
  try {
    const border = element.getBorder();
    if (border) {
      border.getLineFill().setSolidFill(color);
      return null;
    }
  } catch (e) {
    // Element may not support border
  }

  const overlay = page.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, width, height);
  overlay.getFill().setTransparent();
  overlay.getBorder().getLineFill().setSolidFill(color);
  overlay.bringToFront();
  
  return overlay;
}
