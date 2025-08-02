function insertVerticalDashedLineBetween() {
  insertDashedLineBetween('vertical');
}

function insertHorizontalDashedLineBetween() {
  insertDashedLineBetween('horizontal');
}

function insertDashedLineBetween(orientation) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const page = selection.getCurrentPage();
    const pageElementRange = selection.getPageElementRange();

    if (!pageElementRange) {
      throw new Error("No elements selected. Please select exactly TWO shapes or text boxes.");
    }

    const selectedPageElements = pageElementRange.getPageElements();
    if (selectedPageElements.length !== 2) {
      throw new Error("Please select exactly TWO items (Shape or TextBox).");
    }

    const [elem1, elem2] = selectedPageElements;

    // Log info here — inside scope
    logElementInfo(elem1, "Element 1");
    logElementInfo(elem2, "Element 2");

    const rect1 = getElementRect(elem1);
    const rect2 = getElementRect(elem2);

    const line = createDashedLineBetween(page, rect1, rect2, orientation);
    styleLine(line);

    Logger.log("Dashed line inserted successfully.");
  } catch (e) {
    Logger.log("Error: " + e.message);
    SlidesApp.getUi().alert("Error: " + e.message);
  }
}

function getElementRect(elem) {
  return {
    left: elem.getLeft(),
    top: elem.getTop(),
    width: elem.getWidth(),
    height: elem.getHeight()
  };
}

function createDashedLineBetween(page, rect1, rect2, orientation) {
  if (orientation === 'vertical') {
    const [leftRect, rightRect] = rect1.left < rect2.left ? [rect1, rect2] : [rect2, rect1];
    const centerX = leftRect.left + leftRect.width + (rightRect.left - (leftRect.left + leftRect.width)) / 2;
    const topY = Math.min(rect1.top, rect2.top);
    const bottomY = Math.max(rect1.top + rect1.height, rect2.top + rect2.height);
    return page.insertLine(SlidesApp.LineCategory.STRAIGHT, centerX, topY, centerX, bottomY);
  } else if (orientation === 'horizontal') {
    const [topRect, bottomRect] = rect1.top < rect2.top ? [rect1, rect2] : [rect2, rect1];
    const centerY = topRect.top + topRect.height + (bottomRect.top - (topRect.top + topRect.height)) / 2;
    const leftX = Math.min(rect1.left, rect2.left);
    const rightX = Math.max(rect1.left + rect1.width, rect2.left + rect2.width);
    return page.insertLine(SlidesApp.LineCategory.STRAIGHT, leftX, centerY, rightX, centerY);
  } else {
    throw new Error("Unsupported orientation: " + orientation);
  }
}

function styleLine(line) {
  line.getLineFill().setSolidFill('#888888');
  line.setDashStyle(SlidesApp.DashStyle.DASH);
  line.setWeight(1);
}

function isShapeOrTextbox(elem) {
  if (elem.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return false;
  const shapeType = elem.asShape().getShapeType();
  return ['TEXT_BOX', 'RECTANGLE', 'ROUND_RECTANGLE', 'ELLIPSE'].includes(shapeType);
}


function logElementInfo(elem, label) {
  const type = elem.getPageElementType();
  Logger.log(`${label}: Type = ${type}`);
  if (type === SlidesApp.PageElementType.SHAPE) {
    const shapeType = elem.asShape().getShapeType();
    Logger.log(`  ShapeType = ${shapeType}`);
  }
}
