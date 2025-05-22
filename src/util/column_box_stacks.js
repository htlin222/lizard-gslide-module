function addThreeColumn() {
// Creates 3 columns, 30px gap between columns, 180px wide boxes, with blue color
createBoxColumns(3, 30, 180, main_color);
}

function createBoxColumns(columnCount, gapBetweenColumns, boxWidth, colorHex) {
  const slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  const startX = 30;
  const topY = 150;
  const topHeight = 40;
  const bottomY = topY + topHeight;
  const bottomHeight = 120;
  const strokeWidth = 2; // in points
  const white = "#FFFFFF";

  for (let i = 0; i < columnCount; i++) {
    const currentX = startX + i * (boxWidth + gapBetweenColumns);

    // Top box (filled with color)
    const topBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, currentX, topY, boxWidth, topHeight);
    topBox.getFill().setSolidFill(colorHex);
    topBox.getBorder().getLineFill().setSolidFill(colorHex);
    topBox.getBorder().setWeight(strokeWidth);

    // Bottom box (white fill, colored border)
    const bottomBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, currentX, bottomY, boxWidth, bottomHeight);
    bottomBox.getFill().setSolidFill(white);
    bottomBox.getBorder().getLineFill().setSolidFill(colorHex);
    bottomBox.getBorder().setWeight(strokeWidth);
  }
}