/**
 * Converts a selected shape into a callout by adding two additional shapes:
 * 1. A header shape above the main shape
 * 2. A vertical bar to the left of both shapes
 * 
 * The function checks if a shape is selected, then:
 * - Sets the main shape's fill color to white
 * - Sets the main shape's border color to blue
 * - Sets the main shape's text color to black
 * - Creates a header shape above the main shape
 * - Creates a vertical bar to the left of both shapes
 */
function convertShapeToCallout() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const selectionType = selection.getSelectionType();
  
  // Check if a page element is selected
  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    SlidesApp.getUi().alert('Please select a shape first.');
    return;
  }
  
  const pageElement = selection.getPageElementRange().getPageElements()[0];
  
  // Check if the selected element is a shape
  if (pageElement.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
    SlidesApp.getUi().alert('Please select a shape, not another type of element.');
    return;
  }
  
  // Get the shape
  const mainShape = pageElement.asShape();
  const slide = pageElement.getParentPage();
  
  // Get position and size of the main shape
  const mainShapeX = mainShape.getLeft();
  const mainShapeY = mainShape.getTop();
  const mainShapeWidth = mainShape.getWidth();
  const mainShapeHeight = mainShape.getHeight();
  
  // Style the main shape
  mainShape.getFill().setSolidFill('#FFFFFF'); // White fill
  mainShape.getBorder().getLineFill().setSolidFill(main_color); // Blue border
  
  // Set text color to black and alignment to left if there's text
  try {
    const mainText = mainShape.getText();
    mainText.getTextStyle().setForegroundColor('#000000');
    mainText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  } catch (e) {
    // No text in the shape, ignore
  }
  
  // Create the header shape (1st shape)
  const headerHeight = 20;
  const headerShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    mainShapeX,
    mainShapeY - headerHeight,
    mainShapeWidth,
    headerHeight
  );
  
  // Style the header shape
  headerShape.getFill().setSolidFill("#d9d9d9"); // Blue fill
  headerShape.getBorder().getLineFill().setSolidFill(main_color); // Blue border
  
  // Add default text 'INFO' to the header shape
  headerShape.getText().setText('INFO');
  
  // Set text style for header: white, bold, 12pt, centered
  const headerText = headerShape.getText();
  headerText.getTextStyle()
    .setForegroundColor('#FFFFFF') // White text
    .setBold(true) // Bold text
    .setFontSize(12); // 12pt font size
  headerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  headerShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); // Vertically center text
  
  // Create the vertical bar (2nd shape)
  const barWidth = 3;
  const barHeight = mainShapeHeight + headerHeight;
  const barShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    mainShapeX - barWidth,
    mainShapeY - headerHeight,
    barWidth,
    barHeight
  );
  
  // Style the vertical bar
  barShape.getFill().setSolidFill(main_color); // Blue fill
  barShape.getBorder().getLineFill().setSolidFill(main_color); // Blue border
  
  // Group all three shapes together
  const pageElements = [barShape, headerShape, mainShape];
  slide.group(pageElements);
}
