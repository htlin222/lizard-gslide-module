/**
 * Converts a selected shape or textbox into a callout by adding two additional shapes:
 * 1. A header shape above the main element
 * 2. A vertical bar to the left of both shapes
 * 
 * The function checks if a shape or textbox is selected, then:
 * - Sets the main element's fill color to white
 * - Sets the main element's border color to blue
 * - Sets the main element's text color to black
 * - Creates a header shape above the main element
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
  
  // Check if the selected element is a shape or a text box
  const elementType = pageElement.getPageElementType();
  if (elementType !== SlidesApp.PageElementType.SHAPE && elementType !== SlidesApp.PageElementType.TEXT_BOX) {
    SlidesApp.getUi().alert('Please select a shape or a text box.');
    return;
  }
  
  // Get the element (either shape or text box)
  let mainElement;
  if (elementType === SlidesApp.PageElementType.SHAPE) {
    mainElement = pageElement.asShape();
  } else {
    mainElement = pageElement.asTextBox();
  }
  const slide = pageElement.getParentPage();
  
  // Get position and size of the main element
  const mainElementX = mainElement.getLeft();
  const mainElementY = mainElement.getTop();
  const mainElementWidth = mainElement.getWidth();
  const mainElementHeight = mainElement.getHeight();
  
  // Style the main element
  mainElement.getFill().setSolidFill('#FFFFFF'); // White fill
  mainElement.getBorder().getLineFill().setSolidFill(main_color); // Blue border
  
  // Set text color to black and alignment to left if there's text
  try {
    const mainText = mainElement.getText();
    mainText.getTextStyle().setForegroundColor('#000000');
    mainText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  } catch (e) {
    // No text in the element, ignore
  }
  
  // Create the header shape (1st shape)
  const headerHeight = 20;
  const headerShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    mainElementX,
    mainElementY - headerHeight,
    mainElementWidth,
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
    .setForegroundColor(main_color) // Blue text
    .setBold(true) // Bold text
    .setFontSize(12); // 12pt font size
  headerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  headerShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); // Vertically center text
  
  // Create the vertical bar (2nd shape)
  const barWidth = 3;
  const barHeight = mainElementHeight + headerHeight;
  const barShape = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    mainElementX - barWidth,
    mainElementY - headerHeight,
    barWidth,
    barHeight
  );
  
  // Style the vertical bar
  barShape.getFill().setSolidFill(main_color); // Blue fill
  barShape.getBorder().getLineFill().setSolidFill(main_color); // Blue border
  
  // Group all three elements together
  const pageElements = [barShape, headerShape, mainElement];
  slide.group(pageElements);
}
