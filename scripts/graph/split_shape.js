/**
 * Shows a dialog to input grid parameters for splitting a shape.
 */
function showSplitShapeDialog() {
  const ui = SlidesApp.getUi();

  // Check if a shape is selected
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectedShapes = selection.getPageElementRange()
    ? selection
        .getPageElementRange()
        .getPageElements()
        .filter(
          (element) =>
            element.getPageElementType() === SlidesApp.PageElementType.SHAPE
        )
    : [];

  if (selectedShapes.length !== 1) {
    ui.alert(
      "Error",
      "Please select exactly one shape to split.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Create and show the dialog
  const htmlOutput = HtmlService.createHtmlOutput(createSplitShapeDialogHtml())
    .setWidth(300)
    .setHeight(250);

  ui.showModalDialog(htmlOutput, "Split Shape into Grid");
}

/**
 * Creates the HTML content for the split shape dialog.
 * @return {string} The HTML content.
 */
function createSplitShapeDialogHtml() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 10px;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
          }
          input {
            width: 100%;
            padding: 5px;
            box-sizing: border-box;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
            margin-top: 20px;
          }
          button {
            padding: 8px 16px;
            background-color: #4285f4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
          }
          button:hover {
            background-color: #2a75f3;
          }
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="rows">Number of Rows:</label>
          <input type="number" id="rows" min="1" value="2">
        </div>
        <div class="form-group">
          <label for="columns">Number of Columns:</label>
          <input type="number" id="columns" min="1" value="2">
        </div>
        <div class="form-group">
          <label for="gap">Gap Size (points):</label>
          <input type="number" id="gap" min="0" value="5">
        </div>
        <div class="button-container">
          <button onclick="submitForm()">Split Shape</button>
        </div>
        
        <script>
          function submitForm() {
            const rows = parseInt(document.getElementById('rows').value);
            const columns = parseInt(document.getElementById('columns').value);
            const gap = parseInt(document.getElementById('gap').value);
            
            if (rows < 1 || columns < 1 || gap < 0 || isNaN(rows) || isNaN(columns) || isNaN(gap)) {
              alert('Please enter valid values. Rows and columns must be at least 1, and gap must be at least 0.');
              return;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Error: ' + error);
              })
              .splitSelectedShape(rows, columns, gap);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Splits the selected shape into a grid of shapes.
 * @param {number} rows - Number of rows in the grid.
 * @param {number} columns - Number of columns in the grid.
 * @param {number} gap - Gap size between shapes in points.
 */
function splitSelectedShape(rows, columns, gap) {
  try {
    // Get the active presentation and selection
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const slide = selection.getCurrentPage();

    // Get the selected shape
    const selectedElements = selection.getPageElementRange().getPageElements();
    const selectedShapes = selectedElements.filter(
      (element) =>
        element.getPageElementType() === SlidesApp.PageElementType.SHAPE
    );

    if (selectedShapes.length !== 1) {
      throw new Error("Please select exactly one shape to split.");
    }

    const originalShape = selectedShapes[0].asShape();

    // Log shape information for debugging
    console.log("Element type: " + originalShape.getPageElementType());
    console.log("Element ID: " + originalShape.getObjectId());
    console.log(
      "Position: Left " +
        originalShape.getLeft() +
        ", Top " +
        originalShape.getTop()
    );
    console.log(
      "Size: Width " +
        originalShape.getWidth() +
        ", Height " +
        originalShape.getHeight()
    );
    console.log("Rotation: " + originalShape.getRotation() + " degrees");
    console.log("Shape type: " + originalShape.getShapeType());

    // Get the properties of the original shape
    const originalLeft = originalShape.getLeft();
    const originalTop = originalShape.getTop();
    const originalWidth = originalShape.getWidth();
    const originalHeight = originalShape.getHeight();
    const originalRotation = originalShape.getRotation();
    const originalShapeType = originalShape.getShapeType();

    // Calculate the dimensions for each new shape
    const cellWidth = (originalWidth - gap * (columns - 1)) / columns;
    const cellHeight = (originalHeight - gap * (rows - 1)) / rows;

    // Create an array to store all new shapes
    const newShapes = [];

    // Create the grid of shapes
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < columns; col++) {
        // Calculate position for the new shape
        const left = originalLeft + col * (cellWidth + gap);
        const top = originalTop + row * (cellHeight + gap);

        // Create the new shape
        const newShape = slide.insertShape(
          originalShapeType,
          left,
          top,
          cellWidth,
          cellHeight
        );

        // Apply rotation if the original shape had any
        if (originalRotation !== 0) {
          newShape.setRotation(originalRotation);
        }

        // Copy the styling from the original shape
        copyShapeStyle(originalShape, newShape);

        // Add to our array of new shapes
        newShapes.push(newShape);
      }
    }

    // Remove the original shape
    originalShape.remove();

    // Log completion message
    console.log(
      "Successfully created " +
        newShapes.length +
        " shapes in a " +
        rows +
        "x" +
        columns +
        " grid with " +
        gap +
        "pt gaps"
    );
  } catch (error) {
    SlidesApp.getUi().alert(
      "Error",
      "An error occurred: " + error.message,
      SlidesApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Copies the style from one shape to another.
 * @param {Shape} sourceShape - The shape to copy style from.
 * @param {Shape} targetShape - The shape to apply the style to.
 */
function copyShapeStyle(sourceShape, targetShape) {
  try {
    // Copy fill
    const sourceFill = sourceShape.getFill();
    const targetFill = targetShape.getFill();

    if (sourceFill.getType() === SlidesApp.FillType.SOLID) {
      targetFill.setSolidFill(
        sourceFill.getSolidFill().getColor(),
        sourceFill.getSolidFill().getAlpha()
      );
    }

    // Copy border - using proper border handling for Google Slides
    const sourceBorder = sourceShape.getBorder();
    const targetBorder = targetShape.getBorder();

    // Set border weight and dash style
    targetBorder.setWeight(sourceBorder.getWeight());
    targetBorder.setDashStyle(sourceBorder.getDashStyle());

    // Set border color if it's a solid fill
    const borderFill = sourceBorder.getFill();
    if (borderFill.getType() === SlidesApp.FillType.SOLID) {
      targetBorder.setSolidFill(
        borderFill.getSolidFill().getColor(),
        borderFill.getSolidFill().getAlpha()
      );
    }

    // Copy text style if applicable
    if (sourceShape.getText() && targetShape.getText()) {
      const sourceTextStyle = sourceShape.getText().getTextStyle();
      const targetTextStyle = targetShape.getText().getTextStyle();

      // Copy basic text properties
      if (sourceTextStyle.getFontFamily()) {
        targetTextStyle.setFontFamily(sourceTextStyle.getFontFamily());
      }
      if (sourceTextStyle.getFontSize()) {
        targetTextStyle.setFontSize(sourceTextStyle.getFontSize());
      }
      targetTextStyle.setBold(sourceTextStyle.isBold());
      targetTextStyle.setItalic(sourceTextStyle.isItalic());
      targetTextStyle.setUnderline(sourceTextStyle.isUnderline());

      // Copy text color if available
      const fontColor = sourceTextStyle.getForegroundColor();
      if (fontColor) {
        targetTextStyle.setForegroundColor(fontColor);
      }
    }
  } catch (error) {
    // Log error but continue execution
    console.log("Error copying style: " + error.message);
  }
}
