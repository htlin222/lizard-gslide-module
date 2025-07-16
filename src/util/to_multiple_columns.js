/**
 * Shows a dialog to input column parameters for splitting a textbox into multiple columns.
 */
function showMultipleColumnsDialog() {
  const ui = SlidesApp.getUi();

  // Create and show the dialog immediately - validation will happen on submit
  const htmlOutput = HtmlService.createHtmlOutput(createMultipleColumnsDialogHtml())
    .setWidth(350)
    .setHeight(200);

  ui.showModalDialog(htmlOutput, "分割成多欄");
}

/**
 * Creates the HTML content for the multiple columns dialog.
 * @return {string} The HTML content.
 */
function createMultipleColumnsDialogHtml() {
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
          <label for="columns">欄數 (Number of Columns):</label>
          <input type="number" id="columns" min="2" value="2">
        </div>
        <div class="form-group">
          <label for="gap">間距 (Gap Size in points):</label>
          <input type="number" id="gap" min="0" value="10">
        </div>
        <div class="button-container">
          <button onclick="submitForm()">分割文字框</button>
        </div>
        
        <script>
          function submitForm() {
            const columns = parseInt(document.getElementById('columns').value);
            const gap = parseInt(document.getElementById('gap').value);
            
            if (columns < 2 || gap < 0 || isNaN(columns) || isNaN(gap)) {
              alert('請輸入有效值。欄數至少為2，間距至少為0。');
              return;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('錯誤: ' + error);
              })
              .splitTextBoxToColumns(columns, gap);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Splits the selected textbox into multiple columns.
 * @param {number} columns - Number of columns to create.
 * @param {number} gap - Gap size between columns in points.
 */
function splitTextBoxToColumns(columns, gap) {
  try {
    // Get the active presentation and selection
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const selectionType = selection.getSelectionType();
    
    // Check if a page element is selected
    if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
      throw new Error("Please select a textbox first.");
    }
    
    const pageElement = selection.getPageElementRange().getPageElements()[0];
    
    // Check if the selected element is a shape (which includes text boxes in Google Slides)
    const elementType = pageElement.getPageElementType();
    if (elementType !== SlidesApp.PageElementType.SHAPE) {
      throw new Error("Please select a text box or shape.");
    }
    
    const originalShape = pageElement.asShape();
    
    // Check if the shape has text (indicating it's a text box)
    try {
      const text = originalShape.getText();
      if (!text) {
        throw new Error("Please select a shape with text content.");
      }
    } catch (e) {
      throw new Error("Please select a shape with text content.");
    }
    const slide = pageElement.getParentPage();
    
    // Get the properties of the original shape
    const originalLeft = originalShape.getLeft();
    const originalTop = originalShape.getTop();
    const originalWidth = originalShape.getWidth();
    const originalHeight = originalShape.getHeight();
    
    // Calculate the width for each column
    // column_width * column_number + gap_width * (column_count - 1) = original_width
    const columnWidth = (originalWidth - gap * (columns - 1)) / columns;
    
    if (columnWidth <= 0) {
      throw new Error("間距太大，無法創建有效的欄寬。請減少間距或欄數。");
    }
    
    // Create an array to store all new shapes
    const newShapes = [];
    
    // First, resize the original shape to the column width
    originalShape.setWidth(columnWidth);
    newShapes.push(originalShape);
    
    // Create duplicates for the remaining columns
    for (let col = 1; col < columns; col++) {
      // Duplicate the original shape (preserves all styling and text)
      const duplicatedShape = originalShape.duplicate();
      
      // Calculate position for the duplicated shape
      const left = originalLeft + col * (columnWidth + gap);
      
      // Position the duplicated shape
      duplicatedShape.setLeft(left);
      
      // Add to our array of new shapes
      newShapes.push(duplicatedShape);
    }
    
    // Log completion message
    console.log(
      "Successfully created " +
        newShapes.length +
        " shape columns with " +
        columnWidth +
        "pt width and " +
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

