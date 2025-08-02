/**
 * Shows a dialog to input parameters for creating child shapes inside a parent shape.
 */
function showCreateChildShapesDialog() {
	const ui = SlidesApp.getUi();

	// Check if a shape is selected
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectedShapes = selection.getPageElementRange()
		? selection
				.getPageElementRange()
				.getPageElements()
				.filter(
					(element) =>
						element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
				)
		: [];

	if (selectedShapes.length !== 1) {
		ui.alert(
			"Error",
			"Please select exactly one shape to create child shapes in.",
			ui.ButtonSet.OK,
		);
		return;
	}

	// Create and show the dialog
	const htmlOutput = HtmlService.createHtmlOutput(createChildShapesDialogHtml())
		.setWidth(350)
		.setHeight(280);

	ui.showModalDialog(htmlOutput, "Create Child Shapes");
}

/**
 * Creates the HTML content for the child shapes dialog.
 * @return {string} The HTML content.
 */
function createChildShapesDialogHtml() {
	return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 10px;
            font-size: 14px;
          }
          .form-group {
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            justify-content: space-between;
          }
          label {
            font-weight: bold;
            flex: 1;
            margin-right: 10px;
          }
          input[type="number"] {
            width: 80px;
            padding: 4px 8px;
            box-sizing: border-box;
            text-align: center;
            border: 1px solid #ccc;
            border-radius: 3px;
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
          .input-suffix {
            font-size: 12px;
            color: #666;
            margin-left: 5px;
          }
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="rows">Rows:</label>
          <input type="number" id="rows" min="1" value="2">
        </div>
        <div class="form-group">
          <label for="columns">Columns:</label>
          <input type="number" id="columns" min="1" value="1">
        </div>
        <div class="form-group">
          <label for="padding">Padding:</label>
          <div>
            <input type="number" id="padding" min="0" value="7">
            <span class="input-suffix">pt</span>
          </div>
        </div>
        <div class="form-group">
          <label for="paddingTop">Padding Top:</label>
          <div>
            <input type="number" id="paddingTop" min="0" value="30">
            <span class="input-suffix">pt</span>
          </div>
        </div>
        <div class="form-group">
          <label for="gap">Gap:</label>
          <div>
            <input type="number" id="gap" min="0" value="7">
            <span class="input-suffix">pt</span>
          </div>
        </div>
        <div class="button-container">
          <button onclick="submitForm()">Create Child Shapes</button>
        </div>
        
        <script>
          function submitForm() {
            const rows = parseInt(document.getElementById('rows').value);
            const columns = parseInt(document.getElementById('columns').value);
            const padding = parseInt(document.getElementById('padding').value);
            const paddingTop = parseInt(document.getElementById('paddingTop').value);
            const gap = parseInt(document.getElementById('gap').value);
            
            if (rows < 1 || columns < 1 || padding < 0 || paddingTop < 0 || gap < 0 || 
                isNaN(rows) || isNaN(columns) || isNaN(padding) || isNaN(paddingTop) || isNaN(gap)) {
              alert('Please enter valid values. Rows and columns must be at least 1, padding and gap must be at least 0.');
              return;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Error: ' + error);
              })
              .createChildShapesInSelected(rows, columns, padding, paddingTop, gap);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Creates child shapes inside the selected parent shape.
 * @param {number} rows - Number of rows in the grid.
 * @param {number} columns - Number of columns in the grid.
 * @param {number} padding - Padding inside the parent shape in points.
 * @param {number} paddingTop - Top padding inside the parent shape in points.
 * @param {number} gap - Gap size between child shapes in points.
 */
function createChildShapesInSelected(rows, columns, padding, paddingTop, gap) {
	try {
		// Get the active presentation and selection
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const slide = selection.getCurrentPage();

		// Get the selected shape
		const selectedElements = selection.getPageElementRange().getPageElements();
		const selectedShapes = selectedElements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (selectedShapes.length !== 1) {
			throw new Error(
				"Please select exactly one shape to create child shapes in.",
			);
		}

		const parentShape = selectedShapes[0].asShape();

		// Log parent shape information for debugging
		console.log("Parent shape type: " + parentShape.getShapeType());
		console.log("Parent shape ID: " + parentShape.getObjectId());
		console.log(
			"Parent position: Left " +
				parentShape.getLeft() +
				", Top " +
				parentShape.getTop(),
		);
		console.log(
			"Parent size: Width " +
				parentShape.getWidth() +
				", Height " +
				parentShape.getHeight(),
		);
		console.log("Parent rotation: " + parentShape.getRotation() + " degrees");

		// Get the properties of the parent shape
		const parentLeft = parentShape.getLeft();
		const parentTop = parentShape.getTop();
		const parentWidth = parentShape.getWidth();
		const parentHeight = parentShape.getHeight();
		const parentRotation = parentShape.getRotation();

		// Calculate the available space inside the parent shape after padding
		const availableWidth = parentWidth - padding * 2;
		const availableHeight = parentHeight - paddingTop - padding;

		// Calculate the dimensions for each child shape
		const childWidth = (availableWidth - gap * (columns - 1)) / columns;
		const childHeight = (availableHeight - gap * (rows - 1)) / rows;

		// Validate that child shapes will have positive dimensions
		if (childWidth <= 0 || childHeight <= 0) {
			throw new Error(
				"Padding and gap values are too large for the parent shape size.",
			);
		}

		// Create an array to store all child shapes
		const childShapes = [];

		// Create the grid of child shapes
		for (let row = 0; row < rows; row++) {
			for (let col = 0; col < columns; col++) {
				// Calculate position for the child shape relative to parent
				const childLeft = parentLeft + padding + col * (childWidth + gap);
				const childTop = parentTop + paddingTop + row * (childHeight + gap);

				// Create the child shape with the same type as parent
				const childShape = slide.insertShape(
					parentShape.getShapeType(),
					childLeft,
					childTop,
					childWidth,
					childHeight,
				);

				// Explicitly set position to ensure accuracy
				childShape.setLeft(childLeft);
				childShape.setTop(childTop);
				childShape.setWidth(childWidth);
				childShape.setHeight(childHeight);

				// Apply rotation if the parent shape has any
				if (parentRotation !== 0) {
					childShape.setRotation(parentRotation);
				}

				// Apply white fill and white stroke to child shape
				applyWhiteStyle(childShape);

				// Add to our array of child shapes
				childShapes.push(childShape);
			}
		}

		// Bring child shapes just above the parent shape
		// This maintains the parent's relative position with other elements
		for (let i = 0; i < childShapes.length; i++) {
			childShapes[i].bringForward();
		}

		// Log completion message
		console.log(
			"Successfully created " +
				childShapes.length +
				" child shapes in a " +
				rows +
				"x" +
				columns +
				" grid with " +
				padding +
				"pt padding and " +
				gap +
				"pt gaps",
		);
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			"An error occurred: " + error.message,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Applies white fill and white stroke to a shape.
 * @param {Shape} shape - The shape to apply white style to.
 */
function applyWhiteStyle(shape) {
	try {
		// Set white fill
		const fill = shape.getFill();
		fill.setSolidFill("#FFFFFF");

		// Set white border
		const border = shape.getBorder();
		border.setWeight(1); // 1pt border

		// Get the border fill and set it to white
		const borderFill = border.getFill();
		borderFill.setSolidFill("#FFFFFF");

		// Optionally set text color to black for visibility on white background
		if (shape.getText()) {
			const textStyle = shape.getText().getTextStyle();
			textStyle.setForegroundColor("#000000");
		}
	} catch (error) {
		// Log error but continue execution
		console.log("Error applying white style: " + error.message);
	}
}
