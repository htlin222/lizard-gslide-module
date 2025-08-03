/**
 * Shows a dialog to input column parameters for splitting a textbox into multiple columns.
 */
function showMultipleColumnsDialog() {
	const ui = SlidesApp.getUi();

	// Create and show the dialog immediately - validation will happen on submit
	const htmlOutput = HtmlService.createHtmlOutputFromFile(
		"src/components/multiple-columns-dialog.html",
	)
		.setWidth(350)
		.setHeight(200);

	ui.showModalDialog(htmlOutput, "分割成多欄");
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
