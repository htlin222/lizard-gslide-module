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
						element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
				)
		: [];

	if (selectedShapes.length !== 1) {
		ui.alert(
			"Error",
			"Please select exactly one shape to split.",
			ui.ButtonSet.OK,
		);
		return;
	}

	// Create and show the dialog
	const htmlOutput = HtmlService.createHtmlOutputFromFile(
		"src/components/split-shape-dialog.html",
	)
		.setWidth(300)
		.setHeight(250);

	ui.showModalDialog(htmlOutput, "Split Shape into Grid");
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
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
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
				originalShape.getTop(),
		);
		console.log(
			"Size: Width " +
				originalShape.getWidth() +
				", Height " +
				originalShape.getHeight(),
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
					cellHeight,
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
 * Copies the style from one shape to another.
 * @param {Shape} sourceShape - The shape to copy style from.
 * @param {Shape} targetShape - The shape to apply the style to.
 */
function copyShapeStyle(sourceShape, targetShape) {
	try {
		// Copy fill
		try {
			const sourceFill = sourceShape.getFill();
			if (sourceFill && sourceFill.getType() === SlidesApp.FillType.SOLID) {
				const sourceSolidFill = sourceFill.getSolidFill();
				if (sourceSolidFill) {
					const color = sourceSolidFill.getColor();
					const alpha = sourceSolidFill.getAlpha();
					targetShape.getFill().setSolidFill(color, alpha);
				}
			}
		} catch (e) {
			console.log("Warning: Could not copy fill: " + e.message);
		}

		// Copy border - using proper border handling for Google Slides
		const sourceBorder = sourceShape.getBorder();
		const targetBorder = targetShape.getBorder();

		if (sourceBorder && targetBorder) {
			// Set border weight and dash style
			try {
				targetBorder.setWeight(sourceBorder.getWeight());
			} catch (e) {
				console.log("Warning: Could not copy border weight: " + e.message);
			}

			try {
				targetBorder.setDashStyle(sourceBorder.getDashStyle());
			} catch (e) {
				console.log("Warning: Could not copy border dash style: " + e.message);
			}

			// Set border color if it's a solid fill
			const borderLineFill = sourceBorder.getLineFill();
			if (
				borderLineFill &&
				borderLineFill.getType() === SlidesApp.FillType.SOLID
			) {
				const sourceSolidFill = borderLineFill.getSolidFill();
				if (sourceSolidFill) {
					targetBorder
						.getLineFill()
						.setSolidFill(
							sourceSolidFill.getColor(),
							sourceSolidFill.getAlpha(),
						);
				}
			}
		}

		// Copy text style if applicable
		if (sourceShape.getText() && targetShape.getText()) {
			try {
				const sourceTextStyle = sourceShape.getText().getTextStyle();
				const targetTextStyle = targetShape.getText().getTextStyle();

				// Copy basic text properties
				try {
					const fontFamily = sourceTextStyle.getFontFamily();
					if (fontFamily) {
						targetTextStyle.setFontFamily(fontFamily);
					}
				} catch (e) {
					console.log("Warning: Could not copy font family: " + e.message);
				}

				try {
					const fontSize = sourceTextStyle.getFontSize();
					if (fontSize) {
						targetTextStyle.setFontSize(fontSize);
					}
				} catch (e) {
					console.log("Warning: Could not copy font size: " + e.message);
				}

				// Copy bold styling - using safer boolean check
				try {
					const isBold = sourceTextStyle.isBold();
					if (typeof isBold === "boolean") {
						targetTextStyle.setBold(isBold);
					}
				} catch (e) {
					console.log("Warning: Could not copy bold style: " + e.message);
				}

				// Copy italic styling - using safer boolean check
				try {
					const isItalic = sourceTextStyle.isItalic();
					if (typeof isItalic === "boolean") {
						targetTextStyle.setItalic(isItalic);
					}
				} catch (e) {
					console.log("Warning: Could not copy italic style: " + e.message);
				}

				// Copy underline styling - using safer boolean check
				try {
					const isUnderline = sourceTextStyle.isUnderline();
					if (typeof isUnderline === "boolean") {
						targetTextStyle.setUnderline(isUnderline);
					}
				} catch (e) {
					console.log("Warning: Could not copy underline style: " + e.message);
				}

				// Copy text color if available
				try {
					const fontColor = sourceTextStyle.getForegroundColor();
					if (fontColor) {
						targetTextStyle.setForegroundColor(fontColor);
					}
				} catch (e) {
					console.log("Warning: Could not copy font color: " + e.message);
				}
			} catch (e) {
				console.log("Warning: Could not copy text style: " + e.message);
			}
		}
	} catch (error) {
		// Log error but continue execution
		console.log("Error copying style: " + error.message);
	}
}
