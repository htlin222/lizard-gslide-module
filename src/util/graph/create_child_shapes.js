// Configuration constants for child shapes
const DEFAULT_PADDING = 10;
const DEFAULT_PADDING_TOP = 30;
const DEFAULT_GAP = 10;
const FOOTER_BOX_HEIGHT = 15;
const HOME_PLATE_HEIGHT = 10;

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
	const htmlOutput = HtmlService.createHtmlOutputFromFile(
		"src/components/create-child-shapes-dialog.html",
	)
		.setWidth(350)
		.setHeight(280);

	ui.showModalDialog(htmlOutput, "Create Child Shapes");
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

		// Set titles for parent and child shapes
		parentShape.setTitle("PARENT");

		// Set title for each child shape and bring them forward
		for (let i = 0; i < childShapes.length; i++) {
			childShapes[i].setTitle("CHILD");
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
 * Parses grid syntax supporting two formats:
 * 1. Multi-row: {[item1|item2|item3][item4|item5]} - rows with columns
 * 2. Single-row: {item1} {item2} {item3} - simple column layout
 * @param {string} text - The text to parse
 * @return {Object|null} Grid layout object or null if invalid syntax
 */
function parseGridSyntax(text) {
	// First try to match multi-row format: {[...][...]}
	const multiRowRegex = /\{(\[.*?\])+\}/;
	const multiRowMatch = text.match(multiRowRegex);

	if (multiRowMatch) {
		return parseMultiRowSyntax(multiRowMatch[0]);
	}

	// Try to match single-row format: {} {} {}
	const singleRowRegex = /\{([^}]*)\}/g;
	const singleRowMatches = [];
	let singleRowMatch;

	while ((singleRowMatch = singleRowRegex.exec(text)) !== null) {
		const content = singleRowMatch[1].trim();
		if (content) {
			singleRowMatches.push(content);
		}
	}

	if (singleRowMatches.length > 0) {
		return parseSingleRowSyntax(singleRowMatches);
	}

	return null;
}

/**
 * Parses multi-row syntax like {[item1|item2][item3|item4]}
 * @param {string} gridText - The matched grid text
 * @return {Object} Grid layout object
 */
function parseMultiRowSyntax(gridText) {
	// Extract all row patterns [...]
	const rowRegex = /\[([^\]]*)\]/g;
	const rows = [];
	let rowMatch;

	while ((rowMatch = rowRegex.exec(gridText)) !== null) {
		const rowContent = rowMatch[1];
		// Split by | to get columns, filter out empty strings
		const columns = rowContent.split("|").filter((col) => col.trim() !== "");
		if (columns.length > 0) {
			rows.push(columns);
		}
	}

	if (rows.length === 0) {
		return null;
	}

	// Find the maximum number of columns across all rows
	const maxColumns = Math.max(...rows.map((row) => row.length));

	return {
		rows: rows.length,
		maxColumns: maxColumns,
		rowData: rows,
		isVariableColumns: rows.some((row) => row.length !== maxColumns),
		syntaxType: "multi-row",
	};
}

/**
 * Parses single-row syntax like {item1} {item2} {item3}
 * @param {Array} matches - Array of matched content strings
 * @return {Object} Grid layout object
 */
function parseSingleRowSyntax(matches) {
	return {
		rows: 1,
		maxColumns: matches.length,
		rowData: [matches], // Single row with all the columns
		isVariableColumns: false,
		syntaxType: "single-row",
	};
}

/**
 * Creates child shapes with a specific layout structure (supporting variable columns per row).
 * @param {Shape} parentShape - The parent shape
 * @param {Object} layout - The grid layout structure
 * @param {number} padding - Padding in points
 * @param {number} paddingTop - Top padding in points
 * @param {number} gap - Gap between shapes in points
 */
function createChildShapesWithLayout(
	parentShape,
	layout,
	padding,
	paddingTop,
	gap,
) {
	// Get the slide from the presentation selection
	const presentation = SlidesApp.getActivePresentation();
	const selection = presentation.getSelection();
	const slide = selection.getCurrentPage();

	// Get parent properties
	const parentLeft = parentShape.getLeft();
	const parentTop = parentShape.getTop();
	const parentWidth = parentShape.getWidth();
	const parentHeight = parentShape.getHeight();
	const parentRotation = parentShape.getRotation();

	// Calculate available space
	const availableWidth = parentWidth - padding * 2;
	const availableHeight = parentHeight - paddingTop - padding;

	// Calculate row height
	const rowHeight = (availableHeight - gap * (layout.rows - 1)) / layout.rows;

	if (rowHeight <= 0) {
		throw new Error("Parent shape is too small for the specified layout.");
	}

	const childShapes = [];

	// Create shapes for each row
	for (let rowIndex = 0; rowIndex < layout.rows; rowIndex++) {
		const rowInfo = layout.rowData[rowIndex];
		const row = rowInfo.cells || rowInfo; // Handle both new and old format
		const columnsInRow = row.length;
		const homePlates = rowInfo.homePlates || []; // Get home plate positions

		// Calculate column width for this specific row
		const columnWidth =
			(availableWidth - gap * (columnsInRow - 1)) / columnsInRow;

		if (columnWidth <= 0) {
			console.warn(
				`Row ${rowIndex + 1} has too many columns for the available width`,
			);
			continue;
		}

		// Calculate the starting Y position for this row
		const rowTop = parentTop + paddingTop + rowIndex * (rowHeight + gap);

		// Create shapes for each column in this row
		for (let colIndex = 0; colIndex < columnsInRow; colIndex++) {
			const columnLeft = parentLeft + padding + colIndex * (columnWidth + gap);

			// Create the child shape
			const childShape = slide.insertShape(
				parentShape.getShapeType(),
				columnLeft,
				rowTop,
				columnWidth,
				rowHeight,
			);

			// Set precise positioning
			childShape.setLeft(columnLeft);
			childShape.setTop(rowTop);
			childShape.setWidth(columnWidth);
			childShape.setHeight(rowHeight);

			// Apply rotation if needed
			if (parentRotation !== 0) {
				childShape.setRotation(parentRotation);
			}

			// Process the cell text for footer boxes (text) and main content
			const cellText = row[colIndex].trim();
			const footerBoxData = processFooterBoxText(cellText);

			// Adjust child shape height if footer box is needed
			let adjustedRowHeight = rowHeight;
			if (footerBoxData.hasFooter) {
				adjustedRowHeight = rowHeight - FOOTER_BOX_HEIGHT;
				childShape.setHeight(adjustedRowHeight);
			}

			// Set the main text content (without footer text)
			if (footerBoxData.mainText) {
				const textRange = childShape.getText();
				textRange.setText(footerBoxData.mainText);
			}

			// Apply styling after text is set (so we can check for [bold] markers)
			applyWhiteStyle(childShape);

			// Create footer box if needed
			if (footerBoxData.hasFooter) {
				createFooterBox(
					slide,
					columnLeft,
					rowTop + adjustedRowHeight, // Position at bottom of adjusted cell
					columnWidth,
					FOOTER_BOX_HEIGHT,
					footerBoxData.footerText,
					parentRotation,
				);
			}

			// Set title for child shape
			childShape.setTitle("CHILD");

			childShapes.push(childShape);
		}

		// Create HOME_PLATE shapes for this row if any are specified
		for (const homePlatePosition of homePlates) {
			// Calculate position between cells
			const leftCellIndex = homePlatePosition - 1;
			const rightCellIndex = homePlatePosition;

			// Only create if both cells exist
			if (leftCellIndex >= 0 && rightCellIndex < columnsInRow) {
				const leftCellRight =
					parentLeft +
					padding +
					leftCellIndex * (columnWidth + gap) +
					columnWidth;
				const rightCellLeft =
					parentLeft + padding + rightCellIndex * (columnWidth + gap);

				// Position HOME_PLATE in the gap between cells
				const homePlateLeft = leftCellRight;
				const homePlateTop = rowTop + (rowHeight - HOME_PLATE_HEIGHT) / 2; // Center vertically
				const homePlateWidth = gap; // Use gap width
				const homePlateHeight = HOME_PLATE_HEIGHT;

				// Create the HOME_PLATE shape
				const homePlate = slide.insertShape(
					SlidesApp.ShapeType.HOME_PLATE,
					homePlateLeft,
					homePlateTop,
					homePlateWidth,
					homePlateHeight,
				);

				// Apply rotation if needed
				if (parentRotation !== 0) {
					homePlate.setRotation(parentRotation);
				}

				// Set main_color fill and border
				const fill = homePlate.getFill();
				fill.setSolidFill(main_color);

				homePlate.getBorder().setWeight(1);
				homePlate.getBorder().getLineFill().setSolidFill(main_color);

				// Don't set any title for HOME_PLATE shapes (they are connectors, not content)

				// Bring HOME_PLATE forward
				homePlate.bringForward();

				console.log(
					`Created HOME_PLATE at row ${rowIndex + 1}, position ${homePlatePosition}`,
				);
			}
		}
	}

	// Bring all child shapes forward
	for (const childShape of childShapes) {
		childShape.bringForward();
	}

	// Set parent shape text alignment to top
	parentShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);

	// Set title for parent shape
	parentShape.setTitle("PARENT");

	console.log(
		`Created ${childShapes.length} child shapes with variable column layout`,
	);
}

/**
 * Applies bold styling to selected shape if text is wrapped in **asterisks**.
 * This is a standalone function for testing the bold styling feature.
 */
function applyBoldStyleToSelectedShape() {
	try {
		// Get the active presentation and selection
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		// Check if a shape is selected
		const selectedElements = selection.getPageElementRange()
			? selection.getPageElementRange().getPageElements()
			: [];

		const selectedShapes = selectedElements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (selectedShapes.length !== 1) {
			SlidesApp.getUi().alert(
				"Error",
				"Please select exactly one shape to apply bold styling.",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		const shape = selectedShapes[0].asShape();

		// Apply the bold styling transformation
		const result = applyBoldStyleTransformation(shape);

		if (result.applied) {
			SlidesApp.getUi().alert(
				"Success",
				`Bold styling applied. Text changed from "${result.originalText}" to "${result.newText}"`,
				SlidesApp.getUi().ButtonSet.OK,
			);
		} else {
			// Create debug message
			let debugMessage = "No bold styling applied.\n\n";
			debugMessage += `Original text: "${result.originalText}"\n`;
			debugMessage += `Text length: ${result.debug.textLength}\n`;
			debugMessage += `Starts with **: ${result.debug.startsWithAsterisk}\n`;
			debugMessage += `Ends with **: ${result.debug.endsWithAsterisk}\n`;

			if (result.debug.error) {
				debugMessage += `\nError: ${result.debug.error}`;
			} else if (result.debug.reason) {
				debugMessage += `\nReason: ${result.debug.reason}`;
			} else {
				debugMessage +=
					"\nText must be wrapped in **asterisks** (e.g., **text**).";
			}

			SlidesApp.getUi().alert(
				"Debug Info",
				debugMessage,
				SlidesApp.getUi().ButtonSet.OK,
			);
		}
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			`An error occurred: ${error.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Applies bold style transformation to a shape if its text is wrapped in [brackets].
 * @param {Shape} shape - The shape to apply bold style to
 * @return {Object} Result object with applied status, original text, and new text
 */
function applyBoldStyleTransformation(shape) {
	const result = {
		applied: false,
		originalText: "",
		newText: "",
		debug: {},
	};

	try {
		// Check if the shape has text
		const textRange = shape.getText();
		if (!textRange) {
			result.debug.hasText = false;
			return result;
		}

		// Get the text content, remove line breaks and normalize whitespace
		const textContent = textRange
			.asString()
			.replace(/\r?\n|\r/g, " ")
			.replace(/\s+/g, " ")
			.trim();
		result.originalText = textContent;
		result.debug.hasText = true;
		result.debug.textLength = textContent.length;
		result.debug.startsWithBracket = textContent.startsWith("[");
		result.debug.endsWithBracket = textContent.endsWith("]");

		// Check if text is wrapped in [] and has content between them
		if (
			textContent.startsWith("[") &&
			textContent.endsWith("]") &&
			textContent.length > 2 // Must have at least 1 character between the []
		) {
			// Apply special formatting for [text]
			// Set border with main_color and 1pt weight
			shape.getBorder().setWeight(1);
			shape.getBorder().getLineFill().setSolidFill(main_color);

			// Remove the [] markers and set text
			const cleanText = textContent.substring(1, textContent.length - 1);
			textRange.setText(cleanText);

			// Set text to main_color and bold
			const textStyle = textRange.getTextStyle();
			textStyle.setForegroundColor(main_color);
			textStyle.setBold(true);

			result.newText = cleanText;
			result.applied = true;
			result.debug.transformationApplied = true;
		} else {
			result.debug.transformationApplied = false;
			result.debug.reason = "Text not wrapped in [] or too short";
		}
	} catch (error) {
		console.log(`Error in applyBoldStyleTransformation: ${error.message}`);
		result.debug.error = error.message;
	}

	return result;
}

/**
 * Simple syntax parser: Line 1 = title, Line 2 = parent text, Line 3+ = rows (| separated)
 * Supports [bold] text formatting and (footer) boxes. No complex {[]} syntax needed.
 */
function autoCreateChildShapesFromLines() {
	try {
		// Get the active presentation and selection
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		// Check if a shape is selected
		const selectedElements = selection.getPageElementRange()
			? selection.getPageElementRange().getPageElements()
			: [];

		const selectedShapes = selectedElements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (selectedShapes.length !== 1) {
			SlidesApp.getUi().alert(
				"Error",
				"Please select exactly one shape with line-based text to create child shapes.",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		const parentShape = selectedShapes[0].asShape();

		// Get the raw text content (preserve line breaks)
		const rawTextContent = parentShape.getText().asString();
		const lines = rawTextContent.split(/\r?\n/);

		// Filter out empty lines
		const nonEmptyLines = lines.filter((line) => line.trim() !== "");

		if (nonEmptyLines.length < 3) {
			SlidesApp.getUi().alert(
				"Error",
				"Need at least 3 lines: Line 1 = title, Line 2 = parent text, Line 3+ = rows",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		// Parse the lines
		const titleText = nonEmptyLines[0].trim();
		const parentText = nonEmptyLines[1].trim();
		const rowLines = nonEmptyLines.slice(2);

		// Parse rows (split by | but handle -|> syntax)
		const rowData = rowLines
			.map((line) => parseRowWithHomePlates(line))
			.filter((rowInfo) => rowInfo.cells.length > 0);

		if (rowData.length === 0) {
			SlidesApp.getUi().alert(
				"Error",
				"No valid rows found. Use | to separate columns in each row.",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		// Create the layout structure
		const maxColumns = Math.max(
			...rowData.map((rowInfo) => rowInfo.cells.length),
		);
		const layout = {
			rows: rowData.length,
			maxColumns: maxColumns,
			rowData: rowData,
			isVariableColumns: rowData.some(
				(rowInfo) => rowInfo.cells.length !== maxColumns,
			),
			syntaxType: "line-based",
		};

		// Get the slide
		const slide = selection.getCurrentPage();

		// Create title text box if title exists
		if (titleText) {
			createTitleTextBoxFromText(parentShape, titleText, slide);
		}

		// Set parent text
		parentShape.getText().setText(parentText);
		parentShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
		parentShape.setTitle("PARENT");

		// Create child shapes with default settings
		createChildShapesWithLayout(
			parentShape,
			layout,
			DEFAULT_PADDING,
			DEFAULT_PADDING_TOP,
			DEFAULT_GAP,
		);

		console.log(
			`Auto-created child shapes from ${rowData.length} lines with line-based syntax`,
		);
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			`An error occurred: ${error.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Creates a title RECTANGLE above the parent shape with the given text.
 * @param {Shape} parentShape - The parent shape
 * @param {string} titleText - The title text
 * @param {Slide} slide - The slide to add the rectangle to
 * @return {Shape} The created rectangle
 */
function createTitleTextBoxFromText(parentShape, titleText, slide) {
	// Get parent shape properties
	const parentLeft = parentShape.getLeft();
	const parentTop = parentShape.getTop();
	const parentWidth = parentShape.getWidth();
	const parentRotation = parentShape.getRotation();

	// Create rectangle positioned 30pt above parent shape
	const rectangleLeft = parentLeft;
	const rectangleTop = parentTop - 30; // 30pt above
	const rectangleWidth = parentWidth; // Same width as parent
	const rectangleHeight = 30; // 30pt height

	// Create the rectangle
	const titleRectangle = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		rectangleLeft,
		rectangleTop,
		rectangleWidth,
		rectangleHeight,
	);

	// Apply rotation if parent has any
	if (parentRotation !== 0) {
		titleRectangle.setRotation(parentRotation);
	}

	// Set the fill color to main_color
	const fill = titleRectangle.getFill();
	fill.setSolidFill(main_color);

	// Set white border with 0.1pt weight
	titleRectangle.getBorder().setWeight(0.1);
	titleRectangle.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Set the text content
	titleRectangle.getText().setText(titleText);

	// Style the text - 14pt, bold, white
	const textStyle = titleRectangle.getText().getTextStyle();
	textStyle.setFontSize(14);
	textStyle.setBold(true);
	textStyle.setForegroundColor("#FFFFFF");

	// Center the text vertically and horizontally
	titleRectangle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Bring rectangle forward
	titleRectangle.bringForward();

	console.log(`Created title rectangle with text: "${titleText}"`);

	return titleRectangle;
}

/**
 * Parses a row line to detect -|> syntax for home plates and extract cells.
 * @param {string} line - The row line to parse
 * @return {Object} Object with cells array and homePlates array
 */
function parseRowWithHomePlates(line) {
	// Split by | but keep track of home plate positions
	const parts = line.split("|");
	const cells = [];
	const homePlates = [];

	for (let i = 0; i < parts.length; i++) {
		const part = parts[i].trim();

		// Check if this part ends with - (indicating start of home plate)
		if (part.endsWith("-") && i < parts.length - 1) {
			// Next part should start with > for complete -|> syntax
			const nextPart = parts[i + 1].trim();
			if (nextPart.startsWith(">")) {
				// This is a home plate position
				// Add the cell without the trailing -
				const cellText = part.slice(0, -1).trim();
				if (cellText) {
					cells.push(cellText);
				}

				// Record home plate position (after current cell)
				homePlates.push(cells.length);

				// Add the next cell without the leading >
				const nextCellText = nextPart.slice(1).trim();
				if (nextCellText) {
					cells.push(nextCellText);
				}

				// Skip the next part since we already processed it
				i++;
			} else {
				// Not a home plate, just a regular cell
				if (part) {
					cells.push(part);
				}
			}
		} else {
			// Regular cell
			if (part) {
				cells.push(part);
			}
		}
	}

	return {
		cells: cells.filter((cell) => cell !== ""),
		homePlates: homePlates,
	};
}

/**
 * Processes cell text to extract main content and footer box text.
 * Detects (text) pattern for footer boxes.
 * @param {string} cellText - The original cell text
 * @return {Object} Object with hasFooter, mainText, and footerText properties
 */
function processFooterBoxText(cellText) {
	// Check if text contains footer pattern (text)
	const footerRegex = /\(([^)]*)\)/;
	const match = cellText.match(footerRegex);

	if (match) {
		const footerText = match[1]; // Text inside parentheses
		const mainText = cellText.replace(footerRegex, "").trim(); // Text without parentheses

		return {
			hasFooter: true,
			mainText: mainText,
			footerText: footerText,
		};
	}

	return {
		hasFooter: false,
		mainText: cellText,
		footerText: "",
	};
}

/**
 * Creates a footer box at the bottom of a cell.
 * @param {Slide} slide - The slide to add the footer box to
 * @param {number} left - Left position of the footer box
 * @param {number} top - Top position of the footer box
 * @param {number} width - Width of the footer box
 * @param {number} height - Height of the footer box
 * @param {string} text - Text content for the footer box
 * @param {number} rotation - Rotation angle to apply
 * @return {Shape} The created footer box shape
 */
function createFooterBox(slide, left, top, width, height, text, rotation) {
	// Create the footer box rectangle
	const footerBox = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		left,
		top,
		width,
		height,
	);

	// Apply rotation if needed
	if (rotation !== 0) {
		footerBox.setRotation(rotation);
	}

	// Set background color to main_color
	const fill = footerBox.getFill();
	fill.setSolidFill(main_color);

	// Set white border with 0.1pt weight
	footerBox.getBorder().setWeight(0.1);
	footerBox.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Set the text content
	footerBox.getText().setText(text);

	// Style the text - white color, 10pt font size, centered
	const textStyle = footerBox.getText().getTextStyle();
	textStyle.setForegroundColor("#FFFFFF");
	textStyle.setFontSize(10);

	// Center the text both vertically and horizontally
	footerBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Bring footer box forward
	footerBox.bringForward();

	console.log(`Created footer box with text: "${text}"`);

	return footerBox;
}

/**
 * Applies white fill and white stroke to a shape.
 * If the text is wrapped in **asterisks**, applies special formatting.
 * @param {Shape} shape - The shape to apply white style to.
 */
function applyWhiteStyle(shape) {
	try {
		// Always set white fill first
		const fill = shape.getFill();
		fill.setSolidFill("#FFFFFF");

		// Check if we should apply bold style transformation
		const result = applyBoldStyleTransformation(shape);

		if (!result.applied) {
			// Apply normal white style if no bold transformation was applied
			// Set white border with 0.1pt weight
			shape.getBorder().setWeight(0.1);
			shape.getBorder().getLineFill().setSolidFill("#FFFFFF");

			// Optionally set text color to black for visibility on white background
			if (shape.getText()) {
				const textStyle = shape.getText().getTextStyle();
				textStyle.setForegroundColor("#000000");
			}
		}
	} catch (error) {
		// Log error but continue execution
		console.log(`Error applying white style: ${error.message}`);
	}
}
