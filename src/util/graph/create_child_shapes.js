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
 * Automatically creates child shapes supporting nested syntax:
 * - Single level: {[item1|item2][item3|item4]} - creates grid directly
 * - Nested level: {grid1} {grid2} {grid3} - splits into columns, then creates grids in each
 */
function autoCreateChildShapesFromText() {
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
				"Please select exactly one shape with text syntax to auto-create child shapes.",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		const parentShape = selectedShapes[0].asShape();

		// Get the text content from the shape and remove line breaks to avoid syntax errors
		const rawTextContent = parentShape.getText().asString();
		const textContent = rawTextContent.replace(/\r?\n|\r/g, " ").replace(/\s+/g, " ").trim();

		// Check if this is a nested syntax (multiple {} blocks with complex content)
		const nestedLayout = parseNestedSyntax(textContent);

		if (nestedLayout) {
			// Handle nested syntax: split into columns first, then create grids in each column
			createNestedChildShapes(parentShape, nestedLayout);
			console.log(
				`Auto-created nested layout with ${nestedLayout.columns} columns`,
			);
			return;
		}

		// Try single-level parsing
		const gridLayout = parseGridSyntax(textContent);

		if (!gridLayout) {
			SlidesApp.getUi().alert(
				"Error",
				"No valid grid syntax found. Please use format: {[item1|item2][item3|item4]} or nested format",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		// Create child shapes with default settings
		const defaultPadding = 7;
		const defaultPaddingTop = 30;
		const defaultGap = 7;

		createChildShapesWithLayout(
			parentShape,
			gridLayout,
			defaultPadding,
			defaultPaddingTop,
			defaultGap,
		);

		console.log(
			`Auto-created child shapes with ${gridLayout.rows} rows and varying columns`,
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
		const row = layout.rowData[rowIndex];
		const columnsInRow = row.length;

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

			// Set the text content from the parsed data
			const cellText = row[colIndex].trim();
			if (cellText) {
				const textRange = childShape.getText();
				textRange.setText(cellText);
			}

			// Apply styling after text is set (so we can check for ** markers)
			applyWhiteStyle(childShape);
			
			// Set title for child shape
			childShape.setTitle("CHILD");

			childShapes.push(childShape);
		}
	}

	// Bring all child shapes forward
	for (const childShape of childShapes) {
		childShape.bringForward();
	}

	// Set parent shape text alignment to top and set text to content outside {}
	parentShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
	
	// Set title for parent shape
	parentShape.setTitle("PARENT");

	// Extract text outside {} brackets
	const originalText = parentShape.getText().asString();
	const textOutsideBrackets = extractTextOutsideBrackets(originalText);

	// Check if text outside brackets is wrapped in quotes "text"
	const titleTextBox = createTitleTextBoxIfNeeded(
		parentShape,
		textOutsideBrackets,
		slide,
	);

	// Set remaining text to parent shape (after removing quoted text)
	const remainingText = titleTextBox
		? extractTextWithoutQuotes(textOutsideBrackets)
		: textOutsideBrackets;
	parentShape.getText().setText(remainingText);

	console.log(
		`Created ${childShapes.length} child shapes with variable column layout`,
	);
}

/**
 * Parses nested syntax where multiple {} blocks contain complex grid definitions.
 * Example: {[A|B][C|D]} {[E|F][G|H]} {[I|J][K|L]}
 * @param {string} text - The text to parse
 * @return {Object|null} Nested layout object or null if not nested syntax
 */
function parseNestedSyntax(text) {
	// Look for multiple {} blocks that contain [...] patterns (complex grids)
	const complexBlockRegex = /\{(\[.*?\])+\}/g;
	const complexBlocks = [];
	let match;

	while ((match = complexBlockRegex.exec(text)) !== null) {
		complexBlocks.push(match[0]);
	}

	// Only consider it nested if we have multiple complex blocks
	if (complexBlocks.length < 2) {
		return null;
	}

	// Parse each block's content
	const columnLayouts = [];
	for (const block of complexBlocks) {
		const layout = parseMultiRowSyntax(block);
		if (layout) {
			columnLayouts.push({
				content: block,
				layout: layout,
			});
		}
	}

	if (columnLayouts.length === 0) {
		return null;
	}

	return {
		columns: columnLayouts.length,
		columnLayouts: columnLayouts,
		syntaxType: "nested",
	};
}

/**
 * Creates nested child shapes: first splits into columns, then creates grids in each column.
 * @param {Shape} parentShape - The parent shape
 * @param {Object} nestedLayout - The nested layout structure
 */
function createNestedChildShapes(parentShape, nestedLayout) {
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

	// Split into columns first (like split_shape.js)
	const columnCount = nestedLayout.columns;
	const columnGap = 7; // Gap between main columns
	const columnPadding = 7; // Padding around each column

	// Calculate column dimensions
	const availableWidth = parentWidth - columnPadding * 2;
	const columnWidth =
		(availableWidth - columnGap * (columnCount - 1)) / columnCount;

	if (columnWidth <= 0) {
		throw new Error("Parent shape is too small for the number of columns.");
	}

	// Create a column shape for each layout
	for (let colIndex = 0; colIndex < columnCount; colIndex++) {
		const columnLayout = nestedLayout.columnLayouts[colIndex];

		// Calculate column position
		const columnLeft =
			parentLeft + columnPadding + colIndex * (columnWidth + columnGap);
		const columnTop = parentTop;

		// Create a temporary column shape to hold the grid
		const columnShape = slide.insertShape(
			parentShape.getShapeType(),
			columnLeft,
			columnTop,
			columnWidth,
			parentHeight,
		);

		// Apply rotation if needed
		if (parentRotation !== 0) {
			columnShape.setRotation(parentRotation);
		}

		// Apply white styling to column
		applyWhiteStyle(columnShape);
		
		// Set title for column shape (which acts as a parent for nested grids)
		columnShape.setTitle("PARENT");

		// Now create the grid inside this column using the parsed layout
		const gridPadding = 3; // Smaller padding for nested grids
		const gridPaddingTop = 15; // Smaller top padding for nested grids
		const gridGap = 3; // Smaller gaps for nested grids

		createChildShapesWithLayout(
			columnShape,
			columnLayout.layout,
			gridPadding,
			gridPaddingTop,
			gridGap,
		);

		// The column shape is now the container, bring it forward
		columnShape.bringForward();
	}

	console.log(`Created ${columnCount} column layout with nested grids`);
}

/**
 * Extracts text that is outside of {} brackets.
 * @param {string} text - The original text
 * @return {string} Text outside brackets, trimmed
 */
function extractTextOutsideBrackets(text) {
	// Remove line breaks and normalize whitespace
	let result = text.replace(/\r?\n|\r/g, " ");

	// Remove all {} blocks (including nested ones)
	// Remove multi-row format: {[...][...]}
	result = result.replace(/\{(\[.*?\])+\}/g, "");

	// Remove single-row format: {content}
	result = result.replace(/\{[^}]*\}/g, "");

	// Clean up extra whitespace and return
	return result.replace(/\s+/g, " ").trim();
}

/**
 * Creates a title TEXT_BOX above the parent shape if the text is wrapped in quotes "text".
 * @param {Shape} parentShape - The parent shape
 * @param {string} textOutsideBrackets - The text outside {} brackets
 * @param {Slide} slide - The slide to add the text box to
 * @return {Shape|null} The created text box or null if no quoted text found
 */
function createTitleTextBoxIfNeeded(parentShape, textOutsideBrackets, slide) {
	// Check if text contains quoted content "text"
	const quotedTextRegex = /"([^"]*)"/;
	const match = textOutsideBrackets.match(quotedTextRegex);

	if (!match) {
		return null; // No quoted text found
	}

	const quotedText = match[1]; // Text inside quotes

	// Get parent shape properties
	const parentLeft = parentShape.getLeft();
	const parentTop = parentShape.getTop();
	const parentWidth = parentShape.getWidth();
	const parentRotation = parentShape.getRotation();

	// Create text box positioned 30pt above parent shape with same width and 30pt height
	const textBoxLeft = parentLeft;
	const textBoxTop = parentTop - 30; // 30pt above
	const textBoxWidth = parentWidth; // Same width as parent
	const textBoxHeight = 30; // 30pt height

	// Create the text box
	const titleTextBox = slide.insertShape(
		SlidesApp.ShapeType.TEXT_BOX,
		textBoxLeft,
		textBoxTop,
		textBoxWidth,
		textBoxHeight,
	);

	// Apply rotation if parent has any
	if (parentRotation !== 0) {
		titleTextBox.setRotation(parentRotation);
	}

	// Set the text content
	titleTextBox.getText().setText(quotedText);

	// Style the text - 14pt, bold, main_color
	const textStyle = titleTextBox.getText().getTextStyle();
	textStyle.setFontSize(14);
	textStyle.setBold(true);
	textStyle.setForegroundColor(main_color);

	// Style the text box - TEXT_BOX is already transparent by default
	// Don't set border weight to 0 as it causes an error, leave it as default

	// Bring text box forward
	titleTextBox.bringForward();

	console.log(`Created title text box with text: "${quotedText}"`);

	return titleTextBox;
}

/**
 * Extracts text without the quoted portions "text".
 * @param {string} text - The original text
 * @return {string} Text with quoted portions removed
 */
function extractTextWithoutQuotes(text) {
	// Remove line breaks and normalize whitespace
	let result = text.replace(/\r?\n|\r/g, " ");

	// Remove all quoted text "content"
	result = result.replace(/"[^"]*"/g, "");

	// Clean up extra whitespace and return
	return result.replace(/\s+/g, " ").trim();
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
 * Applies bold style transformation to a shape if its text is wrapped in **asterisks**.
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
		result.debug.startsWithAsterisk = textContent.startsWith("**");
		result.debug.endsWithAsterisk = textContent.endsWith("**");

		// Check if text is wrapped in ** and has content between them
		if (
			textContent.startsWith("**") &&
			textContent.endsWith("**") &&
			textContent.length > 4 // Must have at least 1 character between the **
		) {
			// Apply special formatting for **text**
			// Set border with main_color and 1.5pt weight
			shape.getBorder().setWeight(1);
			shape.getBorder().getLineFill().setSolidFill(main_color);

			// Remove the ** markers and set text
			const cleanText = textContent.substring(2, textContent.length - 2);
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
			result.debug.reason = "Text not wrapped in ** or too short";
		}
	} catch (error) {
		console.log(`Error in applyBoldStyleTransformation: ${error.message}`);
		result.debug.error = error.message;
	}

	return result;
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
			// Set white border following the documented approach in basic_style_api.md
			// First set the border weight
			shape.getBorder().setWeight(1);

			// Then set the border color using getLineFill() - the correct documented way
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
