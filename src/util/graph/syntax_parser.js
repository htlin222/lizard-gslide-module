/**
 * Syntax Parser for Child Shapes
 * Handles text parsing and grid layout logic
 */

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
		const DEFAULT_PADDING = 10;
		const DEFAULT_PADDING_TOP = 30;
		const DEFAULT_GAP = 10;

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
