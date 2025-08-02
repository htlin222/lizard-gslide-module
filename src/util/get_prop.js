/**
 * Utility functions to get and log properties of selected objects in Google Slides
 */

/**
 * Logs all accessible properties of the currently selected object in Google Slides
 * This function detects the type of selection and logs appropriate properties
 */
function logSelectedObjectProperties() {
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectionType = selection.getSelectionType();

	Logger.log("Selection Type: " + selectionType);
	console.log("Selection Type: " + selectionType);

	// Different types of selections
	switch (selectionType) {
		case SlidesApp.SelectionType.TEXT:
			logTextSelectionProperties(selection);
			break;

		case SlidesApp.SelectionType.TABLE_CELL:
			logTableCellProperties(selection);
			break;

		case SlidesApp.SelectionType.PAGE_ELEMENT:
			logPageElementProperties(selection);
			break;

		case SlidesApp.SelectionType.CURRENT_PAGE:
			logCurrentPageProperties(selection);
			break;

		default:
			Logger.log("No supported selection type found or nothing is selected");
			console.log("No supported selection type found or nothing is selected");
	}
}

/**
 * Shows a dialog with properties of the currently selected object in Google Slides
 * This function detects the type of selection and displays appropriate properties in a dialog
 */
function showSelectedObjectPropertiesDialog() {
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectionType = selection.getSelectionType();

	let propertiesHtml = "<style>";
	propertiesHtml += "body { font-family: Arial, sans-serif; padding: 20px; }";
	propertiesHtml += "h2 { color: #3D6869; margin-bottom: 10px; }";
	propertiesHtml += "h3 { color: #555; margin-top: 15px; margin-bottom: 5px; }";
	propertiesHtml +=
		".property { margin: 5px 0; padding: 5px 10px; background: #f5f5f5; border-radius: 4px; }";
	propertiesHtml += ".property-name { font-weight: bold; color: #333; }";
	propertiesHtml += ".property-value { color: #666; margin-left: 10px; }";
	propertiesHtml +=
		".no-selection { color: #999; font-style: italic; padding: 20px; text-align: center; }";
	propertiesHtml += "</style>";

	propertiesHtml += "<h2>Selected Object Properties</h2>";
	propertiesHtml +=
		'<div class="property"><span class="property-name">Selection Type:</span><span class="property-value">' +
		selectionType +
		"</span></div>";

	// Different types of selections
	switch (selectionType) {
		case SlidesApp.SelectionType.TEXT:
			propertiesHtml += getTextSelectionPropertiesHtml(selection);
			break;

		case SlidesApp.SelectionType.TABLE_CELL:
			propertiesHtml += getTableCellPropertiesHtml(selection);
			break;

		case SlidesApp.SelectionType.PAGE_ELEMENT:
			propertiesHtml += getPageElementPropertiesHtml(selection);
			break;

		case SlidesApp.SelectionType.CURRENT_PAGE:
			propertiesHtml += getCurrentPagePropertiesHtml(selection);
			break;

		default:
			propertiesHtml +=
				'<div class="no-selection">No supported selection type found or nothing is selected</div>';
	}

	// Create and show the dialog
	const htmlOutput = HtmlService.createHtmlOutput(propertiesHtml)
		.setTitle("Object Properties")
		.setWidth(600)
		.setHeight(500);

	SlidesApp.getUi().showModelessDialog(htmlOutput, "Object Properties");
}

/**
 * Gets HTML representation of text selection properties
 * @param {Selection} selection - The current selection object
 * @return {string} HTML string with properties
 */
function getTextSelectionPropertiesHtml(selection) {
	const textRange = selection.getTextRange();
	const textStyle = textRange.getTextStyle();

	let html = "<h3>Text Selection Properties</h3>";
	html +=
		'<div class="property"><span class="property-name">Text content:</span><span class="property-value">' +
		textRange.getText() +
		"</span></div>";

	// Text style properties
	html +=
		'<div class="property"><span class="property-name">Bold:</span><span class="property-value">' +
		textStyle.isBold() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Italic:</span><span class="property-value">' +
		textStyle.isItalic() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Underline:</span><span class="property-value">' +
		textStyle.isUnderline() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Strikethrough:</span><span class="property-value">' +
		textStyle.isStrikethrough() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Font family:</span><span class="property-value">' +
		textStyle.getFontFamily() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Font size:</span><span class="property-value">' +
		textStyle.getFontSize() +
		"</span></div>";

	// Try to get foreground color
	try {
		const foregroundColor = textStyle.getForegroundColor();
		if (foregroundColor) {
			html +=
				'<div class="property"><span class="property-name">Foreground color:</span><span class="property-value">' +
				JSON.stringify(foregroundColor) +
				"</span></div>";
		}
	} catch (e) {}

	// Get paragraph style if available
	try {
		const paragraphStyle = textRange.getParagraphStyle();
		html +=
			'<div class="property"><span class="property-name">Alignment:</span><span class="property-value">' +
			paragraphStyle.getAlignment() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Line spacing:</span><span class="property-value">' +
			paragraphStyle.getLineSpacing() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Space above:</span><span class="property-value">' +
			paragraphStyle.getSpaceAbove() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Space below:</span><span class="property-value">' +
			paragraphStyle.getSpaceBelow() +
			"</span></div>";
	} catch (e) {}

	return html;
}

/**
 * Logs properties of a text selection
 * @param {Selection} selection - The current selection object
 */
function logTextSelectionProperties(selection) {
	const textRange = selection.getTextRange();
	const textStyle = textRange.getTextStyle();

	Logger.log("=== TEXT SELECTION PROPERTIES ===");
	Logger.log("Text content: " + textRange.getText());

	// Text style properties
	Logger.log("Bold: " + textStyle.isBold());
	Logger.log("Italic: " + textStyle.isItalic());
	Logger.log("Underline: " + textStyle.isUnderline());
	Logger.log("Strikethrough: " + textStyle.isStrikethrough());
	Logger.log("Font family: " + textStyle.getFontFamily());
	Logger.log("Font size: " + textStyle.getFontSize());

	// Try to get foreground color (may be null if mixed)
	try {
		const foregroundColor = textStyle.getForegroundColor();
		if (foregroundColor) {
			Logger.log("Foreground color: " + JSON.stringify(foregroundColor));
		}
	} catch (e) {
		Logger.log("Could not get foreground color: " + e.message);
	}

	// Get paragraph style if available
	try {
		const paragraphStyle = textRange.getParagraphStyle();
		Logger.log("Alignment: " + paragraphStyle.getAlignment());
		Logger.log("Line spacing: " + paragraphStyle.getLineSpacing());
		Logger.log("Space above: " + paragraphStyle.getSpaceAbove());
		Logger.log("Space below: " + paragraphStyle.getSpaceBelow());
	} catch (e) {
		Logger.log("Could not get paragraph style: " + e.message);
	}

	// Also log to console for Apps Script dashboard
	console.log("=== TEXT SELECTION PROPERTIES ===");
	console.log(
		JSON.stringify(
			{
				content: textRange.getText(),
				isBold: textStyle.isBold(),
				isItalic: textStyle.isItalic(),
				isUnderline: textStyle.isUnderline(),
				fontFamily: textStyle.getFontFamily(),
				fontSize: textStyle.getFontSize(),
			},
			null,
			2,
		),
	);
}

/**
 * Gets HTML representation of table cell selection properties
 * @param {Selection} selection - The current selection object
 * @return {string} HTML string with properties
 */
function getTableCellPropertiesHtml(selection) {
	const tableCellRange = selection.getTableCellRange();
	const cells = tableCellRange.getTableCells();

	let html = "<h3>Table Cell Properties</h3>";
	html +=
		'<div class="property"><span class="property-name">Number of selected cells:</span><span class="property-value">' +
		cells.length +
		"</span></div>";

	if (cells.length > 0) {
		const cell = cells[0];
		const table = cell.getParentTable();

		html +=
			'<div class="property"><span class="property-name">Table dimensions:</span><span class="property-value">' +
			table.getNumRows() +
			" rows × " +
			table.getNumColumns() +
			" columns</span></div>";
		html +=
			'<div class="property"><span class="property-name">Selected cell position:</span><span class="property-value">Row ' +
			cell.getRowIndex() +
			", Column " +
			cell.getColumnIndex() +
			"</span></div>";

		// Cell content
		html +=
			'<div class="property"><span class="property-name">Cell content:</span><span class="property-value">' +
			cell.getText().asString() +
			"</span></div>";

		// Cell appearance
		try {
			const fill = cell.getFill();
			if (fill.getSolidFill()) {
				const color = fill.getSolidFill().getColor();
				html +=
					'<div class="property"><span class="property-name">Cell background color:</span><span class="property-value">' +
					JSON.stringify(color) +
					"</span></div>";
			}
		} catch (e) {}
	}

	return html;
}

/**
 * Logs properties of a table cell selection
 * @param {Selection} selection - The current selection object
 */
function logTableCellProperties(selection) {
	const tableCellRange = selection.getTableCellRange();
	const cells = tableCellRange.getTableCells();

	Logger.log("=== TABLE CELL PROPERTIES ===");
	Logger.log("Number of selected cells: " + cells.length);

	if (cells.length > 0) {
		const cell = cells[0];
		const table = cell.getParentTable();

		Logger.log(
			"Table dimensions: " +
				table.getNumRows() +
				" rows × " +
				table.getNumColumns() +
				" columns",
		);
		Logger.log(
			"Selected cell position: Row " +
				cell.getRowIndex() +
				", Column " +
				cell.getColumnIndex(),
		);

		// Cell content
		Logger.log("Cell content: " + cell.getText().asString());

		// Cell appearance
		try {
			const fill = cell.getFill();
			if (fill.getSolidFill()) {
				const color = fill.getSolidFill().getColor();
				Logger.log("Cell background color: " + JSON.stringify(color));
			}
		} catch (e) {
			Logger.log("Could not get cell fill: " + e.message);
		}

		// Also log to console
		console.log("=== TABLE CELL PROPERTIES ===");
		console.log(
			JSON.stringify(
				{
					tableSize: {
						rows: table.getNumRows(),
						columns: table.getNumColumns(),
					},
					selectedCell: {
						row: cell.getRowIndex(),
						column: cell.getColumnIndex(),
						content: cell.getText().asString(),
					},
				},
				null,
				2,
			),
		);
	}
}

/**
 * Gets HTML representation of page element selection properties
 * @param {Selection} selection - The current selection object
 * @return {string} HTML string with properties
 */
function getPageElementPropertiesHtml(selection) {
	const pageElementRange = selection.getPageElementRange();
	const pageElements = pageElementRange.getPageElements();

	let html = "<h3>Page Element Properties</h3>";
	html +=
		'<div class="property"><span class="property-name">Number of selected elements:</span><span class="property-value">' +
		pageElements.length +
		"</span></div>";

	if (pageElements.length > 0) {
		const element = pageElements[0];
		const elementType = element.getPageElementType();

		html +=
			'<div class="property"><span class="property-name">Element type:</span><span class="property-value">' +
			elementType +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Element ID:</span><span class="property-value">' +
			element.getObjectId() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Position:</span><span class="property-value">Left ' +
			element.getLeft() +
			", Top " +
			element.getTop() +
			"</span></div>";

		// Try to get size
		let width = "null";
		let height = "null";
		try {
			width = element.getWidth();
			height = element.getHeight();
		} catch (e) {}
		html +=
			'<div class="property"><span class="property-name">Size:</span><span class="property-value">Width ' +
			width +
			", Height " +
			height +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Rotation:</span><span class="property-value">' +
			element.getRotation() +
			" degrees</span></div>";

		// Type-specific properties
		switch (elementType) {
			case SlidesApp.PageElementType.SHAPE:
				html += getShapePropertiesHtml(element.asShape());
				break;
			case SlidesApp.PageElementType.IMAGE:
				html += getImagePropertiesHtml(element.asImage());
				break;
			case SlidesApp.PageElementType.TABLE:
				html += getTablePropertiesHtml(element.asTable());
				break;
			case SlidesApp.PageElementType.GROUP:
				html +=
					'<div class="property"><span class="property-name">Group contains:</span><span class="property-value">' +
					element.asGroup().getChildren().length +
					" elements</span></div>";
				break;
		}
	}

	return html;
}

/**
 * Logs properties of a page element selection (shapes, images, etc.)
 * @param {Selection} selection - The current selection object
 */
function logPageElementProperties(selection) {
	const pageElementRange = selection.getPageElementRange();
	const pageElements = pageElementRange.getPageElements();

	Logger.log("=== PAGE ELEMENT PROPERTIES ===");
	Logger.log("Number of selected elements: " + pageElements.length);

	if (pageElements.length > 0) {
		const element = pageElements[0];
		const elementType = element.getPageElementType();

		Logger.log("Element type: " + elementType);
		Logger.log("Element ID: " + element.getObjectId());
		Logger.log(
			"Position: Left " + element.getLeft() + ", Top " + element.getTop(),
		);

		// Try to get size - might be null for some elements
		let width = "null";
		let height = "null";
		try {
			width = element.getWidth();
			height = element.getHeight();
		} catch (e) {}
		Logger.log("Size: Width " + width + ", Height " + height);
		Logger.log("Rotation: " + element.getRotation() + " degrees");

		// Log all available methods on the element
		const methods = [];
		for (const prop in element) {
			if (typeof element[prop] === "function" && prop[0] !== "_") {
				methods.push(prop);
			}
		}
		Logger.log("Available methods: " + methods.join(", "));

		// Type-specific properties
		switch (elementType) {
			case SlidesApp.PageElementType.SHAPE:
				logShapeProperties(element.asShape());
				break;
			case SlidesApp.PageElementType.IMAGE:
				logImageProperties(element.asImage());
				break;
			case SlidesApp.PageElementType.TABLE:
				logTableProperties(element.asTable());
				break;
			case SlidesApp.PageElementType.GROUP:
				Logger.log(
					"Group contains " +
						element.asGroup().getChildren().length +
						" elements",
				);
				break;
		}

		// Also log to console
		console.log("=== PAGE ELEMENT PROPERTIES ===");
		console.log(
			JSON.stringify(
				{
					type: elementType,
					id: element.getObjectId(),
					position: {
						left: element.getLeft(),
						top: element.getTop(),
					},
					size: {
						width: width,
						height: height,
					},
					rotation: element.getRotation(),
					availableMethods: methods,
				},
				null,
				2,
			),
		);
	}
}

/**
 * Gets HTML representation of shape properties
 * @param {Shape} shape - The shape object
 * @return {string} HTML string with properties
 */
function getShapePropertiesHtml(shape) {
	let html = '<h4 style="color: #777; margin-top: 10px;">Shape Properties</h4>';
	html +=
		'<div class="property"><span class="property-name">Shape type:</span><span class="property-value">' +
		shape.getShapeType() +
		"</span></div>";

	// Fill
	try {
		const fill = shape.getFill();
		if (fill.getSolidFill()) {
			const color = fill.getSolidFill().getColor();
			html +=
				'<div class="property"><span class="property-name">Fill color:</span><span class="property-value">' +
				JSON.stringify(color) +
				"</span></div>";
		}
	} catch (e) {}

	// Border
	try {
		const border = shape.getBorder();
		if (border) {
			html +=
				'<div class="property"><span class="property-name">Border weight:</span><span class="property-value">' +
				border.getWeight() +
				"</span></div>";
			html +=
				'<div class="property"><span class="property-name">Border dash style:</span><span class="property-value">' +
				border.getDashStyle() +
				"</span></div>";
		}
	} catch (e) {}

	// Text content if any
	if (shape.getText()) {
		html +=
			'<div class="property"><span class="property-name">Shape text:</span><span class="property-value">' +
			shape.getText().asString() +
			"</span></div>";
	}

	return html;
}

/**
 * Logs properties specific to shapes
 * @param {Shape} shape - The shape object
 */
function logShapeProperties(shape) {
	Logger.log("--- Shape Properties ---");
	Logger.log("Shape type: " + shape.getShapeType());

	// Fill
	try {
		const fill = shape.getFill();
		if (fill.getSolidFill()) {
			const color = fill.getSolidFill().getColor();
			Logger.log("Fill color: " + JSON.stringify(color));
		}
	} catch (e) {
		Logger.log("Could not get shape fill: " + e.message);
	}

	// Border
	try {
		const border = shape.getBorder();
		if (border) {
			Logger.log("Border weight: " + border.getWeight());
			Logger.log("Border dash style: " + border.getDashStyle());
		}
	} catch (e) {
		Logger.log("Could not get shape border: " + e.message);
	}

	// Text content if any
	if (shape.getText()) {
		Logger.log("Shape text: " + shape.getText().asString());
	}
}

/**
 * Gets HTML representation of image properties
 * @param {Image} image - The image object
 * @return {string} HTML string with properties
 */
function getImagePropertiesHtml(image) {
	let html = '<h4 style="color: #777; margin-top: 10px;">Image Properties</h4>';
	html +=
		'<div class="property"><span class="property-name">Content URL:</span><span class="property-value">' +
		image.getContentUrl() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Source URL:</span><span class="property-value">' +
		image.getSourceUrl() +
		"</span></div>";

	// Try to get image properties that might not be available
	try {
		html +=
			'<div class="property"><span class="property-name">Brightness:</span><span class="property-value">' +
			image.getBrightness() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Contrast:</span><span class="property-value">' +
			image.getContrast() +
			"</span></div>";
		html +=
			'<div class="property"><span class="property-name">Transparency:</span><span class="property-value">' +
			image.getTransparency() +
			"</span></div>";
	} catch (e) {}

	return html;
}

/**
 * Logs properties specific to images
 * @param {Image} image - The image object
 */
function logImageProperties(image) {
	Logger.log("--- Image Properties ---");
	Logger.log("Content URL: " + image.getContentUrl());
	Logger.log("Source URL: " + image.getSourceUrl());

	// Try to get image properties that might not be available
	try {
		Logger.log("Brightness: " + image.getBrightness());
		Logger.log("Contrast: " + image.getContrast());
		Logger.log("Transparency: " + image.getTransparency());
	} catch (e) {
		Logger.log("Could not get some image properties: " + e.message);
	}
}

/**
 * Gets HTML representation of table properties
 * @param {Table} table - The table object
 * @return {string} HTML string with properties
 */
function getTablePropertiesHtml(table) {
	let html = '<h4 style="color: #777; margin-top: 10px;">Table Properties</h4>';
	html +=
		'<div class="property"><span class="property-name">Table dimensions:</span><span class="property-value">' +
		table.getNumRows() +
		" rows × " +
		table.getNumColumns() +
		" columns</span></div>";

	// Show a sample of table data
	html +=
		'<div style="margin-top: 10px; font-weight: bold;">Table data sample (first 3x3):</div>';
	html += '<table style="border-collapse: collapse; margin-top: 5px;">';

	for (let r = 0; r < Math.min(table.getNumRows(), 3); r++) {
		html += "<tr>";
		for (let c = 0; c < Math.min(table.getNumColumns(), 3); c++) {
			try {
				const cell = table.getCell(r, c);
				const cellText = cell.getText().asString();
				html +=
					'<td style="border: 1px solid #ccc; padding: 5px;">' +
					(cellText || "&nbsp;") +
					"</td>";
			} catch (e) {
				html +=
					'<td style="border: 1px solid #ccc; padding: 5px; color: #999;">Error</td>';
			}
		}
		html += "</tr>";
	}

	html += "</table>";

	return html;
}

/**
 * Logs properties specific to tables
 * @param {Table} table - The table object
 */
function logTableProperties(table) {
	Logger.log("--- Table Properties ---");
	Logger.log(
		"Table dimensions: " +
			table.getNumRows() +
			" rows × " +
			table.getNumColumns() +
			" columns",
	);

	// Log table structure and content
	const tableData = [];
	for (let r = 0; r < Math.min(table.getNumRows(), 5); r++) {
		// Limit to first 5 rows
		const rowData = [];
		for (let c = 0; c < Math.min(table.getNumColumns(), 5); c++) {
			// Limit to first 5 columns
			try {
				const cell = table.getCell(r, c);
				const cellText = cell.getText().asString();

				// Try to get cell background color
				let bgColor = "unknown";
				try {
					const fill = cell.getFill();
					if (fill.getSolidFill()) {
						const color = fill.getSolidFill().getColor();
						if (color) {
							bgColor = JSON.stringify(color);
						}
					}
				} catch (e) {}

				rowData.push({
					text: cellText,
					bgColor: bgColor,
				});
			} catch (e) {
				rowData.push({ error: e.message });
			}
		}
		tableData.push(rowData);
	}

	Logger.log("Table data sample: " + JSON.stringify(tableData, null, 2));

	// Try to get border properties
	try {
		const cell = table.getCell(0, 0);
		const border = cell.getBorder();
		if (border) {
			Logger.log("Border properties available: " + (border !== null));
		}
	} catch (e) {
		Logger.log("Could not get border properties: " + e.message);
	}

	// Log table object properties
	const tableProps = {};
	for (const prop in table) {
		if (typeof table[prop] === "function") {
			tableProps[prop] = "function";
		} else if (prop[0] !== "_") {
			// Skip private properties
			tableProps[prop] = table[prop];
		}
	}
	Logger.log(
		"Available table properties: " + JSON.stringify(tableProps, null, 2),
	);
}

/**
 * Gets HTML representation of current page properties
 * @param {Selection} selection - The current selection object
 * @return {string} HTML string with properties
 */
function getCurrentPagePropertiesHtml(selection) {
	const currentPage = selection.getCurrentPage();

	let html = "<h3>Current Page Properties</h3>";
	html +=
		'<div class="property"><span class="property-name">Page ID:</span><span class="property-value">' +
		currentPage.getObjectId() +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Page type:</span><span class="property-value">' +
		(currentPage.isGrouped() ? "Group" : "Regular") +
		"</span></div>";
	html +=
		'<div class="property"><span class="property-name">Page elements count:</span><span class="property-value">' +
		currentPage.getPageElements().length +
		"</span></div>";

	// Background
	try {
		const background = currentPage.getBackground();
		if (background.getSolidFill()) {
			const color = background.getSolidFill().getColor();
			html +=
				'<div class="property"><span class="property-name">Background color:</span><span class="property-value">' +
				JSON.stringify(color) +
				"</span></div>";
		}
	} catch (e) {}

	return html;
}

/**
 * Logs properties of the current page
 * @param {Selection} selection - The current selection object
 */
function logCurrentPageProperties(selection) {
	const currentPage = selection.getCurrentPage();

	Logger.log("=== CURRENT PAGE PROPERTIES ===");
	Logger.log("Page ID: " + currentPage.getObjectId());
	Logger.log("Page index: " + currentPage.getObjectId());
	Logger.log("Page type: " + (currentPage.isGrouped() ? "Group" : "Regular"));
	Logger.log("Page elements count: " + currentPage.getPageElements().length);

	// Background
	try {
		const background = currentPage.getBackground();
		if (background.getSolidFill()) {
			const color = background.getSolidFill().getColor();
			Logger.log("Background color: " + JSON.stringify(color));
		}
	} catch (e) {
		Logger.log("Could not get page background: " + e.message);
	}

	// Also log to console
	console.log("=== CURRENT PAGE PROPERTIES ===");
	console.log(
		JSON.stringify(
			{
				id: currentPage.getObjectId(),
				elementsCount: currentPage.getPageElements().length,
			},
			null,
			2,
		),
	);
}

/**
 * Specifically logs the structure of a selected table with more detail
 */
function logSelectedTableStructure() {
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectionType = selection.getSelectionType();

	if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
		const pageElements = selection.getPageElementRange().getPageElements();
		if (pageElements.length > 0) {
			const element = pageElements[0];
			if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
				const table = element.asTable();

				// Create a detailed structure representation
				Logger.log("=== DETAILED TABLE STRUCTURE ===");
				Logger.log("Table ID: " + element.getObjectId());
				Logger.log(
					"Dimensions: " +
						table.getNumRows() +
						" rows × " +
						table.getNumColumns() +
						" columns",
				);

				// Log all cells with content and formatting
				for (let r = 0; r < table.getNumRows(); r++) {
					for (let c = 0; c < table.getNumColumns(); c++) {
						try {
							const cell = table.getCell(r, c);
							const text = cell.getText().asString();

							if (text.trim() !== "") {
								Logger.log(`Cell [${r},${c}]: "${text}"`);

								// Try to get text style
								try {
									const textStyle = cell.getText().getTextStyle();
									Logger.log(`  - Bold: ${textStyle.isBold()}`);
									Logger.log(
										`  - Font: ${textStyle.getFontFamily() || "default"}`,
									);
									Logger.log(
										`  - Size: ${textStyle.getFontSize() || "default"}`,
									);
								} catch (e) {}

								// Try to get cell fill
								try {
									const fill = cell.getFill();
									if (fill.getSolidFill()) {
										const color = fill.getSolidFill().getColor();
										if (color) {
											Logger.log(`  - Background: ${JSON.stringify(color)}`);
										}
									}
								} catch (e) {}
							}
						} catch (e) {
							Logger.log(`Error accessing cell [${r},${c}]: ${e.message}`);
						}
					}
				}

				return;
			}
		}
	}

	SlidesApp.getUi().alert("Please select a table first");
}
