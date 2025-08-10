/**
 * Line update utilities for flowchart elements
 * Provides functions for updating existing lines with new styles and properties
 */

/**
 * Updates the selected lines with new line type and arrow styles
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @returns {string} - Result message
 */
function updateSelectedLines(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	try {
		console.log(
			`Starting line update: type=${lineType}, startArrow=${startArrow}, endArrow=${endArrow}`,
		);

		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();

		if (!selection) {
			console.log("No selection found");
			return "Please select one or more lines to update.";
		}

		const elements = getSelectedElements(selection);
		if (!elements) {
			return "Please select one or more line elements to update.";
		}

		console.log(`Found ${elements.length} selected elements`);

		const updateResults = processLineUpdates(
			elements,
			lineType,
			startArrow,
			endArrow,
		);

		return formatUpdateResults(updateResults);
	} catch (e) {
		console.error(`Error updating lines: ${e.message}`);
		console.error(`Stack trace: ${e.stack}`);
		return `Error: ${e.message}`;
	}
}

/**
 * Gets selected elements from the current selection
 * @param {Selection} selection - The current selection
 * @returns {Array|null} Array of selected elements or null if invalid
 */
function getSelectedElements(selection) {
	const selectionType = selection.getSelectionType();
	console.log(`Selection type: ${selectionType}`);

	let elements = [];

	if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT_RANGE) {
		// Multiple elements selected
		const range = selection.getPageElementRange();
		if (!range) {
			console.log("No page element range found");
			return null;
		}
		elements = range.getPageElements();
	} else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
		// Single element selected
		const range = selection.getPageElementRange();
		if (range) {
			elements = range.getPageElements();
		} else {
			console.log("Could not get page element range for single selection");
			return null;
		}
	} else {
		console.log(`Unsupported selection type: ${selectionType}`);
		return null;
	}

	return elements;
}

/**
 * Processes line updates for all selected elements
 * @param {Array} elements - Array of selected elements
 * @param {string} lineType - Line type to apply
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @returns {Object} Update results with counts
 */
function processLineUpdates(elements, lineType, startArrow, endArrow) {
	let updatedCount = 0;
	let skippedCount = 0;
	let errorCount = 0;

	for (let i = 0; i < elements.length; i++) {
		const element = elements[i];
		const elementType = element.getPageElementType();
		console.log(`Element ${i}: type = ${elementType}`);

		if (elementType === SlidesApp.PageElementType.LINE) {
			try {
				const success = updateSingleLine(
					element.asLine(),
					lineType,
					startArrow,
					endArrow,
					i,
				);
				if (success) {
					updatedCount++;
				} else {
					skippedCount++;
				}
			} catch (lineError) {
				console.error(`Error updating line ${i}: ${lineError.message}`);
				console.error(`Stack trace: ${lineError.stack}`);
				errorCount++;
			}
		} else {
			console.log(`Element ${i} is not a line, skipping`);
			skippedCount++;
		}
	}

	return { updatedCount, skippedCount, errorCount };
}

/**
 * Updates a single line element
 * @param {Line} line - The line to update
 * @param {string} lineType - Line type to apply
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @param {number} index - Line index for logging
 * @returns {boolean} True if update was successful
 */
function updateSingleLine(line, lineType, startArrow, endArrow, index) {
	console.log(`Processing line ${index}`);

	// Get connection info before deletion
	const slide = line.getParentPage();
	const startConnection = line.getStart();
	const endConnection = line.getEnd();

	console.log(
		`Line connections - Start: ${startConnection ? "exists" : "null"}, End: ${endConnection ? "exists" : "null"}`,
	);

	if (!startConnection || !endConnection) {
		console.log(`Line ${index} missing connection sites`);
		return false;
	}

	// Get shapes on the slide for reconnection
	const allShapes = slide.getShapes();
	console.log(`Found ${allShapes.length} shapes on slide`);

	if (allShapes.length < 2) {
		console.log(`Line ${index} - not enough shapes on slide to connect`);
		return false;
	}

	// Store line style properties before deletion
	const lineStyle = extractLineStyle(line, index);

	// Get shapes to reconnect (simple heuristic using first two shapes)
	const startShape = allShapes[0];
	const endShape = allShapes[1];

	console.log(`Using shapes: shape1 and shape2 as connection candidates`);
	console.log(
		`Shape1 at (${startShape.getLeft()}, ${startShape.getTop()}), Shape2 at (${endShape.getLeft()}, ${endShape.getTop()})`,
	);

	// Remove the old line first
	line.remove();
	console.log(`Line ${index} removed`);

	// Create new line with updated properties
	return recreateLine(
		startShape,
		endShape,
		lineType,
		startArrow,
		endArrow,
		lineStyle,
		index,
	);
}

/**
 * Extracts style properties from a line before deletion
 * @param {Line} line - The line to extract style from
 * @param {number} index - Line index for logging
 * @returns {Object} Line style properties
 */
function extractLineStyle(line, index) {
	let lineWeight = 2; // default weight
	let lineColor = null;

	try {
		const lineStyle = line.getLineStyle();
		if (lineStyle && lineStyle.getWeight) {
			lineWeight = lineStyle.getWeight();
		}

		if (lineStyle && lineStyle.getSolidFill) {
			const solidFill = lineStyle.getSolidFill();
			if (solidFill) {
				const color = solidFill.getColor();
				if (color && color.getColorType() === SlidesApp.ColorType.RGB) {
					const rgb = color.asRgbColor();
					lineColor = {
						red: rgb.getRed(),
						green: rgb.getGreen(),
						blue: rgb.getBlue(),
					};
				}
			}
		}
	} catch (styleError) {
		console.log(
			`Could not read line style: ${styleError.message}, using defaults`,
		);
	}

	console.log(
		`Line ${index} style - Weight: ${lineWeight}, Color: ${lineColor ? "exists" : "default"}`,
	);

	return { lineWeight, lineColor };
}

/**
 * Recreates a line with new properties
 * @param {Shape} startShape - Starting shape for connection
 * @param {Shape} endShape - Ending shape for connection
 * @param {string} lineType - Line type to apply
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @param {Object} lineStyle - Original line style properties
 * @param {number} index - Line index for logging
 * @returns {boolean} True if recreation was successful
 */
function recreateLine(
	startShape,
	endShape,
	lineType,
	startArrow,
	endArrow,
	lineStyle,
	index,
) {
	// Create new line using the createConnection function which handles this properly
	console.log(`Creating new line with type: ${lineType}`);

	// Determine orientation based on shape positions
	const orientation = calculateLineOrientation(startShape, endShape);
	console.log(`Using orientation: ${orientation}`);

	// Use the createConnection function from connectionUtils
	console.log(
		`Calling createConnection with startShape, endShape, orientation=${orientation}, lineType=${lineType}`,
	);

	const newLine = createConnection(
		startShape,
		endShape,
		orientation,
		lineType,
		startArrow,
		endArrow,
	);

	console.log(`createConnection returned: ${newLine ? "LINE OBJECT" : "NULL"}`);

	if (!newLine) {
		console.log(
			`❌ Failed to create connection for line ${index} - createConnection returned null`,
		);
		return false;
	}

	console.log(
		`✅ New line created for element ${index} with arrows already applied`,
	);

	// Apply original line style properties
	applyLineStyle(newLine, lineStyle);

	console.log(`Successfully updated line ${index}`);
	return true;
}

/**
 * Calculates the orientation for a line between two shapes
 * @param {Shape} startShape - Starting shape
 * @param {Shape} endShape - Ending shape
 * @returns {string} "horizontal" or "vertical"
 */
function calculateLineOrientation(startShape, endShape) {
	const startCenter = {
		x: startShape.getLeft() + startShape.getWidth() / 2,
		y: startShape.getTop() + startShape.getHeight() / 2,
	};
	const endCenter = {
		x: endShape.getLeft() + endShape.getWidth() / 2,
		y: endShape.getTop() + endShape.getHeight() / 2,
	};

	// Determine if it's more horizontal or vertical
	const dx = Math.abs(endCenter.x - startCenter.x);
	const dy = Math.abs(endCenter.y - startCenter.y);

	return dx > dy ? "horizontal" : "vertical";
}

/**
 * Applies style properties to a line
 * @param {Line} line - The line to style
 * @param {Object} lineStyle - Style properties to apply
 */
function applyLineStyle(line, lineStyle) {
	try {
		const newLineStyle = line.getLineStyle();
		if (newLineStyle && newLineStyle.setWeight) {
			newLineStyle.setWeight(lineStyle.lineWeight);
			console.log(`Set line weight: ${lineStyle.lineWeight}`);
		}

		if (lineStyle.lineColor && newLineStyle && newLineStyle.setSolidFill) {
			newLineStyle.setSolidFill(
				lineStyle.lineColor.red,
				lineStyle.lineColor.green,
				lineStyle.lineColor.blue,
			);
			console.log(
				`Set line color: RGB(${lineStyle.lineColor.red}, ${lineStyle.lineColor.green}, ${lineStyle.lineColor.blue})`,
			);
		}
	} catch (styleError) {
		console.log(
			`Could not apply line style: ${styleError.message}, but line was created`,
		);
	}
}

/**
 * Formats the update results into a user-friendly message
 * @param {Object} results - Update results with counts
 * @returns {string} Formatted message
 */
function formatUpdateResults(results) {
	const { updatedCount, skippedCount, errorCount } = results;

	console.log(
		`Update complete - Updated: ${updatedCount}, Skipped: ${skippedCount}, Errors: ${errorCount}`,
	);

	if (updatedCount === 0) {
		return errorCount > 0
			? `❌ Failed to update lines (${errorCount} errors). Check console for details.`
			: "No lines were selected. Please select one or more lines.";
	}

	let message = `✅ Updated ${updatedCount} line${updatedCount > 1 ? "s" : ""}`;
	if (skippedCount > 0) {
		message += ` (skipped ${skippedCount} non-line element${skippedCount > 1 ? "s" : ""})`;
	}
	if (errorCount > 0) {
		message += ` (${errorCount} error${errorCount > 1 ? "s" : ""})`;
	}
	return message;
}
