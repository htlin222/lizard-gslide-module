/**
 * Shape Utilities for Child Shapes
 * Handles utility functions and helper methods
 */

/**
 * Gets the selected shape from the current presentation.
 * @return {Object} Object containing the selected shape and any validation errors
 */
function getSelectedShape() {
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
		return {
			shape: null,
			error: "Please select exactly one shape.",
			hasSelection: selectedShapes.length > 0,
			selectionCount: selectedShapes.length,
		};
	}

	return {
		shape: selectedShapes[0].asShape(),
		error: null,
		hasSelection: true,
		selectionCount: 1,
	};
}

/**
 * Validates shape dimensions and parameters for child shape creation.
 * @param {number} parentWidth - Width of parent shape
 * @param {number} parentHeight - Height of parent shape
 * @param {number} rows - Number of rows
 * @param {number} columns - Number of columns
 * @param {number} padding - Padding value
 * @param {number} paddingTop - Top padding value
 * @param {number} gap - Gap between shapes
 * @return {Object} Validation result with isValid flag and error message
 */
function validateShapeDimensions(
	parentWidth,
	parentHeight,
	rows,
	columns,
	padding,
	paddingTop,
	gap,
) {
	// Calculate available space
	const availableWidth = parentWidth - padding * 2;
	const availableHeight = parentHeight - paddingTop - padding;

	// Calculate child dimensions
	const childWidth = (availableWidth - gap * (columns - 1)) / columns;
	const childHeight = (availableHeight - gap * (rows - 1)) / rows;

	if (childWidth <= 0) {
		return {
			isValid: false,
			error: `Child width would be ${childWidth.toFixed(1)}pt. Reduce padding or gap, or decrease columns.`,
		};
	}

	if (childHeight <= 0) {
		return {
			isValid: false,
			error: `Child height would be ${childHeight.toFixed(1)}pt. Reduce padding or gap, or decrease rows.`,
		};
	}

	// Warn if dimensions are very small
	if (childWidth < 20 || childHeight < 20) {
		return {
			isValid: true,
			warning: `Child shapes will be small (${childWidth.toFixed(1)} x ${childHeight.toFixed(1)}pt). Consider adjusting parameters.`,
		};
	}

	return {
		isValid: true,
		childWidth: childWidth,
		childHeight: childHeight,
	};
}

/**
 * Logs shape information for debugging purposes.
 * @param {Shape} shape - The shape to log information about
 * @param {string} label - Label for the log entry
 */
function logShapeInfo(shape, label = "Shape") {
	console.log(`${label} information:`);
	console.log(`  Type: ${shape.getShapeType()}`);
	console.log(`  ID: ${shape.getObjectId()}`);
	console.log(`  Position: Left ${shape.getLeft()}, Top ${shape.getTop()}`);
	console.log(`  Size: Width ${shape.getWidth()}, Height ${shape.getHeight()}`);
	console.log(`  Rotation: ${shape.getRotation()} degrees`);

	// Log text content if available
	const textRange = shape.getText();
	if (textRange) {
		const textContent = textRange.asString().trim();
		if (textContent) {
			console.log(`  Text: "${textContent}"`);
		}
	}
}

/**
 * Calculates optimal grid dimensions based on number of items and shape aspect ratio.
 * @param {number} itemCount - Number of items to arrange
 * @param {number} shapeWidth - Width of the container shape
 * @param {number} shapeHeight - Height of the container shape
 * @return {Object} Object with recommended rows and columns
 */
function calculateOptimalGrid(itemCount, shapeWidth, shapeHeight) {
	const aspectRatio = shapeWidth / shapeHeight;

	// Start with square root as base
	const sqrt = Math.sqrt(itemCount);

	// Adjust based on aspect ratio
	let columns = Math.ceil(sqrt * Math.sqrt(aspectRatio));
	let rows = Math.ceil(itemCount / columns);

	// Ensure we don't have too many empty cells
	while (rows * columns - itemCount > columns && rows > 1) {
		rows--;
		columns = Math.ceil(itemCount / rows);
	}

	return {
		rows: rows,
		columns: columns,
		efficiency: ((itemCount / (rows * columns)) * 100).toFixed(1),
	};
}

/**
 * Formats error messages consistently.
 * @param {string} operation - The operation that failed
 * @param {string} message - The error message
 * @return {string} Formatted error message
 */
function formatErrorMessage(operation, message) {
	return `${operation}: ${message}`;
}

/**
 * Shows a consistent alert dialog.
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 * @param {string} type - Type of alert ('info', 'warning', 'error')
 */
function showAlert(title, message, type = "info") {
	const ui = SlidesApp.getUi();

	// Add emoji based on type
	const icons = {
		info: "ℹ️",
		warning: "⚠️",
		error: "❌",
		success: "✅",
	};

	const iconTitle = `${icons[type] || ""} ${title}`.trim();

	ui.alert(iconTitle, message, ui.ButtonSet.OK);
}
