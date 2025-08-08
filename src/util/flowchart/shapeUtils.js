/**
 * Shape utility functions for common shape operations
 * Handles styling, positioning, connection sites, and validation
 */

/**
 * Gets center coordinates of a shape
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get center of
 * @returns {Object} - Center coordinates {x, y}
 */
function getCenterOf(shape) {
	return {
		x: shape.getLeft() + shape.getWidth() / 2,
		y: shape.getTop() + shape.getHeight() / 2,
	};
}

/**
 * Gets preferred connection site mapping for different shape types
 * @param {GoogleAppsScript.Slides.ShapeType} shapeType - Type of shape
 * @param {number} connectionCount - Number of connection sites available
 * @returns {Object} - Mapping of sides to connection site indices
 */
function getPreferredConnectionMapping(shapeType, connectionCount) {
	// 8 connection points (common case): original LEFT:7, RIGHT:3 → swap to LEFT:3, RIGHT:7
	if (connectionCount >= 8) {
		return { LEFT: 3, RIGHT: 7, TOP: 1, BOTTOM: 5 };
	}

	// 4 connection points: assume [TOP, RIGHT, BOTTOM, LEFT]
	// Swap left-right → LEFT:1, RIGHT:3 (TOP/BOTTOM unchanged)
	if (connectionCount === 4) {
		return { LEFT: 1, RIGHT: 3, TOP: 0, BOTTOM: 2 };
	}

	// 2 connection points: swap left-right (TOP/BOTTOM first maintain)
	if (connectionCount === 2) {
		return { LEFT: 1, RIGHT: 0, TOP: 0, BOTTOM: 1 };
	}

	// 1 or other non-standard: can only use 0
	if (connectionCount === 1) {
		return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };
	}

	// fallback
	return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };
}

/**
 * Picks the best connection site for a shape on a given side
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get connection site from
 * @param {string} side - Side to connect to (LEFT, RIGHT, TOP, BOTTOM)
 * @returns {GoogleAppsScript.Slides.ConnectionSite|null} - Connection site or null if none available
 */
function pickConnectionSite(shape, side) {
	const sites = shape.getConnectionSites();
	if (!sites || sites.length === 0) return null;

	const mapping = getPreferredConnectionMapping(
		shape.getShapeType(),
		sites.length,
	);
	const index = mapping[side];

	if (index != null && index < sites.length) {
		return sites[index];
	}
	return sites[0];
}

/**
 * Copies style properties from source shape to target shape
 * @param {GoogleAppsScript.Slides.Shape} sourceShape - Shape to copy style from
 * @param {GoogleAppsScript.Slides.Shape} targetShape - Shape to apply style to
 */
function copyShapeStyle(sourceShape, targetShape) {
	try {
		// Copy fill
		const sourceFill = sourceShape.getFill();
		if (sourceFill?.getSolidFill()) {
			targetShape.getFill().setSolidFill(sourceFill.getSolidFill().getColor());
		}

		// Copy border
		const sourceBorder = sourceShape.getBorder();
		if (sourceBorder) {
			const targetBorder = targetShape.getBorder();
			if (sourceBorder.getLineFill()?.getSolidFill()) {
				targetBorder
					.getLineFill()
					.setSolidFill(sourceBorder.getLineFill().getSolidFill().getColor());
			}
			targetBorder.setWeight(sourceBorder.getWeight());
			targetBorder.setDashStyle(sourceBorder.getDashStyle());
		}

		// Copy text style if there's text
		const sourceText = sourceShape.getText();
		const targetText = targetShape.getText();
		if (sourceText && targetText) {
			// Copy text content
			targetText.setText(sourceText.asString());

			// Copy text style
			const sourceStyle = sourceText.getTextStyle();
			const targetStyle = targetText.getTextStyle();

			if (sourceStyle.getFontFamily()) {
				targetStyle.setFontFamily(sourceStyle.getFontFamily());
			}
			if (sourceStyle.getFontSize()) {
				targetStyle.setFontSize(sourceStyle.getFontSize());
			}
			if (sourceStyle.getForegroundColor()) {
				targetStyle.setForegroundColor(sourceStyle.getForegroundColor());
			}
			if (sourceStyle.isBold()) {
				targetStyle.setBold(sourceStyle.isBold());
			}
			if (sourceStyle.isItalic()) {
				targetStyle.setItalic(sourceStyle.isItalic());
			}
		}
	} catch (e) {
		console.log("Warning: Could not copy all style properties: " + e.message);
	}
}

/**
 * Validates that a selection contains valid shapes
 * @param {GoogleAppsScript.Slides.PageElementRange} range - Selection range
 * @param {number} expectedCount - Expected number of shapes
 * @returns {Object} - Validation result with shapes array or error message
 */
function validateShapeSelection(range, expectedCount) {
	if (!range) {
		const message =
			expectedCount === 1
				? "Please select a shape."
				: `Please select exactly ${expectedCount === 2 ? "TWO" : expectedCount} shapes.`;
		return { error: message };
	}

	const elements = range.getPageElements();
	if (elements.length !== expectedCount) {
		const message =
			expectedCount === 1
				? "Please select exactly ONE shape."
				: `Please select exactly ${expectedCount === 2 ? "TWO" : expectedCount} shapes.`;
		return { error: message };
	}

	// Validate all elements are shapes
	for (const element of elements) {
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			const message =
				expectedCount === 1
					? "Selected item must be a SHAPE."
					: "All selected items must be SHAPES.";
			return { error: message };
		}
	}

	// Convert to shapes
	const shapes = elements.map((element) => element.asShape());

	// For multiple shapes, verify they're on the same slide
	if (expectedCount > 1) {
		const firstSlideId = String(shapes[0].getParentPage().getObjectId());
		for (let i = 1; i < shapes.length; i++) {
			if (String(shapes[i].getParentPage().getObjectId()) !== firstSlideId) {
				return { error: "All shapes must be on the SAME slide." };
			}
		}
	}

	return expectedCount === 1 ? { shape: shapes[0] } : { shapes };
}

/**
 * Gets the current selection from the active presentation
 * @returns {GoogleAppsScript.Slides.PageElementRange|null} - Selection range or null
 */
function getCurrentSelection() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		return selection.getPageElementRange();
	} catch (e) {
		console.log(`Warning: Could not get selection: ${e.message}`);
		return null;
	}
}

/**
 * Displays an alert to the user
 * @param {string} message - Message to display
 * @param {string} title - Optional title for the alert
 */
function showAlert(message, title = "Alert") {
	try {
		SlidesApp.getUi().alert(title, message);
	} catch (e) {
		console.error(`Failed to show alert: ${e.message}`);
	}
}

/**
 * Gets shape properties in a standardized format
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get properties from
 * @returns {Object} - Shape properties {left, top, width, height, center}
 */
function getShapeProperties(shape) {
	const left = shape.getLeft();
	const top = shape.getTop();
	const width = shape.getWidth();
	const height = shape.getHeight();

	return {
		left,
		top,
		width,
		height,
		center: {
			x: left + width / 2,
			y: top + height / 2,
		},
	};
}
