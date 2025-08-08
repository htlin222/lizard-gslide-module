/**
 * Connection utilities for connecting shapes with lines
 * Handles validation, positioning, and line creation between shapes
 */

/**
 * Validates that two elements are valid shapes for connection
 * @param {Array} elements - Array of page elements
 * @returns {Object} - Validation result with shapes or error message
 */
function validateConnectionElements(elements) {
	if (!elements || elements.length !== 2) {
		return { error: "Please select exactly TWO shapes." };
	}

	if (
		elements[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		elements[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return { error: "Both selected items must be SHAPES." };
	}

	const shapeA = elements[0].asShape();
	const shapeB = elements[1].asShape();

	// Same slide check
	if (
		String(shapeA.getParentPage().getObjectId()) !==
		String(shapeB.getParentPage().getObjectId())
	) {
		return { error: "Both shapes must be on the SAME slide." };
	}

	return { shapeA, shapeB };
}

/**
 * Determines connection sides based on relative positions
 * @param {GoogleAppsScript.Slides.Shape} shapeA - First shape
 * @param {GoogleAppsScript.Slides.Shape} shapeB - Second shape
 * @param {string} orientation - "horizontal" or "vertical"
 * @returns {Object} - Connection sides for both shapes
 */
function determineConnectionSides(shapeA, shapeB, orientation) {
	const centerA = {
		x: shapeA.getLeft() + shapeA.getWidth() / 2,
		y: shapeA.getTop() + shapeA.getHeight() / 2,
	};
	const centerB = {
		x: shapeB.getLeft() + shapeB.getWidth() / 2,
		y: shapeB.getTop() + shapeB.getHeight() / 2,
	};

	if (orientation === "horizontal") {
		const dx = centerB.x - centerA.x;
		if (dx > 0) {
			return { sideA: "RIGHT", sideB: "LEFT" };
		} else {
			return { sideA: "LEFT", sideB: "RIGHT" };
		}
	} else {
		// vertical
		const dy = centerB.y - centerA.y;
		if (dy > 0) {
			return { sideA: "BOTTOM", sideB: "TOP" };
		} else {
			return { sideA: "TOP", sideB: "BOTTOM" };
		}
	}
}

/**
 * Creates a line connection between two shapes
 * @param {GoogleAppsScript.Slides.Shape} shapeA - First shape
 * @param {GoogleAppsScript.Slides.Shape} shapeB - Second shape
 * @param {string} orientation - "horizontal" or "vertical"
 * @param {string} lineType - Type of line (STRAIGHT, BENT, CURVED)
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @returns {GoogleAppsScript.Slides.Line|null} - Created line or null if failed
 */
function createConnection(
	shapeA,
	shapeB,
	orientation,
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const sides = determineConnectionSides(shapeA, shapeB, orientation);

	const siteA = pickConnectionSite(shapeA, sides.sideA);
	const siteB = pickConnectionSite(shapeB, sides.sideB);

	if (!siteA || !siteB) {
		return null;
	}

	// Convert lineType string to SlidesApp.LineCategory enum
	const lineCategory =
		SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
	const line = shapeA.getParentPage().insertLine(lineCategory, siteA, siteB);

	// Apply arrow styles
	if (startArrow && startArrow !== "NONE" && SlidesApp.ArrowStyle[startArrow]) {
		line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
	}
	if (endArrow && endArrow !== "NONE" && SlidesApp.ArrowStyle[endArrow]) {
		line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
	}

	return line;
}

/**
 * Main function to connect two selected shapes
 * @param {string} orientation - "horizontal" or "vertical"
 * @param {string} lineType - Type of line to use
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function connectSelectedShapes(
	orientation,
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	}

	const elements = range.getPageElements();
	const validation = validateConnectionElements(elements);

	if (validation.error) {
		return SlidesApp.getUi().alert(validation.error);
	}

	const line = createConnection(
		validation.shapeA,
		validation.shapeB,
		orientation,
		lineType,
		startArrow,
		endArrow,
	);

	if (!line) {
		return SlidesApp.getUi().alert(
			"Could not resolve suitable connection sites.",
		);
	}
}
