/**
 * Flowchart utilities for connecting and creating related shapes
 * Provides functionality for linking shapes and creating child shapes in different directions
 */

/**
 * Shows the flowchart sidebar for interactive flowchart operations
 */
function showFlowchartSidebar() {
	try {
		const html = HtmlService.createHtmlOutputFromFile(
			"src/components/flowchartSidebar.html",
		)
			.setWidth(300)
			.setTitle("Flowchart Tools");

		SlidesApp.getUi().showSidebar(html);
	} catch (e) {
		console.error("Error showing flowchart sidebar: " + e.message);
		SlidesApp.getUi().alert(
			"Error",
			"Could not open the flowchart sidebar: " + e.message,
		);
	}
}

/**
 * Connects two selected shapes with a smart line
 * This is the main function called from the sidebar
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 */
function connectSelectedShapesSmart(lineType = "STRAIGHT") {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();
	if (!range)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	const els = range.getPageElements();
	if (els.length !== 2)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	if (
		els[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		els[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return SlidesApp.getUi().alert("Both selected items must be SHAPES.");
	}

	const sA = els[0].asShape();
	const sB = els[1].asShape();

	// Same slide check
	if (
		String(sA.getParentPage().getObjectId()) !==
		String(sB.getParentPage().getObjectId())
	) {
		return SlidesApp.getUi().alert("Both shapes must be on the SAME slide.");
	}

	// This function is kept for backwards compatibility
	// Default to horizontal connection
	return connectSelectedShapesHorizontal(lineType);
}

/**
 * Connects two selected shapes vertically (top/down)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 */
function connectSelectedShapesVertical(lineType = "STRAIGHT") {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();
	if (!range)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	const els = range.getPageElements();
	if (els.length !== 2)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	if (
		els[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		els[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return SlidesApp.getUi().alert("Both selected items must be SHAPES.");
	}

	const sA = els[0].asShape();
	const sB = els[1].asShape();

	// Same slide check
	if (
		String(sA.getParentPage().getObjectId()) !==
		String(sB.getParentPage().getObjectId())
	) {
		return SlidesApp.getUi().alert("Both shapes must be on the SAME slide.");
	}

	// Center points for vertical connection
	const cA = centerOf(sA);
	const cB = centerOf(sB);
	const dy = cB.y - cA.y;

	// Determine which shape is on top
	let sideA;
	let sideB;
	if (dy > 0) {
		// A on top, B on bottom
		sideA = "BOTTOM";
		sideB = "TOP";
	} else {
		// A on bottom, B on top
		sideA = "TOP";
		sideB = "BOTTOM";
	}

	const siteA = pickConnectionSite(sA, sideA);
	const siteB = pickConnectionSite(sB, sideB);
	if (!siteA || !siteB)
		return SlidesApp.getUi().alert(
			"Could not resolve suitable connection sites.",
		);

	// Convert lineType string to SlidesApp.LineCategory enum
	const lineCategory =
		SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
	sA.getParentPage().insertLine(lineCategory, siteA, siteB);
}

/**
 * Connects two selected shapes horizontally (left/right)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 */
function connectSelectedShapesHorizontal(lineType = "STRAIGHT") {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();
	if (!range)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	const els = range.getPageElements();
	if (els.length !== 2)
		return SlidesApp.getUi().alert("Please select exactly TWO shapes.");
	if (
		els[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		els[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return SlidesApp.getUi().alert("Both selected items must be SHAPES.");
	}

	const sA = els[0].asShape();
	const sB = els[1].asShape();

	// Same slide check
	if (
		String(sA.getParentPage().getObjectId()) !==
		String(sB.getParentPage().getObjectId())
	) {
		return SlidesApp.getUi().alert("Both shapes must be on the SAME slide.");
	}

	// Center points for horizontal connection
	const cA = centerOf(sA);
	const cB = centerOf(sB);
	const dx = cB.x - cA.x;

	// Determine which shape is on the left
	let sideA;
	let sideB;
	if (dx > 0) {
		// A on left, B on right
		sideA = "RIGHT";
		sideB = "LEFT";
	} else {
		// A on right, B on left
		sideA = "LEFT";
		sideB = "RIGHT";
	}

	const siteA = pickConnectionSite(sA, sideA);
	const siteB = pickConnectionSite(sB, sideB);
	if (!siteA || !siteB)
		return SlidesApp.getUi().alert(
			"Could not resolve suitable connection sites.",
		);

	// Convert lineType string to SlidesApp.LineCategory enum
	const lineCategory =
		SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
	sA.getParentPage().insertLine(lineCategory, siteA, siteB);
}

/**
 * Base function to create child shapes in any direction
 * @param {string} direction - Direction to create child ("TOP", "RIGHT", "BOTTOM", "LEFT")
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 */
function createChildInDirection(
	direction,
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert(
			`Please select a shape to create a child ${direction.toLowerCase()} it.`,
		);
	}

	const els = range.getPageElements();
	if (els.length !== 1) {
		return SlidesApp.getUi().alert("Please select exactly ONE shape.");
	}

	const element = els[0];
	if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
		return SlidesApp.getUi().alert("Selected item must be a SHAPE.");
	}

	const originalShape = element.asShape();
	const slide = originalShape.getParentPage();

	// Get original shape properties
	const originalLeft = originalShape.getLeft();
	const originalTop = originalShape.getTop();
	const originalWidth = originalShape.getWidth();
	const originalHeight = originalShape.getHeight();

	// Create multiple children
	const createdShapes = [];
	let previousShape = originalShape;

	for (let i = 0; i < count; i++) {
		// Calculate position for each child
		let childLeft = originalLeft;
		let childTop = originalTop;

		switch (direction) {
			case "TOP":
				childTop = originalTop - (originalHeight + gap) * (i + 1);
				break;
			case "RIGHT":
				childLeft = originalLeft + (originalWidth + gap) * (i + 1);
				break;
			case "BOTTOM":
				childTop = originalTop + (originalHeight + gap) * (i + 1);
				break;
			case "LEFT":
				childLeft = originalLeft - (originalWidth + gap) * (i + 1);
				break;
		}

		// Create new shape
		const childShape = slide.insertShape(
			originalShape.getShapeType(),
			childLeft,
			childTop,
			originalWidth,
			originalHeight,
		);

		// Copy styling from original shape
		copyShapeStyle(originalShape, childShape);

		// Connect to previous shape
		const connectionPairs = {
			TOP: { previousSide: "TOP", childSide: "BOTTOM" },
			RIGHT: { previousSide: "RIGHT", childSide: "LEFT" },
			BOTTOM: { previousSide: "BOTTOM", childSide: "TOP" },
			LEFT: { previousSide: "LEFT", childSide: "RIGHT" },
		};

		const pair = connectionPairs[direction];
		const previousSite = pickConnectionSite(previousShape, pair.previousSide);
		const childSite = pickConnectionSite(childShape, pair.childSide);

		if (previousSite && childSite) {
			// Convert lineType string to SlidesApp.LineCategory enum
			const lineCategory =
				SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
			slide.insertLine(lineCategory, previousSite, childSite);
		}

		createdShapes.push(childShape);
		previousShape = childShape;
	}

	return createdShapes;
}

/**
 * Creates child shapes above the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 */
function createChildTop(gap = 20, lineType = "STRAIGHT", count = 1) {
	return createChildInDirection("TOP", gap, lineType, count);
}

/**
 * Creates child shapes to the right of the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 */
function createChildRight(gap = 20, lineType = "STRAIGHT", count = 1) {
	return createChildInDirection("RIGHT", gap, lineType, count);
}

/**
 * Creates child shapes below the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 */
function createChildBottom(gap = 20, lineType = "STRAIGHT", count = 1) {
	return createChildInDirection("BOTTOM", gap, lineType, count);
}

/**
 * Creates child shapes to the left of the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 */
function createChildLeft(gap = 20, lineType = "STRAIGHT", count = 1) {
	return createChildInDirection("LEFT", gap, lineType, count);
}

/**
 * Helper function to get center coordinates of a shape
 */
function centerOf(el) {
	return {
		x: el.getLeft() + el.getWidth() / 2,
		y: el.getTop() + el.getHeight() / 2,
	};
}

/**
 * Helper function to get preferred connection site mapping
 * Left-right indices are swapped (top/bottom remain the same)
 */
function getPreferredMappingForType(shapeType, n) {
	// 8 connection points (common case): original LEFT:7, RIGHT:3 → swap to LEFT:3, RIGHT:7
	if (n >= 8) return { LEFT: 3, RIGHT: 7, TOP: 1, BOTTOM: 5 };

	// 4 connection points: assume [TOP, RIGHT, BOTTOM, LEFT]
	// Swap left-right → LEFT:1, RIGHT:3 (TOP/BOTTOM unchanged)
	if (n === 4) return { LEFT: 1, RIGHT: 3, TOP: 0, BOTTOM: 2 };

	// 2 connection points: swap left-right (TOP/BOTTOM first maintain)
	if (n === 2) return { LEFT: 1, RIGHT: 0, TOP: 0, BOTTOM: 1 };

	// 1 or other non-standard: can only use 0
	if (n === 1) return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };

	// fallback
	return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };
}

/**
 * Helper function to pick connection site for a shape
 */
function pickConnectionSite(shape, side) {
	const sites = shape.getConnectionSites();
	if (!sites || sites.length === 0) return null;
	const map = getPreferredMappingForType(shape.getShapeType(), sites.length);
	const idx = map[side];
	if (idx != null && idx < sites.length) return sites[idx];
	return sites[0];
}

/**
 * Helper function to copy style from one shape to another
 */
function copyShapeStyle(sourceShape, targetShape) {
	try {
		// Copy fill
		const sourceFill = sourceShape.getFill();
		if (sourceFill && sourceFill.getSolidFill()) {
			targetShape.getFill().setSolidFill(sourceFill.getSolidFill().getColor());
		}

		// Copy border
		const sourceBorder = sourceShape.getBorder();
		if (sourceBorder) {
			const targetBorder = targetShape.getBorder();
			if (
				sourceBorder.getLineFill() &&
				sourceBorder.getLineFill().getSolidFill()
			) {
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
