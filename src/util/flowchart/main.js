/**
 * Main flowchart interface functions
 * Provides the main entry points for flowchart operations
 */

/**
 * Shows the flowchart sidebar for interactive flowchart operations
 */
function showFlowchartSidebar() {
	try {
		// Use the modular sidebar approach
		const sidebar = createFlowchartSidebar();
		SlidesApp.getUi().showSidebar(sidebar);
	} catch (e) {
		console.error(`Error showing flowchart sidebar: ${e.message}`);
		SlidesApp.getUi().alert(
			"Error",
			`Could not open the flowchart sidebar: ${e.message}`,
		);
	}
}

// ====================
// CONNECTION FUNCTIONS
// ====================

/**
 * Connects two selected shapes with a smart line
 * This is the main function called from the sidebar
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 */
function connectSelectedShapesSmart(lineType = "STRAIGHT") {
	// Default to horizontal connection for backwards compatibility
	return connectSelectedShapesHorizontal(lineType);
}

/**
 * Connects two selected shapes vertically (top/down)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function connectSelectedShapesVertical(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return connectSelectedShapes("vertical", lineType, startArrow, endArrow);
}

/**
 * Connects two selected shapes horizontally (left/right)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function connectSelectedShapesHorizontal(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return connectSelectedShapes("horizontal", lineType, startArrow, endArrow);
}

/**
 * Updates connection between two existing graph shapes
 * Also handles establishing the parent-child relationship in their IDs
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function connectExistingGraphShapes(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert("Please select exactly TWO graph shapes.");
	}

	const els = range.getPageElements();
	if (els.length !== 2) {
		return SlidesApp.getUi().alert("Please select exactly TWO graph shapes.");
	}

	if (
		els[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		els[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return SlidesApp.getUi().alert("Both selected items must be SHAPES.");
	}

	const shapeA = els[0].asShape();
	const shapeB = els[1].asShape();

	// Get their graph IDs
	const idA = getShapeGraphId(shapeA);
	const idB = getShapeGraphId(shapeB);

	if (!idA || !idB) {
		return SlidesApp.getUi().alert(
			"Both shapes must have graph IDs. Use 'Initialize Root' first.",
		);
	}

	// Parse the IDs to understand the hierarchy
	const parsedA = parseGraphId(idA);
	const parsedB = parseGraphId(idB);

	if (!parsedA || !parsedB) {
		return SlidesApp.getUi().alert(
			"Invalid graph ID format on one or both shapes.",
		);
	}

	// Determine parent-child relationship based on hierarchy level
	// Lower levels (A < B < C) are parents of higher levels
	const levelA = parsedA.current.match(/^([A-Z]+)/)?.[1] || "A";
	const levelB = parsedB.current.match(/^([A-Z]+)/)?.[1] || "A";

	let parentShape;
	let childShape;
	let parentId;
	let childId;

	if (levelA <= levelB) {
		// A is parent, B is child
		parentShape = shapeA;
		childShape = shapeB;
		parentId = parsedA;
		childId = parsedB;
	} else {
		// B is parent, A is child
		parentShape = shapeB;
		childShape = shapeA;
		parentId = parsedB;
		childId = parsedA;
	}

	// Update parent to include this child
	if (!parentId.childrenIds.includes(childId.current)) {
		const updatedChildren = [...parentId.childrenIds, childId.current];
		const newParentId = generateGraphId(
			parentId.parent,
			parentId.layout,
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent with full hierarchy
	// Build parent hierarchy chain
	let parentHierarchy = "";
	if (parentId.parent) {
		// Parent already has hierarchy, append current parent
		parentHierarchy = `${parentId.parent}|${parentId.current}`;
	} else {
		// Parent is root, just use its ID
		parentHierarchy = parentId.current;
	}

	const newChildId = generateGraphId(
		parentHierarchy,
		parentId.layout,
		childId.current,
		childId.children,
	);
	setShapeGraphId(childShape, newChildId);

	// Create the visual connection using the existing connection logic
	if (parentShape === shapeA) {
		return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
	}
	return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
}

// ===================
// CHILD SHAPE CREATION
// ===================

/**
 * Base function to create child shapes in any direction
 * @param {string} direction - Direction to create child ("TOP", "RIGHT", "BOTTOM", "LEFT")
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildInDirection(
	direction,
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildrenInDirection(
		direction,
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes above the selected shape
 */
function createChildTop(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"TOP",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes to the right of the selected shape
 */
function createChildRight(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"RIGHT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes below the selected shape
 */
function createChildBottom(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"BOTTOM",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes to the left of the selected shape
 */
function createChildLeft(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"LEFT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

// =========================
// CHILD CREATION WITH TEXT
// =========================

/**
 * Creates child shapes above the selected shape with custom text
 */
function createChildTopWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"TOP",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes to the right of the selected shape with custom text
 */
function createChildRightWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"RIGHT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes below the selected shape with custom text
 */
function createChildBottomWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"BOTTOM",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes to the left of the selected shape with custom text
 */
function createChildLeftWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"LEFT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

// =================
// MODULAR FUNCTIONS AVAILABLE GLOBALLY
// =================
//
// The following functions are now available from specialized modules:
//
// From debugUtils.js:
// - showSelectedShapeGraphId()
// - clearSelectedShapeGraphId()
// - identifyConnectedShapes()
// - initializeRootGraphShape()
// - debugShowTitlePlaceholders() [alias for showSelectedShapeGraphId]
// - analyzeCurrentSlide()
//
// From lineUpdateUtils.js:
// - updateSelectedLines(lineType, startArrow, endArrow)
//
// From backgroundUtils.js:
// - addBackgroundToSelectedElements(padding, bgColor, opacity)
// - calculateShapesBoundingBox(shapes)
// - createBackgroundRectangle(slide, left, top, width, height, bgColor, opacity)
// - createCustomBackground(shapes, style)
//
// All functions are automatically available globally in Google Apps Script
