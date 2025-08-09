/**
 * Main flowchart interface functions
 * Provides the main entry points for flowchart operations
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
		console.error(`Error showing flowchart sidebar: ${e.message}`);
		SlidesApp.getUi().alert(
			"Error",
			`Could not open the flowchart sidebar: ${e.message}`,
		);
	}
}

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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
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
	if (!parentId.children.includes(childId.current)) {
		const updatedChildren = [...parentId.children, childId.current];
		const newParentId = generateGraphId(
			parentId.parent,
			parentId.layout,
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent using parent's layout
	const newChildId = generateGraphId(
		parentId.current,
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

/**
 * Initializes a selected shape as a root graph node
 * Useful for starting a new flowchart hierarchy
 */
function initializeRootGraphShape() {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert(
			"Please select a shape to initialize as root graph node.",
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

	const shape = element.asShape();
	// Initialize as root without layout annotation
	const rootId = generateGraphId("", "", "A1", []);
	setShapeGraphId(shape, rootId);

	SlidesApp.getUi().alert(
		"Root graph shape initialized",
		`Shape is now: ${rootId}`,
	);
}

/**
 * Shows the Graph ID of the currently selected shape
 * @returns {string} Graph ID information about the selected shape
 */
function showSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to show its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get Graph ID from title (alt text)
		const graphId = getShapeGraphId(shape);

		if (graphId) {
			// Parse the graph ID to show more details
			const parsed = parseGraphId(graphId);
			if (parsed) {
				const details = [];
				details.push(`📊 Graph ID: ${graphId}`);
				details.push(`├─ Parent: ${parsed.parent || "(root)"}`);
				details.push(`├─ Layout: ${parsed.layout || "(none)"}`);
				details.push(`├─ Current: ${parsed.current}`);
				details.push(
					`└─ Children: ${parsed.children.length > 0 ? parsed.children.join(", ") : "(none)"}`,
				);
				return details.join("\n");
			}
			return `📊 Graph ID: ${graphId}`;
		} else {
			return "No Graph ID found. This shape may not be part of a flowchart.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

/**
 * Clears the Graph ID from the currently selected shape
 * @returns {string} Confirmation message
 */
function clearSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to clear its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get current Graph ID for confirmation
		const currentGraphId = getShapeGraphId(shape);

		if (currentGraphId) {
			// Clear the title (alt text)
			shape.setTitle("");
			return `✅ Graph ID cleared successfully!\nPrevious ID was: ${currentGraphId}`;
		} else {
			return "No Graph ID to clear. This shape doesn't have a Graph ID.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

// Keep the old function name for backward compatibility but redirect to new function
function debugShowTitlePlaceholders() {
	return showSelectedShapeGraphId();
}

/**
 * Creates child shapes above the selected shape with custom text
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
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
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
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
