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
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function connectSelectedShapesVertical(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
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
	const line = sA.getParentPage().insertLine(lineCategory, siteA, siteB);

	// Apply arrow styles
	if (startArrow && startArrow !== "NONE" && SlidesApp.ArrowStyle[startArrow]) {
		line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
	}
	if (endArrow && endArrow !== "NONE" && SlidesApp.ArrowStyle[endArrow]) {
		line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
	}
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
	const line = sA.getParentPage().insertLine(lineCategory, siteA, siteB);

	// Apply arrow styles
	if (startArrow && startArrow !== "NONE" && SlidesApp.ArrowStyle[startArrow]) {
		line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
	}
	if (endArrow && endArrow !== "NONE" && SlidesApp.ArrowStyle[endArrow]) {
		line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
	}
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

	// Handle hierarchical naming system
	const parentGraphId = getShapeGraphId(originalShape);
	let parentLevel = "";
	let nextLevel = "B";

	if (parentGraphId) {
		const parsed = parseGraphId(parentGraphId);
		if (parsed && parsed.current) {
			// Extract the level from current ID (e.g., "A1" -> "A", "B2" -> "B")
			const levelMatch = parsed.current.match(/^([A-Z]+)/);
			if (levelMatch) {
				parentLevel = levelMatch[1];
				nextLevel = getNextLevel(parentLevel);
			}
		}
	} else {
		// If parent doesn't have a graph ID, make it the root
		parentLevel = "A";
		nextLevel = "B";
	}

	// Generate sibling IDs for all children
	const childIds = generateSiblingIds(nextLevel, count);

	// Create multiple children as siblings
	const createdShapes = [];

	// Calculate spacing between siblings
	let siblingOffset = 0;
	if (count > 1) {
		// For horizontal directions (LEFT/RIGHT), space siblings vertically
		// For vertical directions (TOP/BOTTOM), space siblings horizontally
		if (direction === "LEFT" || direction === "RIGHT") {
			// Calculate total height needed for all siblings
			const totalHeight = count * originalHeight + (count - 1) * gap;
			// Start position to center the group
			siblingOffset = -(totalHeight - originalHeight) / 2;
		} else {
			// TOP or BOTTOM
			// Calculate total width needed for all siblings
			const totalWidth = count * originalWidth + (count - 1) * gap;
			// Start position to center the group
			siblingOffset = -(totalWidth - originalWidth) / 2;
		}
	}

	for (let i = 0; i < count; i++) {
		// Calculate position for each child
		let childLeft = originalLeft;
		let childTop = originalTop;

		switch (direction) {
			case "TOP":
				childTop = originalTop - originalHeight - gap;
				// Space siblings horizontally
				childLeft = originalLeft + siblingOffset + i * (originalWidth + gap);
				break;
			case "RIGHT":
				childLeft = originalLeft + originalWidth + gap;
				// Space siblings vertically
				childTop = originalTop + siblingOffset + i * (originalHeight + gap);
				break;
			case "BOTTOM":
				childTop = originalTop + originalHeight + gap;
				// Space siblings horizontally
				childLeft = originalLeft + siblingOffset + i * (originalWidth + gap);
				break;
			case "LEFT":
				childLeft = originalLeft - originalWidth - gap;
				// Space siblings vertically
				childTop = originalTop + siblingOffset + i * (originalHeight + gap);
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

		// Set the hierarchical graph ID for the child shape
		const childId = childIds[i];
		const parentCurrentId = parentGraphId
			? parseGraphId(parentGraphId)?.current || "A1"
			: "A1";
		const childGraphId = generateGraphId(parentCurrentId, childId, []);
		setShapeGraphId(childShape, childGraphId);

		// Connect to parent shape (not previous shape)
		const connectionPairs = {
			TOP: { parentSide: "TOP", childSide: "BOTTOM" },
			RIGHT: { parentSide: "RIGHT", childSide: "LEFT" },
			BOTTOM: { parentSide: "BOTTOM", childSide: "TOP" },
			LEFT: { parentSide: "LEFT", childSide: "RIGHT" },
		};

		const pair = connectionPairs[direction];
		const parentSite = pickConnectionSite(originalShape, pair.parentSide);
		const childSite = pickConnectionSite(childShape, pair.childSide);

		if (parentSite && childSite) {
			// Convert lineType string to SlidesApp.LineCategory enum
			const lineCategory =
				SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
			const line = slide.insertLine(lineCategory, parentSite, childSite);

			// Apply arrow styles
			if (
				startArrow &&
				startArrow !== "NONE" &&
				SlidesApp.ArrowStyle[startArrow]
			) {
				line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
			}
			if (endArrow && endArrow !== "NONE" && SlidesApp.ArrowStyle[endArrow]) {
				line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
			}
		}

		createdShapes.push(childShape);
	}

	// Update parent shape to include new children in its graph ID
	updateParentWithChildren(originalShape, childIds);

	return createdShapes;
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
 * Hierarchical Shape Naming System
 * Manages parent-child relationships with structured IDs
 * Format: graph[parent][current][children]
 * Example: graph[A1][B1][C1,C2,C3]
 */

/**
 * Parses a graph ID to extract its components
 * @param {string} graphId - The graph ID to parse
 * @returns {Object|null} - Object with parent, current, children arrays or null if not a graph ID
 */
function parseGraphId(graphId) {
	if (!graphId || !graphId.startsWith("graph[")) {
		return null;
	}

	// Extract content between brackets using regex
	const matches = graphId.match(/graph\[([^\]]*)\]\[([^\]]*)\]\[([^\]]*)\]/);
	if (!matches) {
		return null;
	}

	return {
		parent: matches[1] || "",
		current: matches[2] || "",
		children: matches[3] ? matches[3].split(",").filter((c) => c.trim()) : [],
	};
}

/**
 * Generates a graph ID from components
 * @param {string} parent - Parent ID
 * @param {string} current - Current ID
 * @param {Array} children - Array of child IDs
 * @returns {string} - Generated graph ID
 */
function generateGraphId(parent, current, children = []) {
	const childrenStr = children.join(",");
	return `graph[${parent}][${current}][${childrenStr}]`;
}

/**
 * Generates the next level ID for a child
 * @param {string} parentLevel - Parent's level (e.g., "A", "B", "C")
 * @returns {string} - Next level (e.g., "A" -> "B", "B" -> "C")
 */
function getNextLevel(parentLevel) {
	if (!parentLevel) return "A";
	const lastChar = parentLevel.charAt(parentLevel.length - 1);
	if (lastChar >= "A" && lastChar < "Z") {
		return (
			parentLevel.slice(0, -1) + String.fromCharCode(lastChar.charCodeAt(0) + 1)
		);
	}
	return parentLevel + "A";
}

/**
 * Generates sibling IDs for multiple children
 * @param {string} baseLevel - Base level for siblings (e.g., "C")
 * @param {number} count - Number of siblings to generate
 * @returns {Array} - Array of sibling IDs (e.g., ["C1", "C2", "C3"])
 */
function generateSiblingIds(baseLevel, count) {
	const siblings = [];
	for (let i = 1; i <= count; i++) {
		siblings.push(`${baseLevel}${i}`);
	}
	return siblings;
}

/**
 * Updates a shape's text with a new graph ID
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to update
 * @param {string} graphId - New graph ID to set
 */
function setShapeGraphId(shape, graphId) {
	try {
		// Set the shape's text content to the graph ID
		shape.getText().setText(graphId);
	} catch (e) {
		console.log(`Warning: Could not set shape text: ${e.message}`);
	}
}

/**
 * Gets the graph ID from a shape's text content
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to read from
 * @returns {string|null} - Graph ID or null if not found
 */
function getShapeGraphId(shape) {
	try {
		// Get the shape's text content
		const text = shape.getText().asString().trim();
		if (text && text.startsWith("graph[")) {
			return text;
		}
		return null;
	} catch (e) {
		console.log(`Warning: Could not get shape text: ${e.message}`);
		return null;
	}
}

/**
 * Updates parent shape to include new children in its ID
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape to update
 * @param {Array} newChildIds - Array of new child IDs to add
 */
function updateParentWithChildren(parentShape, newChildIds) {
	const currentId = getShapeGraphId(parentShape);
	if (!currentId) {
		// If parent doesn't have a graph ID, create one
		const newId = generateGraphId("", "A1", newChildIds);
		setShapeGraphId(parentShape, newId);
		return;
	}

	const parsed = parseGraphId(currentId);
	if (!parsed) return;

	// Merge existing children with new ones
	const allChildren = [...parsed.children, ...newChildIds];
	const updatedId = generateGraphId(parsed.parent, parsed.current, allChildren);
	setShapeGraphId(parentShape, updatedId);
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
	const rootId = generateGraphId("", "A1", []);
	setShapeGraphId(shape, rootId);

	SlidesApp.getUi().alert(
		"Root graph shape initialized",
		`Shape is now: ${rootId}`,
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

	let parentShape, childShape, parentId, childId;

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
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent
	const newChildId = generateGraphId(
		parentId.current,
		childId.current,
		childId.children,
	);
	setShapeGraphId(childShape, newChildId);

	// Create the visual connection using the existing connection logic
	if (parentShape === shapeA) {
		return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
	} else {
		return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
	}
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

/**
 * Debug function to show placeholder text for the currently selected shape
 * @returns {string} Debug information about the selected shape's placeholder text
 */
function debugShowTitlePlaceholders() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to debug its placeholder text.";
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
		const debugInfo = [];

		// Check the selected shape's placeholder text
		debugInfo.push("🔍 Selected Shape Debug:");

		// Shape text content
		const textContent = shape.getText().asString().trim();
		debugInfo.push(`Text Content: "${textContent || "(empty)"}"`);

		// Graph ID from our helper function
		const graphId = getShapeGraphId(shape);
		debugInfo.push(`Detected Graph ID: "${graphId || "(none)"}"`);

		const result = debugInfo.join("\n");

		// Also log to console for debugging
		console.log(`Debug Selected Shape:\n${result}`);

		return result;
	} catch (e) {
		const errorMsg = `Error in debugShowTitlePlaceholders: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

/**
 * Creates a sibling shape next to the selected shape
 * Analyzes existing sibling positions and places the new sibling appropriately
 * @param {number} horizontalGap - Horizontal gap in points
 * @param {number} verticalGap - Vertical gap in points
 * @param {string} lineType - Type of line to use
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function createSiblingShape(
	horizontalGap = 20,
	verticalGap = 20,
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert(
			"Please select a shape to create a sibling for.",
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

	const selectedShape = element.asShape();
	const selectedGraphId = getShapeGraphId(selectedShape);

	if (!selectedGraphId) {
		return SlidesApp.getUi().alert(
			"Selected shape must have a graph ID. Please create it as part of a flowchart first.",
		);
	}

	const parsed = parseGraphId(selectedGraphId);
	if (!parsed || !parsed.parent) {
		return SlidesApp.getUi().alert(
			"Selected shape must have a parent. Cannot create sibling for root shapes.",
		);
	}

	// Find the parent shape
	const slide = selectedShape.getParentPage();
	const allShapes = slide.getShapes();
	let parentShape = null;
	const siblingShapes = [];

	// Find parent and all sibling shapes
	for (const shape of allShapes) {
		const graphId = getShapeGraphId(shape);
		if (graphId) {
			const shapeData = parseGraphId(graphId);
			if (shapeData) {
				// Check if this is the parent
				if (shapeData.current === parsed.parent) {
					parentShape = shape;
				}
				// Check if this is a sibling (same parent, different current ID)
				if (
					shapeData.parent === parsed.parent &&
					shapeData.current !== parsed.current
				) {
					siblingShapes.push({
						shape: shape,
						data: shapeData,
						left: shape.getLeft(),
						top: shape.getTop(),
						width: shape.getWidth(),
						height: shape.getHeight(),
					});
				}
			}
		}
	}

	if (!parentShape) {
		return SlidesApp.getUi().alert(
			"Could not find parent shape. The hierarchy may be broken.",
		);
	}

	// Add the selected shape to the sibling list for position analysis
	siblingShapes.push({
		shape: selectedShape,
		data: parsed,
		left: selectedShape.getLeft(),
		top: selectedShape.getTop(),
		width: selectedShape.getWidth(),
		height: selectedShape.getHeight(),
	});

	// Determine the layout pattern (horizontal or vertical)
	let isHorizontalLayout = false;
	let isVerticalLayout = false;

	if (siblingShapes.length > 1) {
		// Check if siblings are arranged horizontally (similar Y positions)
		const firstSibling = siblingShapes[0];
		const tolerance = 10; // pixels

		let horizontalCount = 0;
		let verticalCount = 0;

		for (let i = 1; i < siblingShapes.length; i++) {
			const sibling = siblingShapes[i];
			const deltaY = Math.abs(sibling.top - firstSibling.top);
			const deltaX = Math.abs(sibling.left - firstSibling.left);

			if (deltaY < tolerance) horizontalCount++;
			if (deltaX < tolerance) verticalCount++;
		}

		isHorizontalLayout = horizontalCount > verticalCount;
		isVerticalLayout = verticalCount > horizontalCount;
	}

	// Generate new sibling ID
	const parentData = parseGraphId(getShapeGraphId(parentShape));
	if (!parentData) {
		return SlidesApp.getUi().alert("Parent shape has invalid graph ID format.");
	}

	// Find the highest numbered sibling to determine the next number
	const currentLevel = parsed.current.match(/^([A-Z]+)/)?.[1] || "A";
	let maxSiblingNumber = 0;

	for (const sibling of siblingShapes) {
		const siblingLevel = sibling.data.current.match(/^([A-Z]+)(\d+)$/);
		if (siblingLevel && siblingLevel[1] === currentLevel) {
			const number = Number.parseInt(siblingLevel[2]);
			if (number > maxSiblingNumber) {
				maxSiblingNumber = number;
			}
		}
	}

	const newSiblingId = `${currentLevel}${maxSiblingNumber + 1}`;

	// Calculate position for new sibling and move selected shape
	const selectedLeft = selectedShape.getLeft();
	const selectedTop = selectedShape.getTop();
	const selectedWidth = selectedShape.getWidth();
	const selectedHeight = selectedShape.getHeight();

	let newLeft;
	let newTop;
	let adjustedSelectedTop;

	if (isHorizontalLayout && siblingShapes.length > 1) {
		// Horizontal layout: center group around parent's X center
		const parentCenterX = parentShape.getLeft() + parentShape.getWidth() / 2;
		const totalGroupWidth = selectedWidth + horizontalGap + selectedWidth; // current + gap + new sibling
		const groupStartX = parentCenterX - totalGroupWidth / 2;

		selectedShape.setLeft(groupStartX);
		newLeft = groupStartX + selectedWidth + horizontalGap;
		newTop = selectedTop;
	} else {
		// Vertical layout (default): center group around parent's Y center
		const parentCenterY = parentShape.getTop() + parentShape.getHeight() / 2;
		const totalGroupHeight = selectedHeight + verticalGap + selectedHeight; // current + gap + new sibling
		const groupStartY = parentCenterY - totalGroupHeight / 2;

		selectedShape.setTop(groupStartY);
		newLeft = selectedLeft;
		newTop = groupStartY + selectedHeight + verticalGap;
	}

	// Create the new sibling shape
	const newShape = slide.insertShape(
		selectedShape.getShapeType(),
		newLeft,
		newTop,
		selectedWidth,
		selectedHeight,
	);

	// Copy styling from selected shape
	copyShapeStyle(selectedShape, newShape);

	// Set the hierarchical graph ID
	const newGraphId = generateGraphId(parsed.parent, newSiblingId, []);
	setShapeGraphId(newShape, newGraphId);

	// Update parent to include the new sibling
	const updatedChildren = [...parentData.children, newSiblingId];
	const updatedParentId = generateGraphId(
		parentData.parent,
		parentData.current,
		updatedChildren,
	);
	setShapeGraphId(parentShape, updatedParentId);

	// Connect new sibling to parent
	const connectionPairs = {
		horizontal: { parentSide: "RIGHT", childSide: "LEFT" },
		vertical: { parentSide: "BOTTOM", childSide: "TOP" },
	};

	const connectionType =
		isHorizontalLayout || (!isHorizontalLayout && !isVerticalLayout)
			? "horizontal"
			: "vertical";
	const pair = connectionPairs[connectionType];

	const parentSite = pickConnectionSite(parentShape, pair.parentSide);
	const childSite = pickConnectionSite(newShape, pair.childSide);

	if (parentSite && childSite) {
		const lineCategory =
			SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.STRAIGHT;
		const line = slide.insertLine(lineCategory, parentSite, childSite);

		// Apply arrow styles
		if (
			startArrow &&
			startArrow !== "NONE" &&
			SlidesApp.ArrowStyle[startArrow]
		) {
			line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
		}
		if (endArrow && endArrow !== "NONE" && SlidesApp.ArrowStyle[endArrow]) {
			line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
		}
	}

	return newShape;
}
