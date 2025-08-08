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
 * Updates a shape's title with a new graph ID
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to update
 * @param {string} graphId - New graph ID to set
 */
function setShapeGraphId(shape, graphId) {
	try {
		// Set the shape's alt title for identification
		shape.setTitle(graphId);

		// Try to use title placeholder if it exists
		const slide = shape.getParentPage();
		try {
			const titlePlaceholder = slide.getPlaceholder(
				SlidesApp.PlaceholderType.TITLE,
			);
			if (titlePlaceholder) {
				titlePlaceholder.asShape().getText().setText(graphId);
			}
		} catch (placeholderError) {
			// If no title placeholder, set the shape's own text as fallback
			try {
				shape.getText().setText(graphId);
			} catch (textError) {
				console.log(`Warning: Could not set shape text: ${textError.message}`);
			}
		}
	} catch (e) {
		console.log(`Warning: Could not set shape title: ${e.message}`);
	}
}

/**
 * Gets the graph ID from a shape's title or text
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to read from
 * @returns {string|null} - Graph ID or null if not found
 */
function getShapeGraphId(shape) {
	try {
		// First try to get from alt title
		const title = shape.getTitle();
		if (title && title.startsWith("graph[")) {
			return title;
		}

		// Try to get from title placeholder
		const slide = shape.getParentPage();
		try {
			const titlePlaceholder = slide.getPlaceholder(
				SlidesApp.PlaceholderType.TITLE,
			);
			if (titlePlaceholder) {
				const placeholderText = titlePlaceholder
					.asShape()
					.getText()
					.asString()
					.trim();
				if (placeholderText && placeholderText.startsWith("graph[")) {
					return placeholderText;
				}
			}
		} catch (placeholderError) {
			// Title placeholder not found, continue to next fallback
		}

		// Fallback to shape's own text content
		const text = shape.getText().asString().trim();
		if (text && text.startsWith("graph[")) {
			return text;
		}

		return null;
	} catch (e) {
		console.log(`Warning: Could not get shape title: ${e.message}`);
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
