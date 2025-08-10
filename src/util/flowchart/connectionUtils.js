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
 * Finds the next available A-level ID on the current slide
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to check
 * @returns {string} - The next available A-level ID (A1, A2, A3, etc.)
 */
function findNextAvailableRootId(slide) {
	const allShapes = slide.getShapes();
	const usedRootIds = new Set();

	// Collect all A-level IDs already in use
	for (const shape of allShapes) {
		const graphId = getShapeGraphId(shape);
		if (graphId) {
			const parsed = parseGraphId(graphId);
			if (parsed && parsed.current.startsWith("A")) {
				// Extract the number from IDs like A1, A2, A3
				const match = parsed.current.match(/^A(\d+)$/);
				if (match) {
					usedRootIds.add(parseInt(match[1]));
				}
			}
		}
	}

	// Find the smallest available number
	let nextNumber = 1;
	while (usedRootIds.has(nextNumber)) {
		nextNumber++;
	}

	return `A${nextNumber}`;
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

	const { shapeA, shapeB } = validation;

	// Determine parent-child relationship based on Graph IDs and position
	// For horizontal: left shape is parent, right shape is child
	// For vertical: top shape is parent, bottom shape is child
	let parentShape, childShape;

	if (orientation === "horizontal") {
		// Determine based on horizontal position
		if (shapeA.getLeft() < shapeB.getLeft()) {
			parentShape = shapeA;
			childShape = shapeB;
		} else {
			parentShape = shapeB;
			childShape = shapeA;
		}
	} else {
		// vertical orientation - top shape is parent
		if (shapeA.getTop() < shapeB.getTop()) {
			parentShape = shapeA;
			childShape = shapeB;
		} else {
			parentShape = shapeB;
			childShape = shapeA;
		}
	}

	// Get or initialize parent's Graph ID
	let parentGraphId = getShapeGraphId(parentShape);

	if (!parentGraphId) {
		// Initialize parent as root with next available A-level ID
		const slide = parentShape.getParentPage();
		const nextRootId = findNextAvailableRootId(slide);
		parentGraphId = generateGraphId("", "", nextRootId, []);
		setShapeGraphId(parentShape, parentGraphId);
	}

	const parentData = parseGraphId(parentGraphId);
	if (!parentData) {
		return SlidesApp.getUi().alert("Failed to parse parent Graph ID.");
	}

	// Handle child's Graph ID
	const childGraphId = getShapeGraphId(childShape);
	let childId;

	if (childGraphId) {
		// Child already has a Graph ID - update its parent reference
		const childData = parseGraphId(childGraphId);
		if (childData) {
			childId = childData.current;
			// Update child's parent reference and layout
			const layout = orientation === "horizontal" ? "LR" : "TD";
			const updatedChildGraphId = generateGraphId(
				parentData.current,
				layout,
				childData.current,
				childData.children,
			);
			setShapeGraphId(childShape, updatedChildGraphId);
		}
	} else {
		// Child doesn't have a Graph ID - create one
		const nextLevel = getNextLevel(
			parentData.current.match(/^([A-Z]+)/)?.[1] || "A",
		);

		// Find the next available number for this level
		const existingChildren = parentData.children.filter((id) =>
			id.startsWith(nextLevel),
		);
		const nextNumber = existingChildren.length + 1;
		childId = `${nextLevel}${nextNumber}`;

		// Set child's Graph ID with proper layout
		const layout = orientation === "horizontal" ? "LR" : "TD";
		const newChildGraphId = generateGraphId(
			parentData.current,
			layout,
			childId,
			[],
		);
		setShapeGraphId(childShape, newChildGraphId);
	}

	// Update parent to include this child if not already present
	if (!parentData.children.includes(childId)) {
		const updatedChildren = [...parentData.children, childId];
		const updatedParentId = generateGraphId(
			parentData.parent,
			parentData.layout,
			parentData.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, updatedParentId);
	}

	// Create the visual connection
	const line = createConnection(
		parentShape,
		childShape,
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
