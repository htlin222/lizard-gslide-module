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
 * Creates a specific connection between two shapes with explicit sides
 * @param {GoogleAppsScript.Slides.Shape} shapeA - First shape
 * @param {GoogleAppsScript.Slides.Shape} shapeB - Second shape
 * @param {string} sideA - Connection side for shape A (TOP, RIGHT, BOTTOM, LEFT)
 * @param {string} sideB - Connection side for shape B (TOP, RIGHT, BOTTOM, LEFT)
 * @param {string} lineType - Type of line (STRAIGHT, BENT, CURVED)
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @returns {GoogleAppsScript.Slides.Line|null} - Created line or null if failed
 */
function createSpecificConnection(
	shapeA,
	shapeB,
	sideA,
	sideB,
	lineType = "BENT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const siteA = pickConnectionSite(shapeA, sideA);
	const siteB = pickConnectionSite(shapeB, sideB);

	if (!siteA || !siteB) {
		return null;
	}

	// Convert lineType string to SlidesApp.LineCategory enum
	const lineCategory =
		SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.BENT;
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
 * LLQ+RUQ Connection 1: Top of left shape → Left of right shape
 */
function connectLLQRUQ_TopLeft(
	lineType = "BENT",
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

	// Determine which is left and which is right
	let leftShape, rightShape;
	if (shapeA.getLeft() < shapeB.getLeft()) {
		leftShape = shapeA;
		rightShape = shapeB;
	} else {
		leftShape = shapeB;
		rightShape = shapeA;
	}

	// Update Graph IDs to establish relationship
	updateGraphShapeRelationship(leftShape, rightShape);

	const line = createSpecificConnection(
		leftShape,
		rightShape,
		"TOP",
		"LEFT",
		lineType,
		startArrow,
		endArrow,
	);

	if (!line) {
		return SlidesApp.getUi().alert("Could not create connection.");
	}
}

/**
 * LLQ+RUQ Connection 2: Right of left shape → Bottom of right shape
 */
function connectLLQRUQ_RightBottom(
	lineType = "BENT",
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

	// Determine which is left and which is right
	let leftShape, rightShape;
	if (shapeA.getLeft() < shapeB.getLeft()) {
		leftShape = shapeA;
		rightShape = shapeB;
	} else {
		leftShape = shapeB;
		rightShape = shapeA;
	}

	// Update Graph IDs to establish relationship
	updateGraphShapeRelationship(leftShape, rightShape);

	const line = createSpecificConnection(
		leftShape,
		rightShape,
		"RIGHT",
		"BOTTOM",
		lineType,
		startArrow,
		endArrow,
	);

	if (!line) {
		return SlidesApp.getUi().alert("Could not create connection.");
	}
}

/**
 * LUQ+RLQ Connection 1: Right of left shape → Top of right shape
 */
function connectLUQRLQ_RightTop(
	lineType = "BENT",
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

	// Determine which is left and which is right
	let leftShape, rightShape;
	if (shapeA.getLeft() < shapeB.getLeft()) {
		leftShape = shapeA;
		rightShape = shapeB;
	} else {
		leftShape = shapeB;
		rightShape = shapeA;
	}

	// Update Graph IDs to establish relationship
	updateGraphShapeRelationship(leftShape, rightShape);

	const line = createSpecificConnection(
		leftShape,
		rightShape,
		"RIGHT",
		"TOP",
		lineType,
		startArrow,
		endArrow,
	);

	if (!line) {
		return SlidesApp.getUi().alert("Could not create connection.");
	}
}

/**
 * LUQ+RLQ Connection 2: Bottom of left shape → Left of right shape
 */
function connectLUQRLQ_BottomLeft(
	lineType = "BENT",
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

	// Determine which is left and which is right
	let leftShape, rightShape;
	if (shapeA.getLeft() < shapeB.getLeft()) {
		leftShape = shapeA;
		rightShape = shapeB;
	} else {
		leftShape = shapeB;
		rightShape = shapeA;
	}

	// Update Graph IDs to establish relationship
	updateGraphShapeRelationship(leftShape, rightShape);

	const line = createSpecificConnection(
		leftShape,
		rightShape,
		"BOTTOM",
		"LEFT",
		lineType,
		startArrow,
		endArrow,
	);

	if (!line) {
		return SlidesApp.getUi().alert("Could not create connection.");
	}
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

	// Check if both shapes already have Graph IDs
	const parentGraphId = getShapeGraphId(parentShape);
	const childGraphId = getShapeGraphId(childShape);

	// If both shapes already have Graph IDs, just connect them visually without modifying IDs
	if (parentGraphId && childGraphId) {
		// Both shapes have Graph IDs - don't modify them, just create the visual connection
		// This preserves existing hierarchy relationships
	} else if (!parentGraphId && !childGraphId) {
		// Neither shape has a Graph ID - initialize parent as root and child as its child
		const slide = parentShape.getParentPage();
		const nextRootId = findNextAvailableRootId(slide);
		const newParentGraphId = generateGraphId("", "", nextRootId, ["B1"]);
		setShapeGraphId(parentShape, newParentGraphId);

		// Set child's Graph ID
		const layout = orientation === "horizontal" ? "LR" : "TD";
		const newChildGraphId = generateGraphId(nextRootId, layout, "B1", []);
		setShapeGraphId(childShape, newChildGraphId);
	} else if (parentGraphId && !childGraphId) {
		// Only parent has Graph ID - create Graph ID for child
		const parentData = parseGraphId(parentGraphId);
		if (!parentData) {
			return SlidesApp.getUi().alert("Failed to parse parent Graph ID.");
		}

		// Generate child ID
		const nextLevel = getNextLevel(
			parentData.current.match(/^([A-Z]+)/)?.[1] || "A",
		);
		const existingChildren = parentData.children.filter((id) =>
			id.startsWith(nextLevel),
		);
		const nextNumber = existingChildren.length + 1;
		const childId = `${nextLevel}${nextNumber}`;

		// Set child's Graph ID
		const layout = orientation === "horizontal" ? "LR" : "TD";
		const newChildGraphId = generateGraphId(
			parentData.current,
			layout,
			childId,
			[],
		);
		setShapeGraphId(childShape, newChildGraphId);

		// Update parent to include this child
		const updatedChildren = [...parentData.children, childId];
		const updatedParentId = generateGraphId(
			parentData.parent,
			parentData.layout,
			parentData.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, updatedParentId);
	} else if (!parentGraphId && childGraphId) {
		// Only child has Graph ID - initialize parent as root but don't modify child
		const slide = parentShape.getParentPage();
		const nextRootId = findNextAvailableRootId(slide);
		const newParentGraphId = generateGraphId("", "", nextRootId, []);
		setShapeGraphId(parentShape, newParentGraphId);
		// Note: We're not updating the child's parent reference to preserve its existing relationships
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
