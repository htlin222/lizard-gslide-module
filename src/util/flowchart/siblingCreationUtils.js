/**
 * Sibling creation utilities for flowchart shapes
 * Creates sibling shapes next to selected shapes with layout consistency
 */

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

	// Generate new sibling ID based on parent's children list
	const parentData = parseGraphId(getShapeGraphId(parentShape));
	if (!parentData) {
		return SlidesApp.getUi().alert("Parent shape has invalid graph ID format.");
	}

	// Determine layout from the parent's perspective
	// If parent has children, check the layout annotation in the first child's graph ID
	// Otherwise, use the selected shape's layout as fallback
	let layoutToUse = parsed.layout || "TD"; // Default to TD if no layout specified

	// The layout should be consistent across all siblings
	// Check existing siblings to determine the actual layout being used
	if (siblingShapes.length > 0) {
		// Use the layout from existing siblings
		const firstSibling = siblingShapes[0];
		if (firstSibling.data && firstSibling.data.layout) {
			layoutToUse = firstSibling.data.layout;
		}
	}

	const isHorizontalLayout = layoutToUse === "LR" || layoutToUse === "RL";
	const isVerticalLayout = layoutToUse === "TD" || layoutToUse === "DT";

	// Get the level from current selected shape (e.g., "C1" -> "C")
	const currentLevel = parsed.current.match(/^([A-Z]+)/)?.[1] || "A";

	// Find the highest numbered sibling from parent's children list
	let maxSiblingNumber = 0;

	// Check parent's children list first (most reliable source)
	for (const childId of parentData.childrenIds) {
		const childLevel = childId.match(/^([A-Z]+)(\d+)$/);
		if (childLevel && childLevel[1] === currentLevel) {
			const number = Number.parseInt(childLevel[2]);
			if (number > maxSiblingNumber) {
				maxSiblingNumber = number;
			}
		}
	}

	// Also check actual shapes on slide as backup
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

	// Create the new sibling shape first (we'll position it later)
	const newShape = slide.insertShape(
		selectedShape.getShapeType(),
		selectedLeft, // temporary position
		selectedTop, // temporary position
		selectedWidth,
		selectedHeight,
	);

	// Copy styling from selected shape
	copyShapeStyle(selectedShape, newShape);

	// Set the hierarchical graph ID using the determined layout
	const newGraphId = generateGraphId(
		parsed.parent,
		layoutToUse,
		newSiblingId,
		[],
	);
	setShapeGraphId(newShape, newGraphId);

	// Add the new sibling to our shapes array at the correct position
	// Find the index of the selected shape and insert the new one after it
	const selectedIndex = siblingShapes.findIndex(
		(sibling) => sibling.shape === selectedShape,
	);
	siblingShapes.splice(selectedIndex + 1, 0, {
		shape: newShape,
		data: parseGraphId(newGraphId),
		left: selectedLeft,
		top: selectedTop,
		width: selectedWidth,
		height: selectedHeight,
	});

	// Now reposition ALL siblings to be centered around the parent
	if (isHorizontalLayout) {
		// Horizontal layout (LR/RL): siblings spread vertically, all at same X position
		const parentCenterY = parentShape.getTop() + parentShape.getHeight() / 2;
		const totalGroupHeight =
			siblingShapes.length * selectedHeight +
			(siblingShapes.length - 1) * verticalGap;
		const groupStartY = parentCenterY - totalGroupHeight / 2;

		// Position siblings based on layout direction
		let siblingX;
		if (layoutToUse === "LR") {
			// LR: siblings positioned to the right of parent
			siblingX = parentShape.getLeft() + parentShape.getWidth() + horizontalGap;
		} else {
			// RL
			// RL: siblings positioned to the left of parent
			siblingX = parentShape.getLeft() - selectedWidth - horizontalGap;
		}

		siblingShapes.forEach((sibling, index) => {
			const newY = groupStartY + index * (selectedHeight + verticalGap);
			sibling.shape.setTop(newY);
			sibling.shape.setLeft(siblingX); // Set all siblings at the same X position
		});
	} else {
		// Vertical layout (TD/DT): siblings spread horizontally, all at same Y position
		const parentCenterX = parentShape.getLeft() + parentShape.getWidth() / 2;
		const totalGroupWidth =
			siblingShapes.length * selectedWidth +
			(siblingShapes.length - 1) * horizontalGap;
		const groupStartX = parentCenterX - totalGroupWidth / 2;

		// Position siblings based on layout direction
		let siblingY;
		if (layoutToUse === "TD") {
			// TD: siblings positioned below parent
			siblingY = parentShape.getTop() + parentShape.getHeight() + verticalGap;
		} else {
			// DT
			// DT: siblings positioned above parent
			siblingY = parentShape.getTop() - selectedHeight - verticalGap;
		}

		siblingShapes.forEach((sibling, index) => {
			const newX = groupStartX + index * (selectedWidth + horizontalGap);
			sibling.shape.setLeft(newX);
			sibling.shape.setTop(siblingY); // All siblings at same Y level
		});
	}

	// Update parent to include the new sibling
	const updatedChildren = [...parentData.childrenIds, newSiblingId];
	const updatedParentId = generateGraphId(
		parentData.parent,
		parentData.layout,
		parentData.current,
		updatedChildren,
	);
	setShapeGraphId(parentShape, updatedParentId);

	// Connect based on layout: LR/RL uses RIGHT/LEFT, TD/DT uses BOTTOM/TOP
	let parentSide, childSide;
	if (isHorizontalLayout) {
		if (layoutToUse === "LR") {
			// LR layout: parent connects from RIGHT, child connects from LEFT
			parentSide = "RIGHT";
			childSide = "LEFT";
		} else {
			// RL layout
			// RL layout: parent connects from LEFT, child connects from RIGHT
			parentSide = "LEFT";
			childSide = "RIGHT";
		}
	} else {
		if (layoutToUse === "TD") {
			// TD layout: parent connects from BOTTOM, child connects from TOP
			parentSide = "BOTTOM";
			childSide = "TOP";
		} else {
			// DT layout
			// DT layout: parent connects from TOP, child connects from BOTTOM
			parentSide = "TOP";
			childSide = "BOTTOM";
		}
	}

	const parentSite = pickConnectionSite(parentShape, parentSide);
	const childSite = pickConnectionSite(newShape, childSide);

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
