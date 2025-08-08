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

	// Generate new sibling ID based on parent's children list
	const parentData = parseGraphId(getShapeGraphId(parentShape));
	if (!parentData) {
		return SlidesApp.getUi().alert("Parent shape has invalid graph ID format.");
	}

	// Get the level from current selected shape (e.g., "C1" -> "C")
	const currentLevel = parsed.current.match(/^([A-Z]+)/)?.[1] || "A";

	// Find the highest numbered sibling from parent's children list
	let maxSiblingNumber = 0;

	// Check parent's children list first (most reliable source)
	for (const childId of parentData.children) {
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

	// Set the hierarchical graph ID using the parent's layout
	const newGraphId = generateGraphId(
		parsed.parent,
		parsed.layout,
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
	if (isHorizontalLayout && siblingShapes.length > 2) {
		// Horizontal layout: center all siblings around parent's X center
		const parentCenterX = parentShape.getLeft() + parentShape.getWidth() / 2;
		const totalGroupWidth =
			siblingShapes.length * selectedWidth +
			(siblingShapes.length - 1) * horizontalGap;
		const groupStartX = parentCenterX - totalGroupWidth / 2;

		siblingShapes.forEach((sibling, index) => {
			const newX = groupStartX + index * (selectedWidth + horizontalGap);
			sibling.shape.setLeft(newX);
		});
	} else {
		// Vertical layout (default): center all siblings around parent's Y center
		const parentCenterY = parentShape.getTop() + parentShape.getHeight() / 2;
		const totalGroupHeight =
			siblingShapes.length * selectedHeight +
			(siblingShapes.length - 1) * verticalGap;
		const groupStartY = parentCenterY - totalGroupHeight / 2;

		siblingShapes.forEach((sibling, index) => {
			const newY = groupStartY + index * (selectedHeight + verticalGap);
			sibling.shape.setTop(newY);
		});
	}

	// Update parent to include the new sibling
	const updatedChildren = [...parentData.children, newSiblingId];
	const updatedParentId = generateGraphId(
		parentData.parent,
		parentData.layout,
		parentData.current,
		updatedChildren,
	);
	setShapeGraphId(parentShape, updatedParentId);

	// Simple approach: Don't touch existing connections, just connect the new sibling with left/right pattern
	const parentSite = pickConnectionSite(parentShape, "RIGHT");
	const childSite = pickConnectionSite(newShape, "LEFT");

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
