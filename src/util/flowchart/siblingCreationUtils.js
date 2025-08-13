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
	defaultStyle = null,
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

	// Extract the immediate parent from the hierarchy chain
	// For example: if parent is "A1|B1", the immediate parent is "B1"
	const parentHierarchy = parsed.parent;
	const immediateParent = parentHierarchy.includes("|")
		? parentHierarchy.split("|").pop()
		: parentHierarchy;

	// Find parent and all existing sibling shapes (including selected shape)
	for (const shape of allShapes) {
		const graphId = getShapeGraphId(shape);
		if (graphId) {
			const shapeData = parseGraphId(graphId);
			if (shapeData) {
				// Check if this is the parent
				if (shapeData.current === immediateParent) {
					parentShape = shape;
				}
				// Check if this is a sibling (same parent hierarchy)
				if (shapeData.parent === parentHierarchy) {
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

	// Generate new sibling ID based on parent's children list
	const parentData = parseGraphId(getShapeGraphId(parentShape));
	if (!parentData) {
		return SlidesApp.getUi().alert("Parent shape has invalid graph ID format.");
	}

	// Get selected shape dimensions early
	const selectedLeft = selectedShape.getLeft();
	const selectedTop = selectedShape.getTop();
	const selectedWidth = selectedShape.getWidth();
	const selectedHeight = selectedShape.getHeight();

	// Determine layout from the selected shape's own layout annotation
	// Look for the selected shape in parent's children to get its layout
	let layoutToUse = "LR"; // Default
	for (const child of parentData.children) {
		if (child.id === parsed.current) {
			layoutToUse = child.layout || "LR";
			break;
		}
	}

	// Group existing siblings by their actual layout from parent's children list
	const siblingsByLayout = new Map();
	for (const sibling of siblingShapes) {
		// Find this sibling's layout from parent's children list
		let siblingLayout = "LR"; // default
		for (const child of parentData.children) {
			if (child.id === sibling.data.current) {
				siblingLayout = child.layout || "LR";
				break;
			}
		}

		if (!siblingsByLayout.has(siblingLayout)) {
			siblingsByLayout.set(siblingLayout, []);
		}
		siblingsByLayout.get(siblingLayout).push({
			...sibling,
			actualLayout: siblingLayout, // Store the actual layout
		});
	}

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

	// Add placeholder text for styling purposes
	try {
		newShape.getText().setText("_");
	} catch (e) {
		console.log(`Warning: Could not set sibling shape text: ${e.message}`);
	}

	// Apply default style if specified (overrides inherited styling)
	if (defaultStyle) {
		try {
			applyStyleToShape(newShape, defaultStyle);
		} catch (e) {
			console.log(`Warning: Could not apply default style: ${e.message}`);
		}
	}

	// Set the hierarchical graph ID using the determined layout
	const newGraphId = generateGraphId(
		parsed.parent,
		layoutToUse,
		newSiblingId,
		[],
	);
	setShapeGraphId(newShape, newGraphId);

	// Create new sibling shape info and add it to the appropriate layout group
	const newSiblingInfo = {
		shape: newShape,
		data: parseGraphId(newGraphId),
		left: selectedLeft,
		top: selectedTop,
		width: selectedWidth,
		height: selectedHeight,
		actualLayout: layoutToUse,
	};

	// Add the new sibling to the appropriate layout group
	if (!siblingsByLayout.has(layoutToUse)) {
		siblingsByLayout.set(layoutToUse, []);
	}
	siblingsByLayout.get(layoutToUse).push(newSiblingInfo);

	// Now reposition siblings by layout groups
	// Each layout group positions its children independently and centers them to parent
	for (const [layout, layoutSiblings] of siblingsByLayout) {
		if (layoutSiblings.length === 0) continue;

		// Sort siblings by their current IDs for consistent ordering
		layoutSiblings.sort((a, b) => {
			const aNum = Number.parseInt(a.data.current.match(/\d+$/)?.[0] || "0");
			const bNum = Number.parseInt(b.data.current.match(/\d+$/)?.[0] || "0");
			return aNum - bNum;
		});

		const isHorizontal = layout === "LR" || layout === "RL";
		const parentCenterX = parentShape.getLeft() + parentShape.getWidth() / 2;
		const parentCenterY = parentShape.getTop() + parentShape.getHeight() / 2;

		if (isHorizontal) {
			// Horizontal layout (LR/RL): siblings spread vertically, centered to parent's Y
			const totalGroupHeight =
				layoutSiblings.length * selectedHeight +
				(layoutSiblings.length - 1) * verticalGap;
			const groupStartY = parentCenterY - totalGroupHeight / 2;

			// Position siblings based on layout direction
			let siblingX;
			if (layout === "LR") {
				// LR: siblings positioned to the right of parent
				siblingX =
					parentShape.getLeft() + parentShape.getWidth() + horizontalGap;
			} else {
				// RL: siblings positioned to the left of parent
				siblingX = parentShape.getLeft() - selectedWidth - horizontalGap;
			}

			layoutSiblings.forEach((sibling, index) => {
				const newY = groupStartY + index * (selectedHeight + verticalGap);
				sibling.shape.setTop(newY);
				sibling.shape.setLeft(siblingX);
			});
		} else {
			// Vertical layout (TD/DT): siblings spread horizontally, centered to parent's X
			const totalGroupWidth =
				layoutSiblings.length * selectedWidth +
				(layoutSiblings.length - 1) * horizontalGap;
			const groupStartX = parentCenterX - totalGroupWidth / 2;

			// Position siblings based on layout direction
			let siblingY;
			if (layout === "TD") {
				// TD: siblings positioned below parent
				siblingY = parentShape.getTop() + parentShape.getHeight() + verticalGap;
			} else {
				// DT: siblings positioned above parent
				siblingY = parentShape.getTop() - selectedHeight - verticalGap;
			}

			layoutSiblings.forEach((sibling, index) => {
				const newX = groupStartX + index * (selectedWidth + horizontalGap);
				sibling.shape.setLeft(newX);
				sibling.shape.setTop(siblingY);
			});
		}
	}

	// Update parent to include the new sibling with its layout
	const newChildWithLayout = { id: newSiblingId, layout: layoutToUse };
	const updatedChildren = [...parentData.children, newChildWithLayout];
	const updatedParentId = generateGraphId(
		parentData.parent,
		parentData.layout,
		parentData.current,
		updatedChildren,
	);
	setShapeGraphId(parentShape, updatedParentId);

	// Connect based on the new sibling's specific layout
	let parentSide, childSide;
	if (layoutToUse === "LR") {
		// LR layout: parent connects from RIGHT, child connects from LEFT
		parentSide = "RIGHT";
		childSide = "LEFT";
	} else if (layoutToUse === "RL") {
		// RL layout: parent connects from LEFT, child connects from RIGHT
		parentSide = "LEFT";
		childSide = "RIGHT";
	} else if (layoutToUse === "TD") {
		// TD layout: parent connects from BOTTOM, child connects from TOP
		parentSide = "BOTTOM";
		childSide = "TOP";
	} else if (layoutToUse === "DT") {
		// DT layout: parent connects from TOP, child connects from BOTTOM
		parentSide = "TOP";
		childSide = "BOTTOM";
	} else {
		// Default fallback to LR
		parentSide = "RIGHT";
		childSide = "LEFT";
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
