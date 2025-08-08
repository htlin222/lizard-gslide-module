/**
 * Child creation utilities for creating child shapes in flowcharts
 * Handles positioning, styling, and connection of child shapes
 */

/**
 * Validates that element is suitable for child creation
 * @param {GoogleAppsScript.Slides.PageElementRange} range - Selection range
 * @returns {Object} - Validation result with shape or error message
 */
function validateParentElement(range) {
	if (!range) {
		return { error: "Please select a shape to create a child for." };
	}

	const elements = range.getPageElements();
	if (elements.length !== 1) {
		return { error: "Please select exactly ONE shape." };
	}

	const element = elements[0];
	if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
		return { error: "Selected item must be a SHAPE." };
	}

	return { shape: element.asShape() };
}

/**
 * Calculates positions for multiple child shapes
 * @param {Object} parent - Parent shape properties (left, top, width, height)
 * @param {string} direction - Direction to create children (TOP, RIGHT, BOTTOM, LEFT)
 * @param {number} gap - Gap between shapes
 * @param {number} count - Number of children
 * @returns {Array} - Array of position objects {left, top}
 */
function calculateChildPositions(parent, direction, gap, count) {
	const positions = [];

	// Calculate sibling spacing offset
	let siblingOffset = 0;
	if (count > 1) {
		if (direction === "LEFT" || direction === "RIGHT") {
			// For horizontal directions, space siblings vertically
			const totalHeight = count * parent.height + (count - 1) * gap;
			siblingOffset = -(totalHeight - parent.height) / 2;
		} else {
			// For vertical directions, space siblings horizontally
			const totalWidth = count * parent.width + (count - 1) * gap;
			siblingOffset = -(totalWidth - parent.width) / 2;
		}
	}

	for (let i = 0; i < count; i++) {
		let childLeft = parent.left;
		let childTop = parent.top;

		switch (direction) {
			case "TOP":
				childTop = parent.top - parent.height - gap;
				childLeft = parent.left + siblingOffset + i * (parent.width + gap);
				break;
			case "RIGHT":
				childLeft = parent.left + parent.width + gap;
				childTop = parent.top + siblingOffset + i * (parent.height + gap);
				break;
			case "BOTTOM":
				childTop = parent.top + parent.height + gap;
				childLeft = parent.left + siblingOffset + i * (parent.width + gap);
				break;
			case "LEFT":
				childLeft = parent.left - parent.width - gap;
				childTop = parent.top + siblingOffset + i * (parent.height + gap);
				break;
		}

		positions.push({ left: childLeft, top: childTop });
	}

	return positions;
}

/**
 * Gets connection sides for parent-child connection
 * @param {string} direction - Direction of child creation
 * @returns {Object} - Connection sides {parentSide, childSide}
 */
function getConnectionSides(direction) {
	const connectionMap = {
		TOP: { parentSide: "TOP", childSide: "BOTTOM" },
		RIGHT: { parentSide: "RIGHT", childSide: "LEFT" },
		BOTTOM: { parentSide: "BOTTOM", childSide: "TOP" },
		LEFT: { parentSide: "LEFT", childSide: "RIGHT" },
	};
	return connectionMap[direction];
}

/**
 * Creates a single child shape with styling and connection
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape
 * @param {Object} position - Position {left, top}
 * @param {string} direction - Direction for connection
 * @param {string} lineType - Line type for connection
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @returns {GoogleAppsScript.Slides.Shape} - Created child shape
 */
function createSingleChild(
	parentShape,
	position,
	direction,
	lineType,
	startArrow,
	endArrow,
) {
	const slide = parentShape.getParentPage();

	// Create new shape
	const childShape = slide.insertShape(
		parentShape.getShapeType(),
		position.left,
		position.top,
		parentShape.getWidth(),
		parentShape.getHeight(),
	);

	// Copy styling from parent
	copyShapeStyle(parentShape, childShape);

	// Create connection
	const sides = getConnectionSides(direction);
	const parentSite = pickConnectionSite(parentShape, sides.parentSide);
	const childSite = pickConnectionSite(childShape, sides.childSide);

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

	return childShape;
}

/**
 * Main function to create child shapes in a specific direction
 * @param {string} direction - Direction to create children (TOP, RIGHT, BOTTOM, LEFT)
 * @param {number} gap - Gap between shapes
 * @param {string} lineType - Type of line for connections
 * @param {number} count - Number of children to create
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @returns {Array} - Array of created child shapes
 */
function createChildrenInDirection(
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

	const validation = validateParentElement(range);
	if (validation.error) {
		SlidesApp.getUi().alert(validation.error);
		return [];
	}

	const parentShape = validation.shape;
	const parentProperties = {
		left: parentShape.getLeft(),
		top: parentShape.getTop(),
		width: parentShape.getWidth(),
		height: parentShape.getHeight(),
	};

	// Calculate positions for all children
	const positions = calculateChildPositions(
		parentProperties,
		direction,
		gap,
		count,
	);

	// Create all children
	const createdShapes = [];
	const childIds = [];

	for (let i = 0; i < count; i++) {
		const childShape = createSingleChild(
			parentShape,
			positions[i],
			direction,
			lineType,
			startArrow,
			endArrow,
		);

		// Handle hierarchical naming
		const parentGraphId = getShapeGraphId(parentShape);
		let nextLevel = "B";

		if (parentGraphId) {
			const parsed = parseGraphId(parentGraphId);
			if (parsed && parsed.current) {
				const levelMatch = parsed.current.match(/^([A-Z]+)/);
				if (levelMatch) {
					const parentLevel = levelMatch[1];
					nextLevel = getNextLevel(parentLevel);
				}
			}
		} else {
			// If parent doesn't have a graph ID, make it the root
			nextLevel = "B";
		}

		// Generate unique child ID
		const childId = `${nextLevel}${i + 1}`;
		childIds.push(childId);

		// Set graph ID for child
		const parentCurrentId = parentGraphId
			? parseGraphId(parentGraphId)?.current || "A1"
			: "A1";

		// Determine layout based on direction
		const layout = direction === "TOP" || direction === "BOTTOM" ? "TD" : "LR";
		const childGraphId = generateGraphId(parentCurrentId, layout, childId, []);
		setShapeGraphId(childShape, childGraphId);

		createdShapes.push(childShape);
	}

	// Update parent with new children
	updateParentWithChildren(parentShape, childIds);

	return createdShapes;
}
