/**
 * Child creation utilities for creating child shapes in flowcharts
 * Handles positioning, styling, and connection of child shapes
 */

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
	const slide = parentShape.getParentPage();
	const parentProperties = {
		left: parentShape.getLeft(),
		top: parentShape.getTop(),
		width: parentShape.getWidth(),
		height: parentShape.getHeight(),
	};

	// Handle hierarchical naming and check for existing children
	let parentGraphId = getShapeGraphId(parentShape);
	let nextLevel = "B";
	const existingChildShapes = [];
	let parentData = null;

	if (parentGraphId) {
		parentData = parseGraphId(parentGraphId);
		if (parentData && parentData.current) {
			const levelMatch = parentData.current.match(/^([A-Z]+)/);
			if (levelMatch) {
				const parentLevel = levelMatch[1];
				nextLevel = getNextLevel(parentLevel);
			}

			// Check for existing children
			if (parentData.children && parentData.children.length > 0) {
				// Find existing child shapes on the slide
				const allShapes = slide.getShapes();
				for (const shape of allShapes) {
					const shapeGraphId = getShapeGraphId(shape);
					if (shapeGraphId) {
						const shapeData = parseGraphId(shapeGraphId);
						if (shapeData && parentData.children.includes(shapeData.current)) {
							existingChildShapes.push({
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
		}
	} else {
		// If parent doesn't have a graph ID, initialize it as root with next available A-level ID
		const nextRootId = findNextAvailableRootId(slide);
		parentGraphId = generateGraphId("", "", nextRootId, []);
		setShapeGraphId(parentShape, parentGraphId);
		parentData = parseGraphId(parentGraphId);
		nextLevel = "B";
	}

	// Determine starting child number based on existing children
	const existingChildrenOfLevel = parentData.children.filter((id) =>
		id.startsWith(nextLevel),
	);
	const startingNumber = existingChildrenOfLevel.length + 1;

	// Calculate positions for all children
	let positions;
	const childIds = [];

	if (existingChildShapes.length > 0) {
		// If there are existing children, position new ones as siblings
		positions = [];
		const layout = direction === "TOP" || direction === "BOTTOM" ? "TD" : "LR";

		// Sort existing children by position to find where to place new ones
		if (layout === "LR") {
			// For LR layout, sort by vertical position (top)
			existingChildShapes.sort((a, b) => a.top - b.top);
			const lastChild = existingChildShapes[existingChildShapes.length - 1];

			// Place new children below the last existing child
			for (let i = 0; i < count; i++) {
				positions.push({
					left: lastChild.left,
					top: lastChild.top + (i + 1) * (lastChild.height + gap),
				});
			}
		} else {
			// For TD layout, sort by horizontal position (left)
			existingChildShapes.sort((a, b) => a.left - b.left);
			const lastChild = existingChildShapes[existingChildShapes.length - 1];

			// Place new children to the right of the last existing child
			for (let i = 0; i < count; i++) {
				positions.push({
					left: lastChild.left + (i + 1) * (lastChild.width + gap),
					top: lastChild.top,
				});
			}
		}
	} else {
		// No existing children, use normal positioning
		positions = calculateChildPositions(
			parentProperties,
			direction,
			gap,
			count,
		);
	}

	// Create all children
	const createdShapes = [];

	for (let i = 0; i < count; i++) {
		const childShape = createSingleChild(
			parentShape,
			positions[i],
			direction,
			lineType,
			startArrow,
			endArrow,
		);

		// Generate unique child ID
		const childId = `${nextLevel}${startingNumber + i}`;
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

/**
 * Creates multiple child shapes in a specific direction with custom text for each shape
 * @param {string} direction - Direction to create children ("TOP", "RIGHT", "BOTTOM", "LEFT")
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
 * @returns {Array} - Array of created child shapes
 */
function createChildrenInDirectionWithText(
	direction,
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
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
	const slide = parentShape.getParentPage();
	const parentProperties = {
		left: parentShape.getLeft(),
		top: parentShape.getTop(),
		width: parentShape.getWidth(),
		height: parentShape.getHeight(),
	};

	// Use text count if texts are provided, otherwise use count parameter
	const actualCount = texts.length > 0 ? texts.length : count;

	// Check for existing children just like in createChildrenInDirection
	const parentGraphId = getShapeGraphId(parentShape);
	const existingChildShapes = [];
	let parentData = null;

	if (parentGraphId) {
		parentData = parseGraphId(parentGraphId);
		if (parentData && parentData.children && parentData.children.length > 0) {
			// Find existing child shapes on the slide
			const allShapes = slide.getShapes();
			for (const shape of allShapes) {
				const shapeGraphId = getShapeGraphId(shape);
				if (shapeGraphId) {
					const shapeData = parseGraphId(shapeGraphId);
					if (shapeData && parentData.children.includes(shapeData.current)) {
						existingChildShapes.push({
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
	}

	// Calculate positions for all children
	let positions;

	if (existingChildShapes.length > 0) {
		// If there are existing children, position new ones as siblings
		positions = [];
		const layout = direction === "TOP" || direction === "BOTTOM" ? "TD" : "LR";

		// Sort existing children by position to find where to place new ones
		if (layout === "LR") {
			// For LR layout, sort by vertical position (top)
			existingChildShapes.sort((a, b) => a.top - b.top);
			const lastChild = existingChildShapes[existingChildShapes.length - 1];

			// Place new children below the last existing child
			for (let i = 0; i < actualCount; i++) {
				positions.push({
					left: lastChild.left,
					top: lastChild.top + (i + 1) * (lastChild.height + gap),
				});
			}
		} else {
			// For TD layout, sort by horizontal position (left)
			existingChildShapes.sort((a, b) => a.left - b.left);
			const lastChild = existingChildShapes[existingChildShapes.length - 1];

			// Place new children to the right of the last existing child
			for (let i = 0; i < actualCount; i++) {
				positions.push({
					left: lastChild.left + (i + 1) * (lastChild.width + gap),
					top: lastChild.top,
				});
			}
		}
	} else {
		// No existing children, use normal positioning
		positions = calculateChildPositions(
			parentProperties,
			direction,
			gap,
			actualCount,
		);
	}

	const createdShapes = [];

	// Create each child shape with its text
	for (let i = 0; i < actualCount; i++) {
		const position = positions[i];
		const childText = texts[i] || ""; // Empty text if no custom text provided

		const childShape = createSingleChildWithText(
			parentShape,
			slide,
			position,
			direction,
			lineType,
			startArrow,
			endArrow,
			childText,
		);

		if (childShape) {
			createdShapes.push(childShape);
		}
	}

	return createdShapes;
}

/**
 * Creates a single child shape with custom text and styling
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape
 * @param {GoogleAppsScript.Slides.Slide} slide - Slide to create shape on
 * @param {Object} position - Position object {left, top}
 * @param {string} direction - Direction of child creation
 * @param {string} lineType - Type of line to use
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @param {string} text - Text to set in the child shape
 * @returns {GoogleAppsScript.Slides.Shape} - Created child shape
 */
function createSingleChildWithText(
	parentShape,
	slide,
	position,
	direction,
	lineType,
	startArrow,
	endArrow,
	text,
) {
	// Create new shape at calculated position
	const childShape = slide.insertShape(
		parentShape.getShapeType(),
		position.left,
		position.top,
		parentShape.getWidth(),
		parentShape.getHeight(),
	);

	// Copy styling from parent (except graph ID)
	copyShapeStyle(parentShape, childShape);

	// Set the custom text for this child shape (only if text is provided)
	try {
		if (text && text.trim() !== "") {
			childShape.getText().setText(text);
		}
		// If no text provided, leave the shape empty (don't set any text)
	} catch (e) {
		console.log(`Warning: Could not set child shape text: ${e.message}`);
	}

	// Set up graph ID for flowchart hierarchy
	let parentGraphId = getShapeGraphId(parentShape);

	// Initialize parent as root if it doesn't have a Graph ID
	if (!parentGraphId) {
		const slide = parentShape.getParentPage();
		const nextRootId = findNextAvailableRootId(slide);
		parentGraphId = generateGraphId("", "", nextRootId, []);
		setShapeGraphId(parentShape, parentGraphId);
	}

	const parentData = parseGraphId(parentGraphId);
	if (parentData) {
		// Generate child ID based on existing children
		const nextLevel = getNextLevel(
			parentData.current.match(/^([A-Z]+)/)?.[1] || "A",
		);
		const childNumber = parentData.children.length + 1;
		const childId = `${nextLevel}${childNumber}`;

		// Determine layout based on direction
		const layout = direction === "TOP" || direction === "BOTTOM" ? "TD" : "LR";

		// Set graph ID for child
		const childGraphId = generateGraphId(
			parentData.current,
			layout,
			childId,
			[],
		);
		setShapeGraphId(childShape, childGraphId);

		// Update parent to include this child
		const updatedChildren = [...parentData.children, childId];
		const updatedParentId = generateGraphId(
			parentData.parent,
			parentData.layout,
			parentData.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, updatedParentId);
	}

	// Create connection line
	const connectionSides = getConnectionSides(direction);
	const parentSite = pickConnectionSite(
		parentShape,
		connectionSides.parentSide,
	);
	const childSite = pickConnectionSite(childShape, connectionSides.childSide);

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
