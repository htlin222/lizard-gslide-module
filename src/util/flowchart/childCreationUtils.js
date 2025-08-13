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
					usedRootIds.add(Number.parseInt(match[1]));
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
 * Finds the appropriate parent shape for child creation
 * If a child shape is selected, it finds its parent instead
 * @param {GoogleAppsScript.Slides.Shape} selectedShape - Initially selected shape
 * @param {GoogleAppsScript.Slides.Slide} slide - Current slide
 * @returns {GoogleAppsScript.Slides.Shape} - The appropriate parent shape
 */
function findAppropriateParent(selectedShape, slide) {
	const selectedGraphId = getShapeGraphId(selectedShape);
	if (!selectedGraphId) {
		// No graph ID, use as-is
		return selectedShape;
	}

	const parsed = parseGraphId(selectedGraphId);
	if (!parsed || !parsed.parent) {
		// No parent, this is already a root shape
		return selectedShape;
	}

	// This is a child shape, find its parent
	const allShapes = slide.getShapes();
	for (const shape of allShapes) {
		const graphId = getShapeGraphId(shape);
		if (graphId) {
			const shapeData = parseGraphId(graphId);
			if (shapeData && shapeData.current === parsed.parent) {
				return shape; // Found the parent
			}
		}
	}

	// Parent not found, use the selected shape
	return selectedShape;
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
 * @param {number} customWidth - Custom width for children (optional)
 * @param {number} customHeight - Custom height for children (optional)
 * @returns {Array} - Array of position objects {left, top}
 */
function calculateChildPositions(
	parent,
	direction,
	gap,
	count,
	customWidth,
	customHeight,
) {
	// Use custom dimensions if provided, otherwise use parent dimensions
	const childWidth = customWidth || parent.width;
	const childHeight = customHeight || parent.height;
	const positions = [];

	// Calculate sibling spacing offset using child dimensions
	let siblingOffset = 0;
	if (count > 1) {
		if (direction === "LEFT" || direction === "RIGHT") {
			// For horizontal directions, space siblings vertically
			const totalHeight = count * childHeight + (count - 1) * gap;
			siblingOffset = -(totalHeight - parent.height) / 2;
		} else {
			// For vertical directions, space siblings horizontally
			const totalWidth = count * childWidth + (count - 1) * gap;
			siblingOffset = -(totalWidth - parent.width) / 2;
		}
	}

	for (let i = 0; i < count; i++) {
		let childLeft = parent.left;
		let childTop = parent.top;

		switch (direction) {
			case "TOP":
				childTop = parent.top - childHeight - gap;
				childLeft = parent.left + siblingOffset + i * (childWidth + gap);
				break;
			case "RIGHT":
				childLeft = parent.left + parent.width + gap;
				childTop = parent.top + siblingOffset + i * (childHeight + gap);
				break;
			case "BOTTOM":
				childTop = parent.top + parent.height + gap;
				childLeft = parent.left + siblingOffset + i * (childWidth + gap);
				break;
			case "LEFT":
				childLeft = parent.left - childWidth - gap;
				childTop = parent.top + siblingOffset + i * (childHeight + gap);
				break;
		}

		positions.push({ left: childLeft, top: childTop });
	}

	return positions;
}

/**
 * Gets layout type based on direction
 * @param {string} direction - Direction to create child (TOP, RIGHT, BOTTOM, LEFT)
 * @returns {string} - Layout type (TD, DT, LR, RL)
 */
function getLayoutFromDirection(direction) {
	const layoutMap = {
		TOP: "DT", // Down-Top: parent at bottom, child at top
		RIGHT: "LR", // Left-Right: parent at left, child at right
		BOTTOM: "TD", // Top-Down: parent at top, child at bottom
		LEFT: "RL", // Right-Left: parent at right, child at left
	};
	return layoutMap[direction] || "LR"; // Default to LR for compatibility
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
 * @param {number} customWidth - Custom width for child shape (optional)
 * @param {number} customHeight - Custom height for child shape (optional)
 * @returns {GoogleAppsScript.Slides.Shape} - Created child shape
 */
function createSingleChild(
	parentShape,
	position,
	direction,
	lineType,
	startArrow,
	endArrow,
	customWidth,
	customHeight,
) {
	const slide = parentShape.getParentPage();

	// Use custom dimensions if provided, otherwise use parent dimensions
	const width = customWidth || parentShape.getWidth();
	const height = customHeight || parentShape.getHeight();

	// Create new shape
	const childShape = slide.insertShape(
		parentShape.getShapeType(),
		position.left,
		position.top,
		width,
		height,
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
 * Repositions all children of a parent by their layout groups, centering each group to the parent
 * This uses the same robust logic as the sibling creation for consistency
 * @param {GoogleAppsScript.Slides.Shape} parentShape - The parent shape
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide containing the shapes
 * @param {number} gap - Gap between shapes (used for both horizontal and vertical gaps)
 */
function repositionChildrenByLayout(parentShape, slide, gap = 20) {
	const parentGraphId = getShapeGraphId(parentShape);
	if (!parentGraphId) return;

	const parentData = parseGraphId(parentGraphId);
	if (!parentData || !parentData.children || parentData.children.length === 0)
		return;

	// Find all child shapes on the slide
	const allShapes = slide.getShapes();
	const childShapes = [];

	for (const shape of allShapes) {
		const graphId = getShapeGraphId(shape);
		if (graphId) {
			const shapeData = parseGraphId(graphId);
			if (shapeData && shapeData.parent === parentData.current) {
				// Find this child's layout from parent's children list
				let childLayout = "LR"; // default
				for (const child of parentData.children) {
					if (child.id === shapeData.current) {
						childLayout = child.layout || "LR";
						break;
					}
				}

				childShapes.push({
					shape: shape,
					data: shapeData,
					left: shape.getLeft(),
					top: shape.getTop(),
					width: shape.getWidth(),
					height: shape.getHeight(),
					actualLayout: childLayout,
				});
			}
		}
	}

	if (childShapes.length === 0) return;

	// Group children by their actual layout from parent's children list
	const childrenByLayout = new Map();
	for (const child of childShapes) {
		const layout = child.actualLayout;
		if (!childrenByLayout.has(layout)) {
			childrenByLayout.set(layout, []);
		}
		childrenByLayout.get(layout).push(child);
	}

	// Use the same dimensions for all shapes (from the first child)
	const firstChild = childShapes[0];
	const shapeWidth = firstChild.width;
	const shapeHeight = firstChild.height;

	// Reposition each layout group independently and center them to parent
	for (const [layout, layoutChildren] of childrenByLayout) {
		if (layoutChildren.length === 0) continue;

		// Sort children by their current IDs for consistent ordering
		layoutChildren.sort((a, b) => {
			const aNum = Number.parseInt(a.data.current.match(/\d+$/)?.[0] || "0");
			const bNum = Number.parseInt(b.data.current.match(/\d+$/)?.[0] || "0");
			return aNum - bNum;
		});

		const isHorizontal = layout === "LR" || layout === "RL";
		const parentCenterX = parentShape.getLeft() + parentShape.getWidth() / 2;
		const parentCenterY = parentShape.getTop() + parentShape.getHeight() / 2;

		if (isHorizontal) {
			// Horizontal layout (LR/RL): children spread vertically, centered to parent's Y
			const totalGroupHeight =
				layoutChildren.length * shapeHeight + (layoutChildren.length - 1) * gap;
			const groupStartY = parentCenterY - totalGroupHeight / 2;

			// Position children based on layout direction
			let childX;
			if (layout === "LR") {
				// LR: children positioned to the right of parent
				childX = parentShape.getLeft() + parentShape.getWidth() + gap;
			} else {
				// RL: children positioned to the left of parent
				childX = parentShape.getLeft() - shapeWidth - gap;
			}

			layoutChildren.forEach((child, index) => {
				const newY = groupStartY + index * (shapeHeight + gap);
				child.shape.setTop(newY);
				child.shape.setLeft(childX);
			});
		} else {
			// Vertical layout (TD/DT): children spread horizontally, centered to parent's X
			const totalGroupWidth =
				layoutChildren.length * shapeWidth + (layoutChildren.length - 1) * gap;
			const groupStartX = parentCenterX - totalGroupWidth / 2;

			// Position children based on layout direction
			let childY;
			if (layout === "TD") {
				// TD: children positioned below parent
				childY = parentShape.getTop() + parentShape.getHeight() + gap;
			} else {
				// DT: children positioned above parent
				childY = parentShape.getTop() - shapeHeight - gap;
			}

			layoutChildren.forEach((child, index) => {
				const newX = groupStartX + index * (shapeWidth + gap);
				child.shape.setLeft(newX);
				child.shape.setTop(childY);
			});
		}
	}
}

/**
 * Main function to create child shapes in a specific direction
 * @param {string} direction - Direction to create children (TOP, RIGHT, BOTTOM, LEFT)
 * @param {number} gap - Gap between shapes
 * @param {string} lineType - Type of line for connections
 * @param {number} count - Number of children to create
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 * @param {number} customWidth - Custom width for children (optional)
 * @param {number} customHeight - Custom height for children (optional)
 * @param {boolean} maxWidth - Whether to use max width calculation
 * @param {boolean} maxHeight - Whether to use max height calculation
 * @returns {Array} - Array of created child shapes
 */
function createChildrenInDirection(
	direction,
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	customWidth = null,
	customHeight = null,
	maxWidth = false,
	maxHeight = false,
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	const validation = validateParentElement(range);
	if (validation.error) {
		SlidesApp.getUi().alert(validation.error);
		return [];
	}

	const selectedShape = validation.shape;
	const slide = selectedShape.getParentPage();

	// Use the selected shape directly as the parent for creating children
	const parentShape = selectedShape;
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
			if (parentData.childrenIds && parentData.childrenIds.length > 0) {
				// Find existing child shapes on the slide
				const allShapes = slide.getShapes();
				for (const shape of allShapes) {
					const shapeGraphId = getShapeGraphId(shape);
					if (shapeGraphId) {
						const shapeData = parseGraphId(shapeGraphId);
						if (
							shapeData &&
							parentData.childrenIds.includes(shapeData.current)
						) {
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

	// Calculate actual dimensions based on max settings
	let finalCustomWidth = customWidth;
	let finalCustomHeight = customHeight;

	if (maxWidth && customWidth && count > 0) {
		// For LR/RL layouts, width is distributed among children
		const layout = getLayoutFromDirection(direction);
		if (layout === "LR" || layout === "RL") {
			// Width doesn't need division for LR/RL (children are stacked vertically)
			finalCustomWidth = customWidth;
		} else {
			// TD/DT: divide width among children
			finalCustomWidth = Math.floor((customWidth - (count - 1) * gap) / count);
		}
	}

	if (maxHeight && customHeight && count > 0) {
		// For TD/DT layouts, height is distributed among children
		const layout = getLayoutFromDirection(direction);
		if (layout === "TD" || layout === "DT") {
			// Height doesn't need division for TD/DT (children are stacked horizontally)
			finalCustomHeight = customHeight;
		} else {
			// LR/RL: divide height among children
			finalCustomHeight = Math.floor(
				(customHeight - (count - 1) * gap) / count,
			);
		}
	}

	// Determine starting child number based on existing children
	const existingChildrenOfLevel = parentData.childrenIds.filter((id) =>
		id.startsWith(nextLevel),
	);
	const startingNumber = existingChildrenOfLevel.length + 1;

	// Calculate positions for all children
	let positions;
	const childIds = [];

	if (existingChildShapes.length > 0) {
		// Check if there are existing children with the same layout
		const layout = getLayoutFromDirection(direction);

		// Pre-calculate parent boundaries for efficiency
		const parentRect = {
			left: parentProperties.left,
			top: parentProperties.top,
			right: parentProperties.left + parentProperties.width,
			bottom: parentProperties.top + parentProperties.height,
		};

		const existingChildrenWithSameLayout = existingChildShapes.filter(
			(child) => {
				// Check if this child has the same layout as the new children
				if (child.data && child.data.layout) {
					return child.data.layout === layout;
				}

				// For children without explicit layout, quickly infer from position
				const childCenterX = child.left + child.width / 2;
				const childCenterY = child.top + child.height / 2;

				// Fast layout inference based on position
				if (layout === "LR" && childCenterX > parentRect.right) return true;
				if (layout === "RL" && childCenterX < parentRect.left) return true;
				if (layout === "TD" && childCenterY > parentRect.bottom) return true;
				if (layout === "DT" && childCenterY < parentRect.top) return true;

				return false;
			},
		);

		// We'll handle positioning after creating all children using the robust sibling logic
		positions = [];
		for (let i = 0; i < count; i++) {
			// Create temporary positions - we'll reposition everything later
			positions.push({
				left: parentProperties.left + 50 + i * 20,
				top: parentProperties.top + 50 + i * 20,
			});
		}
	} else {
		// No existing children, create temporary positions
		positions = [];
		for (let i = 0; i < count; i++) {
			positions.push({
				left: parentProperties.left + 50 + i * 20,
				top: parentProperties.top + 50 + i * 20,
			});
		}
	}

	// Create all children
	const createdShapes = [];
	const layout = getLayoutFromDirection(direction);

	for (let i = 0; i < count; i++) {
		const childShape = createSingleChild(
			parentShape,
			positions[i],
			direction,
			lineType,
			startArrow,
			endArrow,
			finalCustomWidth,
			finalCustomHeight,
		);

		// Generate unique child ID
		const childId = `${nextLevel}${startingNumber + i}`;
		// Store child with its layout
		childIds.push({ id: childId, layout: layout });

		// Set graph ID for child with full parent hierarchy
		const parsedParent = parentGraphId ? parseGraphId(parentGraphId) : null;
		const parentCurrentId = parsedParent?.current || "A1";

		// Build parent hierarchy chain
		let parentHierarchy = "";
		if (parsedParent) {
			if (parsedParent.parent) {
				// Parent already has hierarchy, append current parent
				parentHierarchy = `${parsedParent.parent}|${parentCurrentId}`;
			} else {
				// Parent is root, just use its ID
				parentHierarchy = parentCurrentId;
			}
		}

		const childGraphId = generateGraphId(parentHierarchy, layout, childId, []);
		setShapeGraphId(childShape, childGraphId);

		createdShapes.push(childShape);
	}

	// Update parent with new children (now includes layout info)
	updateParentWithChildren(parentShape, childIds);

	// Apply robust positioning: group all children by layout and center them to parent
	repositionChildrenByLayout(parentShape, slide, gap);

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
 * @param {number} customWidth - Custom width for children (optional)
 * @param {number} customHeight - Custom height for children (optional)
 * @param {boolean} maxWidth - Whether to use max width calculation
 * @param {boolean} maxHeight - Whether to use max height calculation
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
	customWidth = null,
	customHeight = null,
	maxWidth = false,
	maxHeight = false,
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	const validation = validateParentElement(range);
	if (validation.error) {
		SlidesApp.getUi().alert(validation.error);
		return [];
	}

	const selectedShape = validation.shape;
	const slide = selectedShape.getParentPage();

	// Use the selected shape directly as the parent for creating children
	const parentShape = selectedShape;
	const parentProperties = {
		left: parentShape.getLeft(),
		top: parentShape.getTop(),
		width: parentShape.getWidth(),
		height: parentShape.getHeight(),
	};

	// Use text count if texts are provided, otherwise use count parameter
	const actualCount = texts.length > 0 ? texts.length : count;

	// Calculate actual dimensions based on max settings
	let finalCustomWidth = customWidth;
	let finalCustomHeight = customHeight;

	if (maxWidth && customWidth && actualCount > 0) {
		// For LR/RL layouts, width is distributed among children
		const layout = getLayoutFromDirection(direction);
		if (layout === "LR" || layout === "RL") {
			// Width doesn't need division for LR/RL (children are stacked vertically)
			finalCustomWidth = customWidth;
		} else {
			// TD/DT: divide width among children
			finalCustomWidth = Math.floor(
				(customWidth - (actualCount - 1) * gap) / actualCount,
			);
		}
	}

	if (maxHeight && customHeight && actualCount > 0) {
		// For TD/DT layouts, height is distributed among children
		const layout = getLayoutFromDirection(direction);
		if (layout === "TD" || layout === "DT") {
			// Height doesn't need division for TD/DT (children are stacked horizontally)
			finalCustomHeight = customHeight;
		} else {
			// LR/RL: divide height among children
			finalCustomHeight = Math.floor(
				(customHeight - (actualCount - 1) * gap) / actualCount,
			);
		}
	}

	// Check for existing children just like in createChildrenInDirection
	const parentGraphId = getShapeGraphId(parentShape);
	const existingChildShapes = [];
	let parentData = null;

	if (parentGraphId) {
		parentData = parseGraphId(parentGraphId);
		if (
			parentData &&
			parentData.childrenIds &&
			parentData.childrenIds.length > 0
		) {
			// Find existing child shapes on the slide
			const allShapes = slide.getShapes();
			for (const shape of allShapes) {
				const shapeGraphId = getShapeGraphId(shape);
				if (shapeGraphId) {
					const shapeData = parseGraphId(shapeGraphId);
					if (shapeData && parentData.childrenIds.includes(shapeData.current)) {
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
		// Check if there are existing children with the same layout
		const layout = getLayoutFromDirection(direction);

		// Pre-calculate parent boundaries for efficiency
		const parentRect = {
			left: parentProperties.left,
			top: parentProperties.top,
			right: parentProperties.left + parentProperties.width,
			bottom: parentProperties.top + parentProperties.height,
		};

		const existingChildrenWithSameLayout = existingChildShapes.filter(
			(child) => {
				// Check if this child has the same layout as the new children
				if (child.data && child.data.layout) {
					return child.data.layout === layout;
				}

				// For children without explicit layout, quickly infer from position
				const childCenterX = child.left + child.width / 2;
				const childCenterY = child.top + child.height / 2;

				// Fast layout inference based on position
				if (layout === "LR" && childCenterX > parentRect.right) return true;
				if (layout === "RL" && childCenterX < parentRect.left) return true;
				if (layout === "TD" && childCenterY > parentRect.bottom) return true;
				if (layout === "DT" && childCenterY < parentRect.top) return true;

				return false;
			},
		);

		if (existingChildrenWithSameLayout.length > 0) {
			// Position new children as siblings to existing ones with same layout
			positions = [];

			// Sort existing children by position to find where to place new ones
			if (layout === "LR" || layout === "RL") {
				// For horizontal layouts, sort by vertical position (top)
				existingChildrenWithSameLayout.sort((a, b) => a.top - b.top);
				const lastChild =
					existingChildrenWithSameLayout[
						existingChildrenWithSameLayout.length - 1
					];

				// Place new children below the last existing child
				for (let i = 0; i < actualCount; i++) {
					positions.push({
						left: lastChild.left,
						top: lastChild.top + (i + 1) * (lastChild.height + gap),
					});
				}
			} else {
				// TD or DT
				// For vertical layouts, sort by horizontal position (left)
				existingChildrenWithSameLayout.sort((a, b) => a.left - b.left);
				const lastChild =
					existingChildrenWithSameLayout[
						existingChildrenWithSameLayout.length - 1
					];

				// Place new children to the right of the last existing child
				for (let i = 0; i < actualCount; i++) {
					positions.push({
						left: lastChild.left + (i + 1) * (lastChild.width + gap),
						top: lastChild.top,
					});
				}
			}
		} else {
			// No existing children with same layout, use normal positioning
			positions = calculateChildPositions(
				parentProperties,
				direction,
				gap,
				actualCount,
				finalCustomWidth,
				finalCustomHeight,
			);
		}
	} else {
		// No existing children, use normal positioning
		positions = calculateChildPositions(
			parentProperties,
			direction,
			gap,
			actualCount,
			finalCustomWidth,
			finalCustomHeight,
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
			finalCustomWidth,
			finalCustomHeight,
		);

		if (childShape) {
			createdShapes.push(childShape);
		}
	}

	// Apply robust positioning: group all children by layout and center them to parent
	repositionChildrenByLayout(parentShape, slide, gap);

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
 * @param {number} customWidth - Custom width for child shape (optional)
 * @param {number} customHeight - Custom height for child shape (optional)
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
	customWidth,
	customHeight,
) {
	// Use custom dimensions if provided, otherwise use parent dimensions
	const width = customWidth || parentShape.getWidth();
	const height = customHeight || parentShape.getHeight();

	// Create new shape at calculated position
	const childShape = slide.insertShape(
		parentShape.getShapeType(),
		position.left,
		position.top,
		width,
		height,
	);

	// Set the custom text for this child shape first (if text is provided)
	try {
		if (text && text.trim() !== "") {
			childShape.getText().setText(text);
		}
		// If no text provided, leave the shape empty (don't set any text)
	} catch (e) {
		console.log(`Warning: Could not set child shape text: ${e.message}`);
	}

	// Copy styling from parent AFTER setting text (to preserve text styling)
	copyShapeStyle(parentShape, childShape);

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
		const childNumber = parentData.childrenIds.length + 1;
		const childId = `${nextLevel}${childNumber}`;

		// Determine layout based on direction
		const layout = getLayoutFromDirection(direction);

		// Set graph ID for child with full parent hierarchy
		// Build parent hierarchy chain
		let parentHierarchy = "";
		if (parentData.parent) {
			// Parent already has hierarchy, append current parent
			parentHierarchy = `${parentData.parent}|${parentData.current}`;
		} else {
			// Parent is root, just use its ID
			parentHierarchy = parentData.current;
		}

		const childGraphId = generateGraphId(parentHierarchy, layout, childId, []);
		setShapeGraphId(childShape, childGraphId);

		// Update parent to include this child with its layout
		const childWithLayout = { id: childId, layout: layout };
		const updatedChildren = [...parentData.children, childWithLayout];
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
