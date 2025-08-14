/**
 * Smart selection utilities for flowchart shapes
 * Provides selection functions based on graph ID relationships
 */

/**
 * Validates that a shape is selected and has a graph ID
 * @returns {Object} - Validation result with shape, slide, and graphId or error message
 */
function validateSelectedShape() {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return { error: "Please select a shape first." };
	}

	const elements = range.getPageElements();
	if (elements.length !== 1) {
		return { error: "Please select exactly ONE shape." };
	}

	const element = elements[0];
	if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
		return { error: "Selected item must be a SHAPE." };
	}

	const shape = element.asShape();
	const graphId = getShapeGraphId(shape);

	if (!graphId) {
		return {
			error:
				"Selected shape must have a graph ID. Please create it as part of a flowchart first.",
		};
	}

	const slide = shape.getParentPage();

	return {
		shape: shape,
		slide: slide,
		graphId: graphId,
	};
}

/**
 * Selects all sibling shapes that share the same parent
 */
function selectAllSiblings() {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.parent) {
		return;
	}

	const allShapes = slide.getShapes();
	const siblingShapes = [shape]; // Include the originally selected shape

	// Find all shapes with the same parent hierarchy
	for (const currentShape of allShapes) {
		if (currentShape === shape) continue; // Skip the originally selected shape (already included)

		const currentGraphId = getShapeGraphId(currentShape);
		if (currentGraphId) {
			const currentParsed = parseGraphId(currentGraphId);
			if (currentParsed?.parent === parsed.parent) {
				siblingShapes.push(currentShape);
			}
		}
	}

	if (siblingShapes.length === 1) {
		return;
	}

	// Select all sibling shapes
	try {
		// First select the first shape, then add the rest
		if (siblingShapes.length > 0) {
			siblingShapes[0].select();

			// Add remaining shapes to selection if more than one
			for (let i = 1; i < siblingShapes.length; i++) {
				siblingShapes[i].select(false); // false means don't replace existing selection
			}
		}
	} catch (e) {
		console.error(`Error selecting siblings: ${e.message}`);
	}
}

/**
 * Selects all shapes at the same level (same starting alphabet in graph ID)
 */
function selectAllLevel() {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.current) {
		return;
	}

	// Extract the level (starting alphabet) from the current ID
	const levelMatch = parsed.current.match(/^([A-Z]+)/);
	if (!levelMatch) {
		return;
	}

	const level = levelMatch[1];
	const allShapes = slide.getShapes();
	const levelShapes = [shape]; // Include the originally selected shape

	// Find all shapes that start with the same level
	for (const currentShape of allShapes) {
		if (currentShape === shape) continue; // Skip the originally selected shape (already included)

		const currentGraphId = getShapeGraphId(currentShape);
		if (currentGraphId) {
			const currentParsed = parseGraphId(currentGraphId);
			if (currentParsed?.current) {
				const currentLevelMatch = currentParsed.current.match(/^([A-Z]+)/);
				if (currentLevelMatch && currentLevelMatch[1] === level) {
					levelShapes.push(currentShape);
				}
			}
		}
	}

	if (levelShapes.length === 1) {
		return;
	}

	// Select all level shapes
	try {
		// First select the first shape, then add the rest
		if (levelShapes.length > 0) {
			levelShapes[0].select();

			// Add remaining shapes to selection if more than one
			for (let i = 1; i < levelShapes.length; i++) {
				levelShapes[i].select(false); // false means don't replace existing selection
			}
		}
	} catch (e) {
		console.error(`Error selecting level shapes: ${e.message}`);
	}
}

/**
 * Selects all parent shapes up to the root
 */
function selectAllParents() {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.parent) {
		return;
	}

	const allShapes = slide.getShapes();
	const parentShapes = [shape]; // Include the originally selected shape

	// Build the parent hierarchy chain
	const parentHierarchy = parsed.parent.split("|");

	// Find each parent in the hierarchy
	for (const parentId of parentHierarchy) {
		for (const currentShape of allShapes) {
			const currentGraphId = getShapeGraphId(currentShape);
			if (currentGraphId) {
				const currentParsed = parseGraphId(currentGraphId);
				if (currentParsed && currentParsed.current === parentId) {
					parentShapes.push(currentShape);
					break; // Found this parent, move to next
				}
			}
		}
	}

	if (parentShapes.length === 1) {
		return;
	}

	// Select all parent shapes
	try {
		// First select the first shape, then add the rest
		if (parentShapes.length > 0) {
			parentShapes[0].select();

			// Add remaining shapes to selection if more than one
			for (let i = 1; i < parentShapes.length; i++) {
				parentShapes[i].select(false); // false means don't replace existing selection
			}
		}
	} catch (e) {
		console.error(`Error selecting parent shapes: ${e.message}`);
	}
}

/**
 * Selects all family members (all descendants/children and grandchildren)
 */
function selectFamily() {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.current) {
		return;
	}

	const allShapes = slide.getShapes();
	const familyShapes = [shape]; // Include the originally selected shape

	/**
	 * Recursively find all descendants of a given shape ID
	 * @param {string} parentId - The ID to find descendants for
	 * @param {Set} visited - Set to track visited IDs to prevent infinite loops
	 */
	function findDescendants(parentId, visited = new Set()) {
		if (visited.has(parentId)) {
			return; // Prevent infinite loops
		}
		visited.add(parentId);

		for (const currentShape of allShapes) {
			if (currentShape === shape && familyShapes.includes(currentShape)) {
				continue; // Skip if already included
			}

			const currentGraphId = getShapeGraphId(currentShape);
			if (currentGraphId) {
				const currentParsed = parseGraphId(currentGraphId);
				if (currentParsed?.parent) {
					// Check if this shape is a direct child of parentId
					const parentHierarchy = currentParsed.parent.split("|");
					if (parentHierarchy[parentHierarchy.length - 1] === parentId) {
						// This is a direct child
						if (!familyShapes.includes(currentShape)) {
							familyShapes.push(currentShape);
						}
						// Recursively find its children
						findDescendants(currentParsed.current, visited);
					}
				}
			}
		}
	}

	// Start the recursive search from the selected shape
	findDescendants(parsed.current);

	if (familyShapes.length === 1) {
		return;
	}

	// Select all family shapes
	try {
		// First select the first shape, then add the rest
		if (familyShapes.length > 0) {
			familyShapes[0].select();

			// Add remaining shapes to selection if more than one
			for (let i = 1; i < familyShapes.length; i++) {
				familyShapes[i].select(false); // false means don't replace existing selection
			}
		}
	} catch (e) {
		console.error(`Error selecting family shapes: ${e.message}`);
	}
}

/**
 * Wrapper function to receive gap value from HTML
 * @param {number} gap - Gap between shapes in pixels
 */
function alignChildrenWithGap(gap = 20) {
	alignChildren(gap);
}

/**
 * Aligns all direct children of the selected shape
 * @param {number} gap - Gap between shapes in pixels (default 20)
 */
function alignChildren(gap = 20) {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.current) {
		return;
	}

	const allShapes = slide.getShapes();
	const childShapes = [];

	// Find all direct children (same pattern as selectFamily)
	for (const currentShape of allShapes) {
		const currentGraphId = getShapeGraphId(currentShape);
		if (currentGraphId) {
			const currentParsed = parseGraphId(currentGraphId);
			if (currentParsed?.parent) {
				// Check if this shape is a direct child
				const parentHierarchy = currentParsed.parent.split("|");
				if (parentHierarchy[parentHierarchy.length - 1] === parsed.current) {
					childShapes.push(currentShape);
				}
			}
		}
	}

	if (childShapes.length === 0) {
		return;
	}

	try {
		// Get parent position and dimensions
		const parentLeft = shape.getLeft();
		const parentTop = shape.getTop();
		const parentWidth = shape.getWidth();
		const parentHeight = shape.getHeight();
		const parentCenterX = parentLeft + parentWidth / 2;
		const parentCenterY = parentTop + parentHeight / 2;

		// Get child dimensions
		const childData = childShapes.map((child) => ({
			shape: child,
			width: child.getWidth(),
			height: child.getHeight(),
		}));

		if (parsed.layout === "TD" || parsed.layout === "DT") {
			// Top-Down layout: align children horizontally below parent
			const startY = parentTop + parentHeight + gap;
			const totalWidth =
				childData.reduce((sum, child) => sum + child.width, 0) +
				gap * (childData.length - 1);
			let currentX = parentCenterX - totalWidth / 2;

			childData.forEach((child) => {
				child.shape.setLeft(currentX);
				child.shape.setTop(startY);
				currentX += child.width + gap;
			});
		} else {
			// Left-Right layout: align children vertically to the right of parent
			const startX = parentLeft + parentWidth + gap;
			const totalHeight =
				childData.reduce((sum, child) => sum + child.height, 0) +
				gap * (childData.length - 1);
			let currentY = parentCenterY - totalHeight / 2;

			childData.forEach((child) => {
				child.shape.setLeft(startX);
				child.shape.setTop(currentY);
				currentY += child.height + gap;
			});
		}

		// Re-select the original shape
		shape.select();
	} catch (e) {
		console.error(`Error aligning children: ${e.message}`);
	}
}

/**
 * Aligns the parent shape based on the center of its children
 */
function alignWithParent() {
	const validation = validateSelectedShape();
	if (validation.error) {
		return;
	}

	const { shape, slide, graphId } = validation;
	const parsed = parseGraphId(graphId);

	if (!parsed || !parsed.parent) {
		return;
	}

	// Get the immediate parent ID
	const parentHierarchy = parsed.parent.split("|");
	const immediateParentId = parentHierarchy[parentHierarchy.length - 1];

	const allShapes = slide.getShapes();
	let parentShape = null;
	let parentLayout = null;

	// Find the parent shape
	for (const currentShape of allShapes) {
		const currentGraphId = getShapeGraphId(currentShape);
		if (currentGraphId) {
			const currentParsed = parseGraphId(currentGraphId);
			if (currentParsed && currentParsed.current === immediateParentId) {
				parentShape = currentShape;
				parentLayout = currentParsed.layout;
				break;
			}
		}
	}

	if (!parentShape) {
		return;
	}

	// Find all siblings (including selected shape)
	const siblingShapes = [shape];
	for (const currentShape of allShapes) {
		if (currentShape === shape) continue;

		const currentGraphId = getShapeGraphId(currentShape);
		if (currentGraphId) {
			const currentParsed = parseGraphId(currentGraphId);
			if (currentParsed?.parent) {
				const currentParentHierarchy = currentParsed.parent.split("|");
				if (
					currentParentHierarchy[currentParentHierarchy.length - 1] ===
					immediateParentId
				) {
					siblingShapes.push(currentShape);
				}
			}
		}
	}

	if (siblingShapes.length === 0) {
		return;
	}

	try {
		// Calculate bounding box of all siblings
		let minX = Number.POSITIVE_INFINITY;
		let maxX = Number.NEGATIVE_INFINITY;
		let minY = Number.POSITIVE_INFINITY;
		let maxY = Number.NEGATIVE_INFINITY;

		for (const sibling of siblingShapes) {
			const left = sibling.getLeft();
			const top = sibling.getTop();
			const right = left + sibling.getWidth();
			const bottom = top + sibling.getHeight();

			minX = Math.min(minX, left);
			maxX = Math.max(maxX, right);
			minY = Math.min(minY, top);
			maxY = Math.max(maxY, bottom);
		}

		// Calculate center of siblings
		const siblingsCenterX = (minX + maxX) / 2;
		const siblingsCenterY = (minY + maxY) / 2;

		// Move parent to align with siblings center
		const parentWidth = parentShape.getWidth();
		const parentHeight = parentShape.getHeight();

		if (parentLayout === "TD" || parentLayout === "DT") {
			// Top-Down: align parent horizontally with siblings center
			const newParentLeft = siblingsCenterX - parentWidth / 2;
			parentShape.setLeft(newParentLeft);
		} else {
			// Left-Right: align parent vertically with siblings center
			const newParentTop = siblingsCenterY - parentHeight / 2;
			parentShape.setTop(newParentTop);
		}

		// Re-select the original shape
		shape.select();
	} catch (e) {
		console.error(`Error aligning parent: ${e.message}`);
	}
}
