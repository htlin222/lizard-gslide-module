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
			if (currentParsed && currentParsed.parent === parsed.parent) {
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
			if (currentParsed && currentParsed.current) {
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
				if (currentParsed && currentParsed.parent) {
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
