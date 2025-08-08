/**
 * Graph ID management utilities for hierarchical shape naming
 * Handles parsing, generating, and managing graph IDs for flowchart shapes
 * Format: graph[parent][current][children]
 */

/**
 * Parses a graph ID to extract its components
 * @param {string} graphId - The graph ID to parse
 * @returns {Object|null} - Object with parent, current, children arrays or null if not a graph ID
 */
function parseGraphId(graphId) {
	if (!graphId || !graphId.startsWith("graph[")) {
		return null;
	}

	// Extract content between brackets using regex
	const matches = graphId.match(/graph\[([^\]]*)\]\[([^\]]*)\]\[([^\]]*)\]/);
	if (!matches) {
		return null;
	}

	return {
		parent: matches[1] || "",
		current: matches[2] || "",
		children: matches[3] ? matches[3].split(",").filter((c) => c.trim()) : [],
	};
}

/**
 * Generates a graph ID from components
 * @param {string} parent - Parent ID
 * @param {string} current - Current ID
 * @param {Array} children - Array of child IDs
 * @returns {string} - Generated graph ID
 */
function generateGraphId(parent, current, children = []) {
	const childrenStr = children.join(",");
	return `graph[${parent}][${current}][${childrenStr}]`;
}

/**
 * Generates the next level ID for a child
 * @param {string} parentLevel - Parent's level (e.g., "A", "B", "C")
 * @returns {string} - Next level (e.g., "A" -> "B", "B" -> "C")
 */
function getNextLevel(parentLevel) {
	if (!parentLevel) return "A";
	const lastChar = parentLevel.charAt(parentLevel.length - 1);
	if (lastChar >= "A" && lastChar < "Z") {
		return (
			parentLevel.slice(0, -1) + String.fromCharCode(lastChar.charCodeAt(0) + 1)
		);
	}
	return parentLevel + "A";
}

/**
 * Generates sibling IDs for multiple children
 * @param {string} baseLevel - Base level for siblings (e.g., "C")
 * @param {number} count - Number of siblings to generate
 * @returns {Array} - Array of sibling IDs (e.g., ["C1", "C2", "C3"])
 */
function generateSiblingIds(baseLevel, count) {
	const siblings = [];
	for (let i = 1; i <= count; i++) {
		siblings.push(`${baseLevel}${i}`);
	}
	return siblings;
}

/**
 * Sets the graph ID on a shape by updating its text content
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to update
 * @param {string} graphId - New graph ID to set
 */
function setShapeGraphId(shape, graphId) {
	try {
		// Set the shape's text content to the graph ID
		shape.getText().setText(graphId);
	} catch (e) {
		console.log(`Warning: Could not set shape text: ${e.message}`);
	}
}

/**
 * Gets the graph ID from a shape's text content
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to read from
 * @returns {string|null} - Graph ID or null if not found
 */
function getShapeGraphId(shape) {
	try {
		// Get the shape's text content
		const text = shape.getText().asString().trim();
		if (text && text.startsWith("graph[")) {
			return text;
		}
		return null;
	} catch (e) {
		console.log(`Warning: Could not get shape text: ${e.message}`);
		return null;
	}
}

/**
 * Updates parent shape to include new children in its ID
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape to update
 * @param {Array} newChildIds - Array of new child IDs to add
 */
function updateParentWithChildren(parentShape, newChildIds) {
	const currentId = getShapeGraphId(parentShape);
	if (!currentId) {
		// If parent doesn't have a graph ID, create one
		const newId = generateGraphId("", "A1", newChildIds);
		setShapeGraphId(parentShape, newId);
		return;
	}

	const parsed = parseGraphId(currentId);
	if (!parsed) return;

	// Merge existing children with new ones
	const allChildren = [...parsed.children, ...newChildIds];
	const updatedId = generateGraphId(parsed.parent, parsed.current, allChildren);
	setShapeGraphId(parentShape, updatedId);
}

/**
 * Initializes a shape as a root graph node
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to initialize
 * @returns {string} - Generated root graph ID
 */
function initializeAsRootGraphShape(shape) {
	const rootId = generateGraphId("", "A1", []);
	setShapeGraphId(shape, rootId);
	return rootId;
}

/**
 * Determines parent-child relationship between two shapes based on their graph IDs
 * @param {string} idA - First shape's graph ID
 * @param {string} idB - Second shape's graph ID
 * @returns {Object|null} - Relationship info {parent, child, parentId, childId} or null if invalid
 */
function determineParentChildRelationship(idA, idB) {
	const parsedA = parseGraphId(idA);
	const parsedB = parseGraphId(idB);

	if (!parsedA || !parsedB) {
		return null;
	}

	// Determine parent-child relationship based on hierarchy level
	// Lower levels (A < B < C) are parents of higher levels
	const levelA = parsedA.current.match(/^([A-Z]+)/)?.[1] || "A";
	const levelB = parsedB.current.match(/^([A-Z]+)/)?.[1] || "A";

	if (levelA <= levelB) {
		return {
			parent: "A",
			child: "B",
			parentId: parsedA,
			childId: parsedB,
		};
	} else {
		return {
			parent: "B",
			child: "A",
			parentId: parsedB,
			childId: parsedA,
		};
	}
}

/**
 * Updates the relationship between two existing graph shapes
 * @param {GoogleAppsScript.Slides.Shape} shapeA - First shape
 * @param {GoogleAppsScript.Slides.Shape} shapeB - Second shape
 * @returns {Object|null} - Updated relationship or null if failed
 */
function updateGraphShapeRelationship(shapeA, shapeB) {
	const idA = getShapeGraphId(shapeA);
	const idB = getShapeGraphId(shapeB);

	if (!idA || !idB) {
		return null;
	}

	const relationship = determineParentChildRelationship(idA, idB);
	if (!relationship) {
		return null;
	}

	const parentShape = relationship.parent === "A" ? shapeA : shapeB;
	const childShape = relationship.parent === "A" ? shapeB : shapeA;
	const parentId = relationship.parentId;
	const childId = relationship.childId;

	// Update parent to include this child
	if (!parentId.children.includes(childId.current)) {
		const updatedChildren = [...parentId.children, childId.current];
		const newParentId = generateGraphId(
			parentId.parent,
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent
	const newChildId = generateGraphId(
		parentId.current,
		childId.current,
		childId.children,
	);
	setShapeGraphId(childShape, newChildId);

	return {
		parentShape,
		childShape,
		updated: true,
	};
}
