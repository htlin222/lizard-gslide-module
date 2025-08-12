/**
 * Graph ID management utilities for hierarchical shape naming
 * Handles parsing, generating, and managing graph IDs for flowchart shapes
 * Format: graph[parent1|parent2|...](layout)[current][children]
 * - Parent field contains full parent hierarchy separated by |
 * - Layout types: LR (Left-Right), TD (Top-Down), RL (Right-Left), DT (Down-Top)
 * - Children can have individual layouts: [B1:RL,B2:TD,B3]
 * - Child without layout inherits parent's layout
 * Example: graph[A1|B1](LR)[C1][] means C1's parent is B1, and B1's parent is A1
 */

/**
 * Parses a child ID to extract its ID and optional layout
 * @param {string} childStr - Child string like "B1" or "B1:RL"
 * @returns {Object} - Object with {id, layout}
 */
function parseChildWithLayout(childStr) {
	const parts = childStr.trim().split(":");
	return {
		id: parts[0],
		layout: parts[1] || null, // null means inherit from parent
	};
}

/**
 * Parses a graph ID to extract its components
 * @param {string} graphId - The graph ID to parse
 * @returns {Object|null} - Object with parent, layout, current, children arrays or null if not a graph ID
 */
function parseGraphId(graphId) {
	if (!graphId || !graphId.startsWith("graph[")) {
		return null;
	}

	// New format: graph[parent](layout)[current][children]
	const newFormatMatches = graphId.match(
		/graph\[([^\]]*)\]\(([^)]*)\)\[([^\]]*)\]\[([^\]]*)\]/,
	);
	if (newFormatMatches) {
		const childrenStr = newFormatMatches[4] || "";
		const children = childrenStr
			? childrenStr
					.split(",")
					.filter((c) => c.trim())
					.map(parseChildWithLayout)
			: [];

		return {
			parent: newFormatMatches[1] || "",
			layout: newFormatMatches[2] || "LR", // Default to LR
			current: newFormatMatches[3] || "",
			children: children,
			// Keep backward compatibility with simple array
			childrenIds: children.map((c) => c.id),
		};
	}

	// Legacy format: graph[parent][current][children] - assume LR layout
	const legacyMatches = graphId.match(
		/graph\[([^\]]*)\]\[([^\]]*)\]\[([^\]]*)\]/,
	);
	if (legacyMatches) {
		const childrenStr = legacyMatches[3] || "";
		const children = childrenStr
			? childrenStr
					.split(",")
					.filter((c) => c.trim())
					.map((c) => ({ id: c.trim(), layout: null }))
			: [];

		return {
			parent: legacyMatches[1] || "",
			layout: "LR", // Default legacy to LR
			current: legacyMatches[2] || "",
			children: children,
			// Keep backward compatibility with simple array
			childrenIds: children.map((c) => c.id),
		};
	}

	return null;
}

/**
 * Formats a child for the graph ID string
 * @param {string|Object} child - Child ID string or object with {id, layout}
 * @returns {string} - Formatted child string (e.g., "B1" or "B1:RL")
 */
function formatChildForGraphId(child) {
	if (typeof child === "string") {
		return child;
	}
	if (child && typeof child === "object") {
		return child.layout ? `${child.id}:${child.layout}` : child.id;
	}
	return "";
}

/**
 * Generates a graph ID from components with layout information
 * @param {string} parent - Parent ID
 * @param {string} layout - Layout type ("LR", "TD", "RL", or "DT")
 * @param {string} current - Current ID
 * @param {Array} children - Array of child IDs (strings or objects with {id, layout})
 * @returns {string} - Generated graph ID
 */
function generateGraphId(parent, layout, current, children = []) {
	// Handle legacy calls with 3 parameters (parent, current, children)
	let actualParent = parent || "";
	let actualLayout = layout;
	let actualCurrent = current;
	let actualChildren = children;

	if (Array.isArray(layout)) {
		actualChildren = layout;
		actualCurrent = current || "";
		actualLayout = ""; // No layout for legacy calls (they're typically parents)
		actualParent = "";
		actualCurrent = parent;
	} else if (
		typeof layout === "string" &&
		layout !== "" &&
		!["LR", "TD", "RL", "DT"].includes(layout)
	) {
		// Handle legacy calls: generateGraphId(parent, current, children)
		actualChildren = current || [];
		actualCurrent = layout;
		actualLayout = ""; // No layout for legacy calls
	}

	// Only add layout annotation to children (when parent is not empty)
	const finalLayout = actualParent ? actualLayout : "";

	// Format children with their layouts
	const childrenStr = Array.isArray(actualChildren)
		? actualChildren
				.map(formatChildForGraphId)
				.filter((c) => c)
				.join(",")
		: "";

	// Handle different graph ID formats based on context
	// For root shapes (empty parent and empty layout), use legacy format without parentheses
	if (actualParent === "" && actualLayout === "") {
		return `graph[][${actualCurrent}][${childrenStr}]`;
	} else {
		// Non-root shape or shape with explicit layout: include layout parentheses
		return `graph[${actualParent}](${actualLayout})[${actualCurrent}][${childrenStr}]`;
	}
}

/**
 * Legacy wrapper for generateGraphId to maintain backwards compatibility
 * @param {string} parent - Parent ID
 * @param {string} current - Current ID
 * @param {Array} children - Array of child IDs
 * @returns {string} - Generated graph ID with LR layout
 */
function generateGraphIdLegacy(parent, current, children = []) {
	return generateGraphId(parent, "LR", current, children);
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
 * Sets the graph ID on a shape by updating its title (alt text)
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to update
 * @param {string} graphId - New graph ID to set
 */
function setShapeGraphId(shape, graphId) {
	try {
		// Set the shape's title (alt text) to the graph ID
		// This keeps the graph ID hidden while allowing custom text in the shape
		shape.setTitle(graphId);
	} catch (e) {
		console.log(`Warning: Could not set shape title: ${e.message}`);
	}
}

/**
 * Gets the graph ID from a shape's title (alt text)
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to read from
 * @returns {string|null} - Graph ID or null if not found
 */
function getShapeGraphId(shape) {
	try {
		// Get the shape's title (alt text)
		const title = shape.getTitle();
		if (title && title.startsWith("graph[")) {
			return title;
		}
		// Fallback: check text content for backward compatibility
		// This helps migrate old shapes that stored graph ID in text
		const text = shape.getText().asString().trim();
		if (text && text.startsWith("graph[")) {
			// Migrate to title and clear the text
			shape.setTitle(text);
			shape.getText().clear();
			return text;
		}
		return null;
	} catch (e) {
		console.log(`Warning: Could not get shape title: ${e.message}`);
		return null;
	}
}

/**
 * Updates parent shape to include new children in its ID
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape to update
 * @param {Array} newChildIds - Array of new child IDs (strings or objects with {id, layout})
 */
function updateParentWithChildren(parentShape, newChildIds) {
	const currentId = getShapeGraphId(parentShape);
	if (!currentId) {
		// If parent doesn't have a graph ID, create one without layout annotation
		const newId = generateGraphId("", "", "A1", newChildIds);
		setShapeGraphId(parentShape, newId);
		return;
	}

	const parsed = parseGraphId(currentId);
	if (!parsed) return;

	// Convert new children to objects if they're strings
	const newChildren = newChildIds.map((child) =>
		typeof child === "string" ? { id: child, layout: null } : child,
	);

	// Create a map of existing children to preserve their layouts
	const childMap = new Map();
	parsed.children.forEach((child) => {
		childMap.set(child.id, child);
	});

	// Add new children to the map
	newChildren.forEach((child) => {
		if (!childMap.has(child.id)) {
			childMap.set(child.id, child);
		}
	});

	// Convert map back to array
	const allChildren = Array.from(childMap.values());

	const updatedId = generateGraphId(
		parsed.parent,
		parsed.layout,
		parsed.current,
		allChildren,
	);
	setShapeGraphId(parentShape, updatedId);
}

/**
 * Initializes a shape as a root graph node
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to initialize
 * @returns {string} - Generated root graph ID
 */
function initializeAsRootGraphShape(shape) {
	const rootId = generateGraphId("", "", "A1", []);
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
	if (!parentId.childrenIds.includes(childId.current)) {
		const updatedChildren = [...parentId.childrenIds, childId.current];
		const newParentId = generateGraphId(
			parentId.parent,
			parentId.layout,
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent using parent's layout
	const newChildId = generateGraphId(
		parentId.current,
		parentId.layout,
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
