/**
 * Layout detection utilities for flowchart arrangements
 * Detects whether existing shapes are arranged horizontally (LR) or vertically (TD)
 */

/**
 * Analyzes connections to determine the primary layout direction
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape to analyze
 * @param {Array} siblingShapes - Array of sibling shape data
 * @returns {string} - "horizontal" or "vertical"
 */
function detectLayoutFromConnections(parentShape, siblingShapes) {
	const slide = parentShape.getParentPage();
	const allLines = slide.getLines();

	const parentCenter = getCenterOf(parentShape);
	const connections = { horizontal: 0, vertical: 0 };

	// Analyze all lines connected to parent or siblings
	for (const line of allLines) {
		try {
			const startSite = line.getLineConnections().getStartConnection();
			const endSite = line.getLineConnections().getEndConnection();

			if (!startSite || !endSite) continue;

			// Check if line connects parent to any sibling
			const startShape = startSite.getConnectedPage()
				? startSite.getConnectedPage()
				: null;
			const endShape = endSite.getConnectedPage()
				? endSite.getConnectedPage()
				: null;

			const isParentConnection =
				startShape === parentShape || endShape === parentShape;

			if (isParentConnection) {
				const otherShape = startShape === parentShape ? endShape : startShape;
				if (otherShape) {
					const otherCenter = getCenterOf(otherShape);
					const dx = Math.abs(otherCenter.x - parentCenter.x);
					const dy = Math.abs(otherCenter.y - parentCenter.y);

					// Determine primary direction of this connection
					if (dx > dy) {
						connections.horizontal++;
					} else {
						connections.vertical++;
					}
				}
			}
		} catch (e) {
			// Skip problematic connections
			continue;
		}
	}

	// Return the predominant direction
	return connections.horizontal >= connections.vertical
		? "horizontal"
		: "vertical";
}

/**
 * Analyzes sibling positions to detect layout pattern
 * @param {Array} siblingShapes - Array of sibling shape data
 * @param {number} tolerance - Position tolerance in pixels
 * @returns {string} - "horizontal" or "vertical"
 */
function detectLayoutFromPositions(siblingShapes, tolerance = 20) {
	if (siblingShapes.length < 2) {
		return "horizontal"; // Default fallback
	}

	const firstSibling = siblingShapes[0];
	let horizontalAligned = 0;
	let verticalAligned = 0;

	// Compare each sibling with the first one
	for (let i = 1; i < siblingShapes.length; i++) {
		const sibling = siblingShapes[i];
		const deltaX = Math.abs(sibling.left - firstSibling.left);
		const deltaY = Math.abs(sibling.top - firstSibling.top);

		if (deltaY < tolerance) horizontalAligned++;
		if (deltaX < tolerance) verticalAligned++;
	}

	return horizontalAligned > verticalAligned ? "horizontal" : "vertical";
}

/**
 * Enhanced layout detection combining multiple methods
 * @param {GoogleAppsScript.Slides.Shape} parentShape - Parent shape
 * @param {Array} siblingShapes - Array of sibling shapes
 * @returns {string} - "horizontal" or "vertical"
 */
function detectLayout(parentShape, siblingShapes) {
	// Method 1: Try connection-based detection first
	try {
		const connectionLayout = detectLayoutFromConnections(
			parentShape,
			siblingShapes,
		);
		if (connectionLayout) {
			return connectionLayout;
		}
	} catch (e) {
		console.log(`Connection detection failed: ${e.message}`);
	}

	// Method 2: Fallback to position-based detection
	try {
		return detectLayoutFromPositions(siblingShapes);
	} catch (e) {
		console.log(`Position detection failed: ${e.message}`);
	}

	// Method 3: Ultimate fallback
	return "horizontal";
}
