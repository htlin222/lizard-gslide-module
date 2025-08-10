/**
 * Debug utilities for flowchart development and troubleshooting
 * Provides functions for inspecting and debugging flowchart elements
 */

/**
 * Shows the Graph ID of the currently selected shape
 * @returns {string} Graph ID information about the selected shape
 */
function showSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to show its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get Graph ID from title (alt text)
		const graphId = getShapeGraphId(shape);

		if (graphId) {
			return formatGraphIdInfo(graphId);
		} else {
			return "No Graph ID found. This shape may not be part of a flowchart.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

/**
 * Formats Graph ID information for display
 * @param {string} graphId - The raw graph ID string
 * @returns {string} Formatted graph ID information
 */
function formatGraphIdInfo(graphId) {
	// Parse the graph ID to show more details
	const parsed = parseGraphId(graphId);
	if (parsed) {
		const details = [];
		details.push(`📊 Graph ID:\n${graphId}`);
		details.push(`├─ Parent: ${parsed.parent || "(root)"}`);
		details.push(`├─ Layout: ${parsed.layout || "(none)"}`);
		details.push(`├─ Current: ${parsed.current}`);
		details.push(
			`└─ Children: ${parsed.children.length > 0 ? parsed.children.join(", ") : "(none)"}`,
		);
		return details.join("\n");
	}
	return `📊 Graph ID:\n${graphId}`;
}

/**
 * Clears the Graph ID from the currently selected shape
 * @returns {string} Confirmation message
 */
function clearSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to clear its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get current Graph ID for confirmation
		const currentGraphId = getShapeGraphId(shape);

		if (currentGraphId) {
			// Clear the title (alt text) completely
			// Try multiple methods to ensure it's properly cleared
			try {
				// Method 1: Set to null (preferred)
				shape.setTitle(null);
			} catch (e1) {
				try {
					// Method 2: Set to undefined
					shape.setTitle(undefined);
				} catch (e2) {
					// Method 3: Set to empty string as fallback
					shape.setTitle("");
				}
			}

			// Also clear text content if it contains a graph ID (backward compatibility)
			try {
				const text = shape.getText().asString().trim();
				if (text && text.startsWith("graph[")) {
					shape.getText().setText("");
				}
			} catch (textError) {
				console.log(`Note: Could not clear text content: ${textError.message}`);
			}

			// Verify the Graph ID was actually cleared
			const verifyCleared = getShapeGraphId(shape);
			if (verifyCleared) {
				return `⚠️ Graph ID partially cleared. Previous ID was: ${currentGraphId}\nNote: Some Graph ID data may still remain.`;
			} else {
				return `✅ Graph ID cleared successfully!\nPrevious ID was: ${currentGraphId}`;
			}
		} else {
			return "No Graph ID to clear. This shape doesn't have a Graph ID.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

/**
 * Identifies which shapes are connected to the selected line
 * This helps debug connection issues by showing what shapes would be connected
 * @returns {string} - Information about connected shapes
 */
function identifyConnectedShapes() {
	try {
		console.log("identifyConnectedShapes function called");
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		console.log("Got selection:", selection ? "exists" : "null");

		if (!selection) {
			return "Please select a line to identify its connected shapes.";
		}

		const selectionType = selection.getSelectionType();
		const elements = getElementsFromSelection(selection, selectionType);

		if (!elements) {
			return "Please select a line element.";
		}

		if (elements.length !== 1) {
			return "Please select exactly ONE line.";
		}

		const element = elements[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.LINE) {
			return "Selected element must be a LINE.";
		}

		return analyzeLineConnections(element.asLine());
	} catch (e) {
		return `Error: ${e.message}`;
	}
}

/**
 * Gets elements from selection based on selection type
 * @param {Selection} selection - The current selection
 * @param {SelectionType} selectionType - Type of selection
 * @returns {Array|null} Array of elements or null if invalid
 */
function getElementsFromSelection(selection, selectionType) {
	let elements = [];

	if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT_RANGE) {
		const range = selection.getPageElementRange();
		if (!range) {
			return null;
		}
		elements = range.getPageElements();
	} else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
		const range = selection.getPageElementRange();
		if (range) {
			elements = range.getPageElements();
		} else {
			return null;
		}
	} else {
		return null;
	}

	return elements;
}

/**
 * Analyzes line connections and returns detailed information
 * @param {Line} line - The line to analyze
 * @returns {string} Analysis results
 */
function analyzeLineConnections(line) {
	const slide = line.getParentPage();

	// Get line properties
	const lineLeft = line.getLeft();
	const lineTop = line.getTop();
	const lineWidth = line.getWidth();
	const lineHeight = line.getHeight();

	// Get all shapes on slide
	const allShapes = slide.getShapes();

	const results = [];
	results.push(`📊 Line Analysis:`);
	results.push(`Position: (${Math.round(lineLeft)}, ${Math.round(lineTop)})`);
	results.push(`Size: ${Math.round(lineWidth)} × ${Math.round(lineHeight)}`);
	results.push(`Total shapes on slide: ${allShapes.length}`);
	results.push(``);

	// Show first few shapes as candidates
	results.push(`🎯 Shape Candidates:`);
	for (let i = 0; i < Math.min(4, allShapes.length); i++) {
		const shape = allShapes[i];
		const shapeInfo = analyzeShapeDistanceFromLine(shape, line);
		results.push(
			`Shape ${i + 1}: (${shapeInfo.left}, ${shapeInfo.top}) ${shapeInfo.width}×${shapeInfo.height} - Distance: ${shapeInfo.distance}px`,
		);
	}

	results.push(``);
	results.push(`💡 Current logic will connect:`);
	if (allShapes.length >= 2) {
		results.push(
			`• Shape 1: (${Math.round(allShapes[0].getLeft())}, ${Math.round(allShapes[0].getTop())})`,
		);
		results.push(
			`• Shape 2: (${Math.round(allShapes[1].getLeft())}, ${Math.round(allShapes[1].getTop())})`,
		);
	} else {
		results.push(`• Not enough shapes on slide for connection analysis`);
	}

	return results.join("\n");
}

/**
 * Analyzes the distance between a shape and a line
 * @param {Shape} shape - The shape to analyze
 * @param {Line} line - The line to measure distance from
 * @returns {Object} Shape information including distance from line
 */
function analyzeShapeDistanceFromLine(shape, line) {
	const shapeLeft = Math.round(shape.getLeft());
	const shapeTop = Math.round(shape.getTop());
	const shapeWidth = Math.round(shape.getWidth());
	const shapeHeight = Math.round(shape.getHeight());

	// Calculate distance from line
	const lineCenter = {
		x: line.getLeft() + line.getWidth() / 2,
		y: line.getTop() + line.getHeight() / 2,
	};
	const shapeCenter = {
		x: shapeLeft + shapeWidth / 2,
		y: shapeTop + shapeHeight / 2,
	};
	const distance = Math.round(
		Math.sqrt(
			(lineCenter.x - shapeCenter.x) ** 2 + (lineCenter.y - shapeCenter.y) ** 2,
		),
	);

	return {
		left: shapeLeft,
		top: shapeTop,
		width: shapeWidth,
		height: shapeHeight,
		distance,
	};
}

/**
 * Initializes a selected shape as a root graph node
 * Useful for starting a new flowchart hierarchy
 */
function initializeRootGraphShape() {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert(
			"Please select a shape to initialize as root graph node.",
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

	const shape = element.asShape();
	// Initialize as root without layout annotation
	const rootId = generateGraphId("", "", "A1", []);
	setShapeGraphId(shape, rootId);

	SlidesApp.getUi().alert(
		"Root graph shape initialized",
		`Shape is now: ${rootId}`,
	);
}

// Keep the old function name for backward compatibility but redirect to new function
function debugShowTitlePlaceholders() {
	return showSelectedShapeGraphId();
}

/**
 * Analyzes the current slide and provides debugging information
 * @returns {string} Comprehensive slide analysis
 */
function analyzeCurrentSlide() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const slide = pres.getSelection().getCurrentPage();

		if (!slide) {
			return "No active slide found.";
		}

		const shapes = slide.getShapes();
		const lines = slide.getLines();
		const results = [];

		results.push(`📊 Slide Analysis:`);
		results.push(`Total shapes: ${shapes.length}`);
		results.push(`Total lines: ${lines.length}`);
		results.push(``);

		if (shapes.length > 0) {
			results.push(`🎯 Shapes with Graph IDs:`);
			let graphShapeCount = 0;
			shapes.forEach((shape, index) => {
				const graphId = getShapeGraphId(shape);
				if (graphId) {
					graphShapeCount++;
					results.push(`Shape ${index + 1}: ${graphId}`);
				}
			});
			results.push(`Total shapes with Graph IDs: ${graphShapeCount}`);
		}

		return results.join("\n");
	} catch (e) {
		return `Error analyzing slide: ${e.message}`;
	}
}
