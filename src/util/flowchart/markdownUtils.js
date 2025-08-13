/**
 * Markdown to flowchart utilities
 * Converts markdown hierarchy to flowchart shapes using graph ID system
 */

/**
 * Parses markdown text into hierarchical structure
 * @param {string} markdownText - The markdown text to parse
 * @returns {Array} Array of parsed markdown items with level, text, and hierarchy info
 */
function parseMarkdownHierarchy(markdownText) {
	if (!markdownText || !markdownText.trim()) {
		return [];
	}

	const lines = markdownText.split("\n");
	const items = [];
	const levelStack = []; // Track parent hierarchy

	for (const line of lines) {
		const trimmedLine = line.trim();
		if (!trimmedLine || !trimmedLine.startsWith("#")) {
			continue; // Skip empty lines and non-header lines
		}

		// Count heading level (number of # symbols)
		const levelMatch = trimmedLine.match(/^(#+)\s+(.+)$/);
		if (!levelMatch) {
			continue;
		}

		const level = levelMatch[1].length; // Number of # symbols
		const text = levelMatch[2].trim();

		// Convert level to alphabet (1=A, 2=B, etc.)
		const alphabetLevel = String.fromCharCode(64 + level); // 65 is 'A'

		// Adjust level stack to current level
		while (levelStack.length >= level) {
			levelStack.pop();
		}

		// Generate current ID
		let currentId;
		if (levelStack.length === 0) {
			// Root level - just use alphabet + number
			const rootCount = items.filter((item) => item.level === 1).length + 1;
			currentId = alphabetLevel + rootCount;
		} else {
			// Child level - append to parent
			const parentId = levelStack[levelStack.length - 1];
			const siblingCount =
				items.filter(
					(item) =>
						item.level === level &&
						item.parentHierarchy.length > 0 &&
						item.parentHierarchy[item.parentHierarchy.length - 1] === parentId,
				).length + 1;
			currentId = alphabetLevel + siblingCount;
		}

		// Build parent hierarchy
		const parentHierarchy = [...levelStack];

		const item = {
			level: level,
			text: text,
			alphabetLevel: alphabetLevel,
			currentId: currentId,
			parentHierarchy: parentHierarchy,
			parentId:
				parentHierarchy.length > 0
					? parentHierarchy[parentHierarchy.length - 1]
					: null,
		};

		items.push(item);
		levelStack.push(currentId);
	}

	return items;
}

/**
 * Generates preview text showing the structure
 * @param {Array} items - Parsed markdown items
 * @returns {string} Preview text showing hierarchy
 */
function generateMarkdownPreview(items) {
	if (!items || items.length === 0) {
		return "No markdown to preview";
	}

	let preview = "";
	for (const item of items) {
		const indent = "  ".repeat(item.level - 1);
		const parentInfo = item.parentId
			? ` (parent: ${item.parentId})`
			: " (root)";
		preview += `${indent}${item.currentId}: ${item.text}${parentInfo}\n`;
	}

	return preview.trim();
}

/**
 * Creates flowchart from markdown hierarchy using existing child creation functions
 * @param {string} markdownText - The markdown text
 * @param {string} layout - Layout type ('TD' or 'LR')
 */
function createFromMarkdown(markdownText, layout) {
	try {
		markdownText = markdownText || "";
		layout = layout || "LR";

		if (!markdownText || !markdownText.trim()) {
			console.error("No markdown text provided");
			return;
		}

		// Parse markdown hierarchy
		const items = parseMarkdownHierarchy(markdownText);
		if (items.length === 0) {
			console.error("No valid markdown headers found");
			return;
		}

		// Get current slide and create root shape first
		const pres = SlidesApp.getActivePresentation();
		const slide = pres.getSelection().getCurrentPage();

		if (!slide) {
			console.error("No slide selected");
			return;
		}

		// Create the flowchart using existing child creation functions
		createMarkdownFlowchart(items, slide, layout);

		console.log(`Created flowchart with ${items.length} shapes from markdown`);
	} catch (e) {
		console.error(`Error creating flowchart from markdown: ${e.message}`);
	}
}

/**
 * Creates markdown flowchart using existing child creation functions
 * @param {Array} items - Parsed markdown items
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to create shapes on
 * @param {string} layout - Layout type ('TD' or 'LR')
 */
function createMarkdownFlowchart(items, slide, layout) {
	// Group items by level
	const itemsByLevel = new Map();
	for (const item of items) {
		if (!itemsByLevel.has(item.level)) {
			itemsByLevel.set(item.level, []);
		}
		itemsByLevel.get(item.level).push(item);
	}

	// Create root shapes (level 1) first
	const rootItems = itemsByLevel.get(1) || [];
	const createdShapes = new Map(); // currentId -> shape
	const shapeWidth = 60;
	const shapeHeight = 20;

	// Create all root shapes first
	for (let i = 0; i < rootItems.length; i++) {
		const item = rootItems[i];
		const x = 50 + i * (shapeWidth + 40);
		const y = 50;

		const rootShape = slide.insertShape(
			SlidesApp.ShapeType.RECTANGLE,
			x,
			y,
			shapeWidth,
			shapeHeight,
		);

		// Set text and style
		rootShape.getText().setText(item.text);

		// Apply default style if available
		const defaultStyle = getDefaultStyle();
		if (defaultStyle) {
			applyStyleToShape(rootShape, defaultStyle);
		}

		// Initialize as root shape with graph ID
		initializeAsRootGraphShape(rootShape);
		createdShapes.set(item.currentId, rootShape);
	}

	// Create children level by level using existing functions
	for (
		let level = 2;
		level <= Math.max(...items.map((item) => item.level));
		level++
	) {
		const levelItems = itemsByLevel.get(level) || [];

		for (const item of levelItems) {
			const parentShape = createdShapes.get(item.parentId);
			if (!parentShape) continue;

			// Select the parent shape temporarily
			parentShape.select();

			// Determine direction based on layout
			const direction = layout === "LR" ? "RIGHT" : "BOTTOM";

			// Create child using existing function with text
			const childShapes = createChildrenInDirectionWithText(
				direction,
				20, // gap
				"STRAIGHT", // lineType
				1, // count
				"NONE", // startArrow
				"FILL_ARROW", // endArrow
				[item.text], // texts array
				shapeWidth, // customWidth
				shapeHeight, // customHeight
				false, // maxWidth
				false, // maxHeight
				getDefaultStyle(), // defaultStyle
			);

			if (childShapes && childShapes.length > 0) {
				createdShapes.set(item.currentId, childShapes[0]);
			}
		}
	}

	// Clear selection at the end
	const selection = SlidesApp.getActivePresentation().getSelection();
	selection.unselectAll();
}

/**
 * Updates parent shapes' graph IDs with their children
 * @param {Map} createdShapes - Map of currentId -> shape
 * @param {Array} items - Parsed markdown items
 */
function updateParentChildrenInGraphIds(createdShapes, items) {
	// Build parent-children mapping
	const parentChildren = new Map();

	for (const item of items) {
		if (item.parentId) {
			if (!parentChildren.has(item.parentId)) {
				parentChildren.set(item.parentId, []);
			}
			parentChildren.get(item.parentId).push(item.currentId);
		}
	}

	// Update parent graph IDs
	for (const [parentId, childIds] of parentChildren.entries()) {
		const parentShape = createdShapes.get(parentId);
		if (parentShape) {
			const currentGraphId = getShapeGraphId(parentShape);
			if (currentGraphId) {
				const parsed = parseGraphId(currentGraphId);
				if (parsed) {
					const updatedGraphId = generateGraphId(
						parsed.parent,
						parsed.layout,
						parsed.current,
						childIds.join(","),
					);
					setShapeGraphId(parentShape, updatedGraphId);
				}
			}
		}
	}
}

/**
 * Gets default style from localStorage
 * @returns {Object|null} Default style object or null if not found
 */
function getDefaultStyle() {
	try {
		const styleData = localStorage.getItem("flowchart_default_style");
		if (styleData) {
			return JSON.parse(styleData);
		}
	} catch (e) {
		console.warn("Failed to parse default style from localStorage:", e);
	}
	return null;
}

/**
 * Applies style to a shape
 * @param {GoogleAppsScript.Slides.Shape} shape - The shape to style
 * @param {Object} style - Style object with properties
 */
function applyStyleToShape(shape, style) {
	try {
		if (style.fillColor) {
			shape.getFill().setSolidFill(style.fillColor);
		}
		if (style.borderColor) {
			shape.getBorder().getLineFill().setSolidFill(style.borderColor);
		}
		if (style.borderWeight) {
			shape.getBorder().setWeight(style.borderWeight);
		}
		if (style.fontFamily) {
			shape.getText().getTextStyle().setFontFamily(style.fontFamily);
		}
		if (style.fontSize) {
			shape.getText().getTextStyle().setFontSize(style.fontSize);
		}
		if (style.fontColor) {
			shape.getText().getTextStyle().setForegroundColor(style.fontColor);
		}
	} catch (e) {
		console.warn("Failed to apply style to shape:", e);
	}
}
