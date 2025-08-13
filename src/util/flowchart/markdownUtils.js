/**
 * Markdown to flowchart utilities
 * Converts markdown hierarchy to flowchart shapes using graph ID system
 */

/**
 * Parses markdown text into hierarchical structure following markmap-style hierarchy
 * @param {string} markdownText - The markdown text to parse
 * @returns {Array} Array of parsed markdown items with level, text, and hierarchy info
 */
function parseMarkdownHierarchy(markdownText) {
	if (!markdownText || !markdownText.trim()) {
		return [];
	}

	const lines = markdownText.split("\n");
	const items = [];
	const levelStack = []; // Track parent IDs at each level

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

		// Adjust level stack to current level depth
		// Keep only parent levels (remove deeper levels)
		while (levelStack.length >= level) {
			levelStack.pop();
		}

		// Count siblings at this level under the same parent
		const parentId =
			levelStack.length > 0 ? levelStack[levelStack.length - 1] : null;
		const siblingCount =
			items.filter((item) => item.level === level && item.parentId === parentId)
				.length + 1;

		// Generate current ID
		const currentId = alphabetLevel + siblingCount;

		// Build complete parent hierarchy path
		const parentHierarchy = [...levelStack];

		const item = {
			level: level,
			text: text,
			alphabetLevel: alphabetLevel,
			currentId: currentId,
			parentHierarchy: parentHierarchy,
			parentId: parentId,
		};

		items.push(item);
		levelStack.push(currentId); // Add current ID to stack for potential children
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
 * Creates flowchart from markdown hierarchy using canvas-based intelligent sizing
 * @param {string} markdownText - The markdown text
 * @param {string} layout - Layout type ('TD' or 'LR')
 * @param {number} horizontalGap - Horizontal gap between shapes
 * @param {number} verticalGap - Vertical gap between shapes
 * @param {string} lineType - Line type for connections ('STRAIGHT', 'BENT', 'CURVED')
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function createFromMarkdown(
	markdownText,
	layout,
	horizontalGap,
	verticalGap,
	lineType,
	startArrow,
	endArrow,
) {
	try {
		markdownText = markdownText || "";
		layout = layout || "LR";
		horizontalGap = horizontalGap || 20;
		verticalGap = verticalGap || 20;
		lineType = lineType || "BENT";
		startArrow = startArrow || "NONE";
		endArrow = endArrow || "FILL_ARROW";

		if (!markdownText || !markdownText.trim()) {
			console.error("No markdown text provided");
			return;
		}

		// First, validate and get the canvas shape
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		const validation = validateParentElement(range);
		if (validation.error) {
			console.error(validation.error);
			return;
		}

		const canvasShape = validation.shape;
		const slide = canvasShape.getParentPage();

		// Parse markdown hierarchy
		const items = parseMarkdownHierarchy(markdownText);
		if (items.length === 0) {
			console.error("No valid markdown headers found");
			return;
		}

		// Calculate intelligent dimensions based on canvas and markdown structure
		const canvasProps = {
			left: canvasShape.getLeft(),
			top: canvasShape.getTop(),
			width: canvasShape.getWidth(),
			height: canvasShape.getHeight(),
		};

		const dimensions = calculateIntelligentDimensions(
			items,
			layout,
			canvasProps,
			horizontalGap,
			verticalGap,
		);

		// Remove the canvas shape since we're replacing it
		canvasShape.remove();

		// Create the flowchart using intelligent positioning
		createMarkdownFlowchartIntelligent(
			items,
			slide,
			layout,
			canvasProps,
			dimensions,
			horizontalGap,
			verticalGap,
			lineType,
			startArrow,
			endArrow,
		);

		console.log(`Created flowchart with ${items.length} shapes from markdown`);
	} catch (e) {
		console.error(`Error creating flowchart from markdown: ${e.message}`);
	}
}

/**
 * Calculates intelligent dimensions based on canvas size and markdown structure
 * @param {Array} items - Parsed markdown items
 * @param {string} layout - Layout type ('TD' or 'LR')
 * @param {Object} canvasProps - Canvas properties {left, top, width, height}
 * @param {number} horizontalGap - Horizontal gap between shapes
 * @param {number} verticalGap - Vertical gap between shapes
 * @returns {Object} Dimensions object with shapeWidth, shapeHeight, and positioning info
 */
function calculateIntelligentDimensions(
	items,
	layout,
	canvasProps,
	horizontalGap,
	verticalGap,
) {
	// Analyze the markdown structure
	const maxLevel = Math.max(...items.map((item) => item.level));
	const itemsByLevel = new Map();
	let maxShapesInLevel = 0;

	for (const item of items) {
		if (!itemsByLevel.has(item.level)) {
			itemsByLevel.set(item.level, []);
		}
		itemsByLevel.get(item.level).push(item);
		maxShapesInLevel = Math.max(
			maxShapesInLevel,
			itemsByLevel.get(item.level).length,
		);
	}

	let shapeWidth, shapeHeight;

	if (layout === "LR") {
		// LR: levels spread horizontally, max shapes in a level spread vertically
		// Width: canvas width divided by number of levels (minus gaps)
		shapeWidth = Math.floor(
			(canvasProps.width - (maxLevel - 1) * horizontalGap) / maxLevel,
		);

		// Height: canvas height divided by max shapes in a level (minus gaps)
		shapeHeight = Math.floor(
			(canvasProps.height - (maxShapesInLevel - 1) * verticalGap) /
				maxShapesInLevel,
		);
	} else {
		// TD: levels spread vertically, max shapes in a level spread horizontally
		// Width: canvas width divided by max shapes in a level (minus gaps)
		shapeWidth = Math.floor(
			(canvasProps.width - (maxShapesInLevel - 1) * horizontalGap) /
				maxShapesInLevel,
		);

		// Height: canvas height divided by number of levels (minus gaps)
		shapeHeight = Math.floor(
			(canvasProps.height - (maxLevel - 1) * verticalGap) / maxLevel,
		);
	}

	// Ensure minimum dimensions
	shapeWidth = Math.max(shapeWidth, 40);
	shapeHeight = Math.max(shapeHeight, 20);

	return {
		shapeWidth,
		shapeHeight,
		maxLevel,
		maxShapesInLevel,
		itemsByLevel,
	};
}

/**
 * Creates markdown flowchart with intelligent positioning within canvas bounds
 * @param {Array} items - Parsed markdown items
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to create shapes on
 * @param {string} layout - Layout type ('TD' or 'LR')
 * @param {Object} canvasProps - Canvas properties {left, top, width, height}
 * @param {Object} dimensions - Calculated dimensions
 * @param {number} horizontalGap - Horizontal gap between shapes
 * @param {number} verticalGap - Vertical gap between shapes
 * @param {string} lineType - Line type for connections ('STRAIGHT', 'BENT', 'CURVED')
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function createMarkdownFlowchartIntelligent(
	items,
	slide,
	layout,
	canvasProps,
	dimensions,
	horizontalGap,
	verticalGap,
	lineType,
	startArrow,
	endArrow,
) {
	const { shapeWidth, shapeHeight, itemsByLevel } = dimensions;
	const createdShapes = new Map(); // currentId -> shape
	const defaultStyle = getDefaultStyle();

	// Create all shapes with calculated positions
	for (const [level, levelItems] of itemsByLevel.entries()) {
		for (let i = 0; i < levelItems.length; i++) {
			const item = levelItems[i];

			// Calculate position based on layout
			let x, y;
			if (layout === "LR") {
				// LR: level determines X, index in level determines Y
				x = canvasProps.left + (level - 1) * (shapeWidth + horizontalGap);

				// Center the level vertically within canvas
				const levelHeight =
					levelItems.length * shapeHeight +
					(levelItems.length - 1) * verticalGap;
				const levelStartY =
					canvasProps.top + (canvasProps.height - levelHeight) / 2;
				y = levelStartY + i * (shapeHeight + verticalGap);
			} else {
				// TD: level determines Y, index in level determines X
				y = canvasProps.top + (level - 1) * (shapeHeight + verticalGap);

				// Center the level horizontally within canvas
				const levelWidth =
					levelItems.length * shapeWidth +
					(levelItems.length - 1) * horizontalGap;
				const levelStartX =
					canvasProps.left + (canvasProps.width - levelWidth) / 2;
				x = levelStartX + i * (shapeWidth + horizontalGap);
			}

			// Create shape with intelligent dimensions
			const shape = slide.insertShape(
				SlidesApp.ShapeType.RECTANGLE,
				x,
				y,
				shapeWidth,
				shapeHeight,
			);

			// Set text and style
			shape.getText().setText(item.text);

			if (defaultStyle) {
				applyStyleToShape(shape, defaultStyle);
			}

			// Set graph ID with proper parent hierarchy
			const parentHierarchyString =
				item.parentHierarchy.length > 0 ? item.parentHierarchy.join("|") : "";
			const graphId = generateGraphId(
				parentHierarchyString,
				layout,
				item.currentId,
				"", // Children will be added later in updateParentChildrenInGraphIds
			);
			setShapeGraphId(shape, graphId);

			createdShapes.set(item.currentId, shape);
		}
	}

	// Create connections between parent and child shapes
	createMarkdownConnections(
		items,
		createdShapes,
		layout,
		lineType,
		startArrow,
		endArrow,
	);

	// Update parent graph IDs with children
	updateParentChildrenInGraphIds(createdShapes, items);
}

/**
 * Updates parent shapes' graph IDs with their children information
 * @param {Map} createdShapes - Map of currentId -> shape
 * @param {Array} items - Parsed markdown items
 */
function updateParentChildrenInGraphIds(createdShapes, items) {
	// Build parent-children mapping with layout information
	const parentChildren = new Map();

	for (const item of items) {
		if (item.parentId) {
			if (!parentChildren.has(item.parentId)) {
				parentChildren.set(item.parentId, []);
			}
			// Store child with layout info (all children in markdown have same layout)
			parentChildren.get(item.parentId).push({
				id: item.currentId,
				layout: "LR", // Default layout, will be updated based on actual layout parameter
			});
		}
	}

	// Update parent graph IDs with children information
	for (const [parentId, children] of parentChildren.entries()) {
		const parentShape = createdShapes.get(parentId);
		if (parentShape) {
			const currentGraphId = getShapeGraphId(parentShape);
			if (currentGraphId) {
				const parsed = parseGraphId(currentGraphId);
				if (parsed) {
					// Update children with the correct layout from the parent's graph ID
					const childrenWithLayout = children.map((child) => ({
						id: child.id,
						layout: parsed.layout || "LR",
					}));

					const updatedGraphId = generateGraphId(
						parsed.parent,
						parsed.layout,
						parsed.current,
						childrenWithLayout,
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
 * Creates connections between parent and child shapes based on markdown hierarchy
 * @param {Array} items - Parsed markdown items with hierarchy information
 * @param {Map} createdShapes - Map of currentId -> shape
 * @param {string} layout - Layout type ('TD' or 'LR')
 * @param {string} lineType - Line type for connections ('STRAIGHT', 'BENT', 'CURVED')
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function createMarkdownConnections(
	items,
	createdShapes,
	layout,
	lineType,
	startArrow,
	endArrow,
) {
	for (const item of items) {
		if (item.parentId && createdShapes.has(item.parentId)) {
			const parentShape = createdShapes.get(item.parentId);
			const childShape = createdShapes.get(item.currentId);

			if (parentShape && childShape) {
				// Determine connection sides based on layout
				let parentSide, childSide;
				if (layout === "LR") {
					parentSide = "RIGHT";
					childSide = "LEFT";
				} else {
					parentSide = "BOTTOM";
					childSide = "TOP";
				}

				// Create connection using existing utilities
				const parentSite = pickConnectionSite(parentShape, parentSide);
				const childSite = pickConnectionSite(childShape, childSide);

				if (parentSite && childSite) {
					const slide = parentShape.getParentPage();
					const lineCategory =
						SlidesApp.LineCategory[lineType] || SlidesApp.LineCategory.BENT;
					const line = slide.insertLine(lineCategory, parentSite, childSite);

					// Apply arrow styles from line settings
					if (
						startArrow &&
						startArrow !== "NONE" &&
						SlidesApp.ArrowStyle[startArrow]
					) {
						line.setStartArrow(SlidesApp.ArrowStyle[startArrow]);
					}
					if (
						endArrow &&
						endArrow !== "NONE" &&
						SlidesApp.ArrowStyle[endArrow]
					) {
						line.setEndArrow(SlidesApp.ArrowStyle[endArrow]);
					}
				}
			}
		}
	}
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
