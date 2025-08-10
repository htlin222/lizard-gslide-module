/**
 * Background utilities for flowchart elements
 * Provides functions for creating background rectangles and styling
 */

/**
 * Adds a background rectangle to encompass all selected elements
 * Creates a rectangle with specified padding, color, and opacity behind the selected shapes
 * @param {number} padding - Padding around the selection in points (default 10)
 * @param {string} bgColor - Background color in hex format (default #f3f3f3)
 * @param {number} opacity - Opacity from 0.0 to 1.0 (default 0.5)
 */
function addBackgroundToSelectedElements(
	padding = 10,
	bgColor = "#f3f3f3",
	opacity = 0.5,
) {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			throw new Error("Please select one or more shapes to add background.");
		}

		const elements = range.getPageElements();
		if (elements.length === 0) {
			throw new Error("No elements selected.");
		}

		// Filter to only include shapes
		const shapes = elements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (shapes.length === 0) {
			throw new Error("Please select at least one shape to add background.");
		}

		// Calculate bounding box of all selected shapes
		const boundingBox = calculateShapesBoundingBox(shapes);

		// Calculate background rectangle dimensions with padding
		const bgLeft = boundingBox.minLeft - padding;
		const bgTop = boundingBox.minTop - padding;
		const bgWidth = boundingBox.width + padding * 2;
		const bgHeight = boundingBox.height + padding * 2;

		// Get the slide to add the background rectangle
		const slide = shapes[0].asShape().getParentPage();

		// Create and style the background rectangle
		const bgRect = createBackgroundRectangle(
			slide,
			bgLeft,
			bgTop,
			bgWidth,
			bgHeight,
			bgColor,
			opacity,
		);

		console.log(
			`Background rectangle created: ${bgWidth}x${bgHeight} at (${bgLeft}, ${bgTop}) with color ${bgColor} and opacity ${opacity}`,
		);

		return `Background rectangle added successfully! Size: ${Math.round(bgWidth)}×${Math.round(bgHeight)} with ${Math.round(opacity * 100)}% opacity.`;
	} catch (e) {
		const errorMsg = `Error creating background: ${e.message}`;
		console.error(errorMsg);
		throw new Error(errorMsg);
	}
}

/**
 * Calculates the bounding box for an array of shapes
 * @param {Array} shapes - Array of shape elements
 * @returns {Object} Bounding box with minLeft, minTop, maxRight, maxBottom, width, height
 */
function calculateShapesBoundingBox(shapes) {
	let minLeft = Number.MAX_VALUE;
	let minTop = Number.MAX_VALUE;
	let maxRight = Number.MIN_VALUE;
	let maxBottom = Number.MIN_VALUE;

	shapes.forEach((element) => {
		const shape = element.asShape();
		const left = shape.getLeft();
		const top = shape.getTop();
		const width = shape.getWidth();
		const height = shape.getHeight();

		minLeft = Math.min(minLeft, left);
		minTop = Math.min(minTop, top);
		maxRight = Math.max(maxRight, left + width);
		maxBottom = Math.max(maxBottom, top + height);
	});

	return {
		minLeft,
		minTop,
		maxRight,
		maxBottom,
		width: maxRight - minLeft,
		height: maxBottom - minTop,
	};
}

/**
 * Creates and styles a background rectangle
 * @param {Slide} slide - The slide to add the rectangle to
 * @param {number} left - Left position
 * @param {number} top - Top position
 * @param {number} width - Rectangle width
 * @param {number} height - Rectangle height
 * @param {string} bgColor - Background color in hex format
 * @param {number} opacity - Opacity from 0.0 to 1.0
 * @returns {Shape} The created background rectangle
 */
function createBackgroundRectangle(
	slide,
	left,
	top,
	width,
	height,
	bgColor,
	opacity,
) {
	// Create background rectangle
	const bgRect = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		left,
		top,
		width,
		height,
	);

	// Style the background rectangle
	bgRect.getFill().setSolidFill(bgColor, opacity);

	// Set white border as specified
	bgRect.getBorder().setWeight(1);
	bgRect.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Send the background rectangle to the back so it appears behind other shapes
	bgRect.sendToBack();

	return bgRect;
}

/**
 * Creates a background for a specific group of shapes with custom styling
 * @param {Array} shapes - Array of shape elements to encompass
 * @param {Object} style - Style configuration object
 * @param {number} style.padding - Padding around shapes (default 10)
 * @param {string} style.color - Background color (default #f3f3f3)
 * @param {number} style.opacity - Opacity 0-1 (default 0.5)
 * @param {number} style.borderWidth - Border width (default 1)
 * @param {string} style.borderColor - Border color (default #FFFFFF)
 * @returns {Shape} The created background rectangle
 */
function createCustomBackground(shapes, style = {}) {
	const config = {
		padding: style.padding || 10,
		color: style.color || "#f3f3f3",
		opacity: style.opacity || 0.5,
		borderWidth: style.borderWidth || 1,
		borderColor: style.borderColor || "#FFFFFF",
		...style,
	};

	if (!shapes || shapes.length === 0) {
		throw new Error("No shapes provided for background creation.");
	}

	// Calculate bounding box
	const boundingBox = calculateShapesBoundingBox(shapes);

	// Calculate background rectangle dimensions with padding
	const bgLeft = boundingBox.minLeft - config.padding;
	const bgTop = boundingBox.minTop - config.padding;
	const bgWidth = boundingBox.width + config.padding * 2;
	const bgHeight = boundingBox.height + config.padding * 2;

	// Get the slide from the first shape
	const slide = shapes[0].asShape().getParentPage();

	// Create background rectangle
	const bgRect = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		bgLeft,
		bgTop,
		bgWidth,
		bgHeight,
	);

	// Apply styling
	bgRect.getFill().setSolidFill(config.color, config.opacity);
	bgRect.getBorder().setWeight(config.borderWidth);
	bgRect.getBorder().getLineFill().setSolidFill(config.borderColor);

	// Send to back
	bgRect.sendToBack();

	return bgRect;
}
