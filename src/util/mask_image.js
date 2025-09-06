// Global alpha value for overlay transparency
const OVERLAY_ALPHA = 0.8;

/**
 * Masks an image by highlighting specific areas defined by multiple shapes.
 * Creates semi-transparent white rectangles using a grid-based approach to prevent overlaps.
 * Requires selecting multiple shapes and one image.
 * @return {boolean} True if successful, false otherwise
 */
function maskImage() {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const currentSlide = selection.getCurrentPage();

		// Get the selected elements
		const pageElementRange = selection.getPageElementRange();

		if (!pageElementRange) {
			SlidesApp.getUi().alert(
				"No selection found. Please select shapes and an image.",
			);
			return false;
		}

		const selectedElements = pageElementRange.getPageElements();

		if (!selectedElements || selectedElements.length < 2) {
			SlidesApp.getUi().alert(
				"Please select at least one shape and one image.",
			);
			return false;
		}

		// Find which elements are shapes and which is the image
		const shapeElements = [];
		let imageElement = null;

		for (const element of selectedElements) {
			if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
				shapeElements.push(element.asShape());
			} else if (
				element.getPageElementType() === SlidesApp.PageElementType.IMAGE
			) {
				if (imageElement !== null) {
					SlidesApp.getUi().alert("Please select only one image.");
					return false;
				}
				imageElement = element.asImage();
			}
		}

		if (shapeElements.length === 0 || !imageElement) {
			SlidesApp.getUi().alert(
				"Please select at least one shape and exactly one image.",
			);
			return false;
		}

		// Get image dimensions and position
		const imageLeft = imageElement.getLeft();
		const imageTop = imageElement.getTop();
		const imageWidth = imageElement.getWidth();
		const imageHeight = imageElement.getHeight();

		// Validate all shapes are at least partially within the image bounds
		for (let i = 0; i < shapeElements.length; i++) {
			const shape = shapeElements[i];
			const shapeLeft = shape.getLeft();
			const shapeTop = shape.getTop();
			const shapeWidth = shape.getWidth();
			const shapeHeight = shape.getHeight();

			if (
				shapeLeft > imageLeft + imageWidth ||
				shapeLeft + shapeWidth < imageLeft ||
				shapeTop > imageTop + imageHeight ||
				shapeTop + shapeHeight < imageTop
			) {
				SlidesApp.getUi().alert(
					`Shape ${i + 1} and the image don't overlap. Please adjust their positions.`,
				);
				return false;
			}
		}

		// Use grid-based masking approach to prevent overlaps
		const overlays = createGridBasedMask(
			currentSlide,
			imageElement,
			shapeElements,
		);

		// Create an array of all elements to group (overlays and original image, but not the shapes)
		const pageElements = [imageElement, ...overlays].filter(Boolean);

		// Group all elements together
		if (pageElements.length > 1) {
			currentSlide.group(pageElements);
		}

		// Delete all shapes as they're no longer needed
		for (const shape of shapeElements) {
			shape.remove();
		}

		return true;
	} catch (error) {
		SlidesApp.getUi().alert("Error: " + error.message);
		console.log("Error in maskImage: " + error.message);
		console.log(error.stack);
		return false;
	}
}

/**
 * Creates non-overlapping mask overlays for multiple shapes on an image.
 * Uses a grid-based approach to prevent duplicate overlapping masks.
 * @param {Object} currentSlide - The current slide
 * @param {Object} imageElement - The image element to mask
 * @param {Array} shapeElements - Array of shape elements to use as masks
 * @return {Array} Array of overlay shapes created
 */
function createGridBasedMask(currentSlide, imageElement, shapeElements) {
	const overlays = [];

	// Get image boundaries
	const imageLeft = imageElement.getLeft();
	const imageTop = imageElement.getTop();
	const imageWidth = imageElement.getWidth();
	const imageHeight = imageElement.getHeight();
	const imageRight = imageLeft + imageWidth;
	const imageBottom = imageTop + imageHeight;

	// Collect all unique X and Y coordinates to create a grid (force to integers)
	const xCoordinates = [Math.round(imageLeft), Math.round(imageRight)];
	const yCoordinates = [Math.round(imageTop), Math.round(imageBottom)];

	for (const shape of shapeElements) {
		const shapeLeft = Math.round(shape.getLeft());
		const shapeTop = Math.round(shape.getTop());
		const shapeRight = Math.round(shapeLeft + shape.getWidth());
		const shapeBottom = Math.round(shapeTop + shape.getHeight());

		// Add X coordinates within image bounds
		if (
			shapeLeft > Math.round(imageLeft) &&
			shapeLeft < Math.round(imageRight)
		) {
			xCoordinates.push(shapeLeft);
		}
		if (
			shapeRight > Math.round(imageLeft) &&
			shapeRight < Math.round(imageRight)
		) {
			xCoordinates.push(shapeRight);
		}

		// Add Y coordinates within image bounds
		if (shapeTop > Math.round(imageTop) && shapeTop < Math.round(imageBottom)) {
			yCoordinates.push(shapeTop);
		}
		if (
			shapeBottom > Math.round(imageTop) &&
			shapeBottom < Math.round(imageBottom)
		) {
			yCoordinates.push(shapeBottom);
		}
	}

	// Sort and remove duplicates
	const sortedXCoords = [...new Set(xCoordinates)].sort((a, b) => a - b);
	const sortedYCoords = [...new Set(yCoordinates)].sort((a, b) => a - b);

	// Process each grid cell
	for (let i = 0; i < sortedXCoords.length - 1; i++) {
		for (let j = 0; j < sortedYCoords.length - 1; j++) {
			const cellLeft = sortedXCoords[i];
			const cellRight = sortedXCoords[i + 1];
			const cellTop = sortedYCoords[j];
			const cellBottom = sortedYCoords[j + 1];
			const cellWidth = cellRight - cellLeft;
			const cellHeight = cellBottom - cellTop;

			// Check if any shape covers this cell
			let cellCoveredByShape = false;

			for (const shape of shapeElements) {
				const shapeLeft = Math.round(shape.getLeft());
				const shapeTop = Math.round(shape.getTop());
				const shapeRight = Math.round(shapeLeft + shape.getWidth());
				const shapeBottom = Math.round(shapeTop + shape.getHeight());

				// Check if shape completely covers this cell
				if (
					shapeLeft <= cellLeft &&
					shapeRight >= cellRight &&
					shapeTop <= cellTop &&
					shapeBottom >= cellBottom
				) {
					cellCoveredByShape = true;
					break;
				}
			}

			// If cell is not covered by any shape, create a mask
			if (!cellCoveredByShape) {
				const overlay = currentSlide.insertShape(
					SlidesApp.ShapeType.RECTANGLE,
					Math.round(cellLeft),
					Math.round(cellTop),
					Math.round(cellWidth),
					Math.round(cellHeight),
				);
				overlay.getFill().setSolidFill("#FFFFFF", OVERLAY_ALPHA);
				overlay.getBorder().setWeight(0.1);
				overlay.getBorder().getLineFill().setSolidFill("#FFFFFF", 0);
				overlays.push(overlay);
			}
		}
	}

	return overlays;
}
