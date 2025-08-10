/**
 * Stage Bar utilities for flowchart elements
 * Provides functions for creating HOME_PLATE stage bars for process flows and timelines
 * Creates visual process indicators behind selected shapes
 */

/**
 * Adds a stage bar (HOME_PLATE shape) behind selected shapes
 * @param {number} baseY - Base Y position for the stage bar (default 100)
 * @param {number} offsetX - X offset from the leftmost shape (default -20)
 * @param {number} extraWidth - Extra width beyond the rightmost shape (default 30)
 * @param {number} height - Height of the stage bar (default 15)
 * @param {string} fillColor - Fill color for the stage bar (default #3D6869)
 * @param {number} opacity - Opacity of the stage bar (default 1.0)
 * @param {string} strokeColor - Border stroke color (default #FFFFFF)
 * @param {string} text - Optional text to display on the stage bar (default "")
 * @param {number} fontSize - Font size for the text (default 10)
 */
function addStageBar(
	baseY = 100,
	offsetX = -20,
	extraWidth = 30,
	height = 15,
	fillColor = "#3D6869",
	opacity = 1.0,
	strokeColor = "#FFFFFF",
	text = "",
	fontSize = 10,
) {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();

		if (!selection) {
			throw new Error("Please select one or more shapes to add a stage bar.");
		}

		const range = selection.getPageElementRange();
		if (!range) {
			throw new Error("Please select one or more shapes to add a stage bar.");
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
			throw new Error("Please select at least one shape to add a stage bar.");
		}

		// Calculate the leftmost and rightmost positions
		let minLeft = Number.MAX_VALUE;
		let maxRight = Number.MIN_VALUE;

		shapes.forEach((element) => {
			const shape = element.asShape();
			const left = shape.getLeft();
			const width = shape.getWidth();
			const right = left + width;

			minLeft = Math.min(minLeft, left);
			maxRight = Math.max(maxRight, right);
		});

		// Calculate stage bar dimensions
		const stageLeft = minLeft + offsetX;
		const stageWidth = maxRight - minLeft + extraWidth;
		const stageTop = baseY;

		// Get the slide from the first shape
		const slide = shapes[0].asShape().getParentPage();

		// Create the HOME_PLATE shape as stage bar
		const stageBar = slide.insertShape(
			SlidesApp.ShapeType.HOME_PLATE,
			stageLeft,
			stageTop,
			stageWidth,
			height,
		);

		// Style the stage bar
		stageBar.getFill().setSolidFill(fillColor, opacity);

		// Set white border by default
		stageBar.getBorder().setWeight(1);
		stageBar.getBorder().getLineFill().setSolidFill(strokeColor);

		// Add text if provided
		if (text && text.trim() !== "") {
			stageBar.getText().setText(text);
			const textStyle = stageBar.getText().getTextStyle();
			textStyle.setFontSize(fontSize);
			textStyle.setForegroundColor("#FFFFFF"); // White text for contrast
			textStyle.setBold(false);

			// Center align the text
			const paragraphStyle = stageBar.getText().getParagraphStyle();
			paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
		}

		// Set title for identification
		stageBar.setTitle("flowchartstagesbar");

		// Reorder all stage bars on the slide (leftmost on top)
		// This will handle sending all bars to back in the correct order
		reorderStageBars(slide);

		// Select the created stage bar
		stageBar.select();

		console.log(
			`Stage bar created: ${stageWidth}x${height} at (${stageLeft}, ${stageTop}) with color ${fillColor}`,
		);

		return `Stage bar added successfully! Width: ${Math.round(stageWidth)} at Y: ${baseY}`;
	} catch (e) {
		const errorMsg = `Error creating stage bar: ${e.message}`;
		console.error(errorMsg);
		throw new Error(errorMsg);
	}
}

/**
 * Quick function to add a default stage bar
 * Uses default settings for rapid prototyping
 */
function addDefaultStageBar() {
	return addStageBar();
}

/**
 * Add a stage bar with custom color theme
 * @param {string} theme - Color theme: 'blue', 'green', 'orange', 'purple', or hex color
 * @param {string} text - Optional text to display on the stage bar
 */
function addThemedStageBar(theme = "blue", text = "") {
	const themes = {
		blue: "#2196F3",
		green: "#4CAF50",
		orange: "#FF9800",
		purple: "#9C27B0",
		teal: "#009688",
		red: "#F44336",
		gray: "#757575",
	};

	const color = themes[theme] || theme;
	return addStageBar(100, -20, 30, 15, color, 1.0, "#FFFFFF", text, 10);
}

/**
 * Add multiple stage bars at different Y positions
 * Useful for creating multi-level process diagrams
 * @param {Array} yPositions - Array of Y positions for multiple stage bars
 * @param {string} fillColor - Fill color for all stage bars
 */
function addMultipleStageBar(
	yPositions = [100, 150, 200],
	fillColor = "#3D6869",
) {
	const results = [];

	yPositions.forEach((yPos, index) => {
		try {
			const result = addStageBar(
				yPos,
				-20,
				30,
				15,
				fillColor,
				1.0 - index * 0.1,
				"#FFFFFF",
			);
			results.push(`Level ${index + 1}: ${result}`);
		} catch (e) {
			results.push(`Level ${index + 1}: Failed - ${e.message}`);
		}
	});

	return results.join("\n");
}

/**
 * Reorders all stage bars on a slide so that leftmost bars appear on top
 * This creates a proper layering effect for overlapping stage bars
 * @param {Slide} slide - The slide containing the stage bars
 */
function reorderStageBars(slide) {
	try {
		// Get all shapes on the slide
		const shapes = slide.getShapes();

		// Filter to only get stage bars (shapes with title "flowchartstagesbar")
		const stageBars = [];
		shapes.forEach((shape) => {
			if (shape.getTitle() === "flowchartstagesbar") {
				stageBars.push({
					shape: shape,
					x: shape.getLeft(),
				});
			}
		});

		// Always send all stage bars to back, even if there's only one
		// This ensures they're behind other content
		if (stageBars.length === 0) {
			return;
		}

		// Sort stage bars by X position (left to right)
		stageBars.sort((a, b) => a.x - b.x);

		// Send stage bars to back from LEFT TO RIGHT (X=100, X=150, X=200, X=300)
		// The LAST one sent to back (X=300) will be at the very bottom
		// The FIRST one sent to back (X=100) will be on top of other stage bars
		for (let i = 0; i < stageBars.length; i++) {
			stageBars[i].shape.sendToBack();
		}

		console.log(
			`Reordered ${stageBars.length} stage bars (leftmost on top, all behind other content)`,
		);
	} catch (e) {
		console.error(`Error reordering stage bars: ${e.message}`);
		// Don't throw error to avoid interrupting the main flow
	}
}

/**
 * Manually reorder all stage bars on the current slide
 * Useful for fixing layering after manual adjustments
 */
function reorderCurrentSlideStageBars() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const currentPage = selection.getCurrentPage();

		if (!currentPage) {
			throw new Error("No current slide selected");
		}

		reorderStageBars(currentPage);

		return "Stage bars reordered successfully (leftmost on top)";
	} catch (e) {
		const errorMsg = `Error reordering stage bars: ${e.message}`;
		console.error(errorMsg);
		throw new Error(errorMsg);
	}
}
