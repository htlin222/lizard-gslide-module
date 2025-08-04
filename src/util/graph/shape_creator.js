/**
 * Shape Creator for Child Shapes
 * Handles all shape creation and positioning logic
 */

// Configuration constants for child shapes
const DEFAULT_PADDING = 10;
const DEFAULT_PADDING_TOP = 30;
const DEFAULT_GAP = 10;
const FOOTER_BOX_HEIGHT = 15;
const HOME_PLATE_HEIGHT = 10;

/**
 * Creates child shapes inside the selected parent shape.
 * @param {number} rows - Number of rows in the grid.
 * @param {number} columns - Number of columns in the grid.
 * @param {number} padding - Padding inside the parent shape in points.
 * @param {number} paddingTop - Top padding inside the parent shape in points.
 * @param {number} gap - Gap size between child shapes in points.
 */
function createChildShapesInSelected(rows, columns, padding, paddingTop, gap) {
	try {
		// Get the active presentation and selection
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const slide = selection.getCurrentPage();

		// Get the selected shape
		const selectedElements = selection.getPageElementRange().getPageElements();
		const selectedShapes = selectedElements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (selectedShapes.length !== 1) {
			throw new Error(
				"Please select exactly one shape to create child shapes in.",
			);
		}

		const parentShape = selectedShapes[0].asShape();

		// Log parent shape information for debugging
		console.log("Parent shape type: " + parentShape.getShapeType());
		console.log("Parent shape ID: " + parentShape.getObjectId());
		console.log(
			"Parent position: Left " +
				parentShape.getLeft() +
				", Top " +
				parentShape.getTop(),
		);
		console.log(
			"Parent size: Width " +
				parentShape.getWidth() +
				", Height " +
				parentShape.getHeight(),
		);
		console.log("Parent rotation: " + parentShape.getRotation() + " degrees");

		// Get the properties of the parent shape
		const parentLeft = parentShape.getLeft();
		const parentTop = parentShape.getTop();
		const parentWidth = parentShape.getWidth();
		const parentHeight = parentShape.getHeight();
		const parentRotation = parentShape.getRotation();

		// Calculate the available space inside the parent shape after padding
		const availableWidth = parentWidth - padding * 2;
		const availableHeight = parentHeight - paddingTop - padding;

		// Calculate the dimensions for each child shape
		const childWidth = (availableWidth - gap * (columns - 1)) / columns;
		const childHeight = (availableHeight - gap * (rows - 1)) / rows;

		// Validate that child shapes will have positive dimensions
		if (childWidth <= 0 || childHeight <= 0) {
			throw new Error(
				"Padding and gap values are too large for the parent shape size.",
			);
		}

		// Create an array to store all child shapes
		const childShapes = [];

		// Create the grid of child shapes
		for (let row = 0; row < rows; row++) {
			for (let col = 0; col < columns; col++) {
				// Calculate position for the child shape relative to parent
				const childLeft = parentLeft + padding + col * (childWidth + gap);
				const childTop = parentTop + paddingTop + row * (childHeight + gap);

				// Create the child shape with the same type as parent
				const childShape = slide.insertShape(
					parentShape.getShapeType(),
					childLeft,
					childTop,
					childWidth,
					childHeight,
				);

				// Explicitly set position to ensure accuracy
				childShape.setLeft(childLeft);
				childShape.setTop(childTop);
				childShape.setWidth(childWidth);
				childShape.setHeight(childHeight);

				// Apply rotation if the parent shape has any
				if (parentRotation !== 0) {
					childShape.setRotation(parentRotation);
				}

				// Apply white fill and white stroke to child shape
				applyWhiteStyle(childShape);

				// Add to our array of child shapes
				childShapes.push(childShape);
			}
		}

		// Set titles for parent and child shapes
		parentShape.setTitle("PARENT");

		// Set title for each child shape and bring them forward
		for (let i = 0; i < childShapes.length; i++) {
			childShapes[i].setTitle("CHILD");
			childShapes[i].bringForward();
		}

		// Log completion message
		console.log(
			"Successfully created " +
				childShapes.length +
				" child shapes in a " +
				rows +
				"x" +
				columns +
				" grid with " +
				padding +
				"pt padding and " +
				gap +
				"pt gaps",
		);
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			"An error occurred: " + error.message,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Creates child shapes with a specific layout structure (supporting variable columns per row).
 * @param {Shape} parentShape - The parent shape
 * @param {Object} layout - The grid layout structure
 * @param {number} padding - Padding in points
 * @param {number} paddingTop - Top padding in points
 * @param {number} gap - Gap between shapes in points
 */
function createChildShapesWithLayout(
	parentShape,
	layout,
	padding,
	paddingTop,
	gap,
) {
	// Get the slide from the presentation selection
	const presentation = SlidesApp.getActivePresentation();
	const selection = presentation.getSelection();
	const slide = selection.getCurrentPage();

	// Get parent properties
	const parentLeft = parentShape.getLeft();
	const parentTop = parentShape.getTop();
	const parentWidth = parentShape.getWidth();
	const parentHeight = parentShape.getHeight();
	const parentRotation = parentShape.getRotation();

	// Calculate available space
	const availableWidth = parentWidth - padding * 2;
	const availableHeight = parentHeight - paddingTop - padding;

	// Calculate row height
	const rowHeight = (availableHeight - gap * (layout.rows - 1)) / layout.rows;

	if (rowHeight <= 0) {
		throw new Error("Parent shape is too small for the specified layout.");
	}

	const childShapes = [];

	// Create shapes for each row
	for (let rowIndex = 0; rowIndex < layout.rows; rowIndex++) {
		const rowInfo = layout.rowData[rowIndex];
		const row = rowInfo.cells || rowInfo; // Handle both new and old format
		const columnsInRow = row.length;
		const homePlates = rowInfo.homePlates || []; // Get home plate positions

		// Calculate column width for this specific row
		const columnWidth =
			(availableWidth - gap * (columnsInRow - 1)) / columnsInRow;

		if (columnWidth <= 0) {
			console.warn(
				`Row ${rowIndex + 1} has too many columns for the available width`,
			);
			continue;
		}

		// Calculate the starting Y position for this row
		const rowTop = parentTop + paddingTop + rowIndex * (rowHeight + gap);

		// Create shapes for each column in this row
		for (let colIndex = 0; colIndex < columnsInRow; colIndex++) {
			const columnLeft = parentLeft + padding + colIndex * (columnWidth + gap);

			// Create the child shape
			const childShape = slide.insertShape(
				parentShape.getShapeType(),
				columnLeft,
				rowTop,
				columnWidth,
				rowHeight,
			);

			// Set precise positioning
			childShape.setLeft(columnLeft);
			childShape.setTop(rowTop);
			childShape.setWidth(columnWidth);
			childShape.setHeight(rowHeight);

			// Apply rotation if needed
			if (parentRotation !== 0) {
				childShape.setRotation(parentRotation);
			}

			// Process cell text for footer boxes (but not vertical splits yet)
			const cellText = row[colIndex];

			// Only process footer if this cell doesn't have vertical splits (no --)
			if (cellText && !cellText.includes("--")) {
				// Process for footer box syntax (text)
				const footerData = processFooterBoxTextForSegment(cellText);

				// Adjust child shape height if footer box is needed
				let adjustedRowHeight = rowHeight;
				if (footerData.hasFooter) {
					adjustedRowHeight = rowHeight - FOOTER_BOX_HEIGHT;
					childShape.setHeight(adjustedRowHeight);
				}

				// Set the main text content (without footer text)
				if (footerData.mainText) {
					const textRange = childShape.getText();
					textRange.setText(footerData.mainText);
				}

				// Apply styling after text is set
				applyWhiteStyle(childShape);

				// Create footer box if needed
				if (footerData.hasFooter) {
					createFooterBox(
						slide,
						columnLeft,
						rowTop + adjustedRowHeight, // Position at bottom of adjusted cell
						columnWidth,
						FOOTER_BOX_HEIGHT,
						footerData.footerText,
						parentRotation,
					);
				}
			} else {
				// For cells with -- or no text, just set the raw text
				if (cellText) {
					const textRange = childShape.getText();
					textRange.setText(cellText);
				}

				// Apply basic styling
				applyWhiteStyle(childShape);
			}

			// Set title for child shape
			childShape.setTitle("CHILD");

			childShapes.push(childShape);
		}

		// Create HOME_PLATE shapes for this row if any are specified
		for (const homePlatePosition of homePlates) {
			// Calculate position between cells
			const leftCellIndex = homePlatePosition - 1;
			const rightCellIndex = homePlatePosition;

			// Only create if both cells exist
			if (leftCellIndex >= 0 && rightCellIndex < columnsInRow) {
				const leftCellRight =
					parentLeft +
					padding +
					leftCellIndex * (columnWidth + gap) +
					columnWidth;
				const rightCellLeft =
					parentLeft + padding + rightCellIndex * (columnWidth + gap);

				// Position HOME_PLATE in the gap between cells
				const homePlateLeft = leftCellRight;
				const homePlateTop = rowTop + (rowHeight - HOME_PLATE_HEIGHT) / 2; // Center vertically
				const homePlateWidth = gap; // Use gap width
				const homePlateHeight = HOME_PLATE_HEIGHT;

				// Create the HOME_PLATE shape
				const homePlate = slide.insertShape(
					SlidesApp.ShapeType.HOME_PLATE,
					homePlateLeft,
					homePlateTop,
					homePlateWidth,
					homePlateHeight,
				);

				// Apply rotation if needed
				if (parentRotation !== 0) {
					homePlate.setRotation(parentRotation);
				}

				// Set main_color fill and border
				const fill = homePlate.getFill();
				fill.setSolidFill(main_color);

				homePlate.getBorder().setWeight(1);
				homePlate.getBorder().getLineFill().setSolidFill(main_color);

				// Don't set any title for HOME_PLATE shapes (they are connectors, not content)

				// Bring HOME_PLATE forward
				homePlate.bringForward();

				console.log(
					`Created HOME_PLATE at row ${rowIndex + 1}, position ${homePlatePosition}`,
				);
			}
		}
	}

	// Bring all child shapes forward
	for (const childShape of childShapes) {
		childShape.bringForward();
	}

	// Set parent shape text alignment to top
	parentShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);

	// Set title for parent shape
	parentShape.setTitle("PARENT");

	// Post-process vertical splits after all shapes are created
	postProcessVerticalSplits(
		slide,
		layout,
		childShapes,
		parentRotation,
		padding,
		paddingTop,
		gap,
		parentLeft,
		parentTop,
		parentWidth,
		parentHeight,
	);

	console.log(
		`Created ${childShapes.length} child shapes with variable column layout`,
	);
}

/**
 * Creates a title RECTANGLE above the parent shape with the given text.
 * @param {Shape} parentShape - The parent shape
 * @param {string} titleText - The title text
 * @param {Slide} slide - The slide to add the rectangle to
 * @return {Shape} The created rectangle
 */
function createTitleTextBoxFromText(parentShape, titleText, slide) {
	// Get parent shape properties
	const parentLeft = parentShape.getLeft();
	const parentTop = parentShape.getTop();
	const parentWidth = parentShape.getWidth();
	const parentRotation = parentShape.getRotation();

	// Create rectangle positioned 30pt above parent shape
	const rectangleLeft = parentLeft;
	const rectangleTop = parentTop - 30; // 30pt above
	const rectangleWidth = parentWidth; // Same width as parent
	const rectangleHeight = 30; // 30pt height

	// Create the rectangle
	const titleRectangle = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		rectangleLeft,
		rectangleTop,
		rectangleWidth,
		rectangleHeight,
	);

	// Apply rotation if parent has any
	if (parentRotation !== 0) {
		titleRectangle.setRotation(parentRotation);
	}

	// Set the fill color to main_color
	const fill = titleRectangle.getFill();
	fill.setSolidFill(main_color);

	// Set white border with 0.1pt weight
	titleRectangle.getBorder().setWeight(0.1);
	titleRectangle.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Set the text content
	titleRectangle.getText().setText(titleText);

	// Style the text - 14pt, bold, white
	const textStyle = titleRectangle.getText().getTextStyle();
	textStyle.setFontSize(14);
	textStyle.setBold(true);
	textStyle.setForegroundColor("#FFFFFF");

	// Center the text vertically and horizontally
	titleRectangle.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Bring rectangle forward
	titleRectangle.bringForward();

	console.log(`Created title rectangle with text: "${titleText}"`);

	return titleRectangle;
}

/**
 * Creates a footer box at the bottom of a cell.
 * @param {Slide} slide - The slide to add the footer box to
 * @param {number} left - Left position of the footer box
 * @param {number} top - Top position of the footer box
 * @param {number} width - Width of the footer box
 * @param {number} height - Height of the footer box
 * @param {string} text - Text content for the footer box
 * @param {number} rotation - Rotation angle to apply
 * @return {Shape} The created footer box shape
 */
function createFooterBox(slide, left, top, width, height, text, rotation) {
	// Create the footer box rectangle
	const footerBox = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		left,
		top,
		width,
		height,
	);

	// Apply rotation if needed
	if (rotation !== 0) {
		footerBox.setRotation(rotation);
	}

	// Set background color to main_color
	const fill = footerBox.getFill();
	fill.setSolidFill(main_color);

	// Set white border with 0.1pt weight
	footerBox.getBorder().setWeight(0.1);
	footerBox.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Set the text content
	footerBox.getText().setText(text);

	// Style the text - white color, 10pt font size, centered
	const textStyle = footerBox.getText().getTextStyle();
	textStyle.setForegroundColor("#FFFFFF");
	textStyle.setFontSize(10);

	// Center the text both vertically and horizontally
	footerBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Bring footer box forward
	footerBox.bringForward();

	console.log(`Created footer box with text: "${text}"`);

	return footerBox;
}

/**
 * Creates a vertically split cell with multiple segments and divider lines.
 * @param {Slide} slide - The slide to add shapes to
 * @param {Shape} parentCell - The parent cell shape to modify
 * @param {Object} footerBoxData - The processed footer box data with segments
 * @param {number} left - Left position of the cell
 * @param {number} top - Top position of the cell
 * @param {number} width - Width of the cell
 * @param {number} height - Height of the cell
 * @param {number} rotation - Rotation angle to apply
 */
function createVerticalSplitCell(
	slide,
	parentCell,
	footerBoxData,
	left,
	top,
	width,
	height,
	rotation,
) {
	const segments = footerBoxData.segments;
	const segmentCount = segments.length;

	// Validate segment count to prevent division by zero
	if (segmentCount <= 0) {
		console.error("Invalid segment count for vertical split:", segmentCount);
		return;
	}

	const segmentHeight = height / segmentCount;

	// Validate dimensions to prevent Google Apps Script errors
	if (segmentHeight <= 0 || width <= 0) {
		console.error("Invalid dimensions for vertical split segments:", {
			segmentHeight,
			width,
			height,
			segmentCount,
		});
		return;
	}

	// Set the parent cell to be invisible (no fill, no border) as it will be a container
	const parentFill = parentCell.getFill();
	parentFill.setSolidFill("#FFFFFF"); // White background
	parentCell.getBorder().setWeight(0.1);
	parentCell.getBorder().getLineFill().setSolidFill("#FFFFFF");

	// Clear any existing text from parent cell
	parentCell.getText().setText("");

	// Create each segment as a separate text box
	for (let i = 0; i < segmentCount; i++) {
		const segment = segments[i];
		const segmentTop = top + i * segmentHeight;

		// Create a text box for this segment
		const segmentShape = slide.insertShape(
			SlidesApp.ShapeType.TEXT_BOX,
			left,
			segmentTop,
			width,
			segmentHeight,
		);

		// Apply rotation if needed
		if (rotation !== 0) {
			segmentShape.setRotation(rotation);
		}

		// Set transparent fill and no border for the text box
		const segmentFill = segmentShape.getFill();
		segmentFill.setTransparent();
		segmentShape.getBorder().setWeight(0);

		// Handle footer box for this segment
		let segmentTextHeight = segmentHeight;
		if (segment.hasFooter) {
			segmentTextHeight = segmentHeight - FOOTER_BOX_HEIGHT;
			segmentShape.setHeight(segmentTextHeight);

			// Create footer box for this segment
			createFooterBox(
				slide,
				left,
				segmentTop + segmentTextHeight,
				width,
				FOOTER_BOX_HEIGHT,
				segment.footerText,
				rotation,
			);
		}

		// Set the segment text
		if (segment.mainText) {
			const textRange = segmentShape.getText();
			textRange.setText(segment.mainText);

			// Apply styling (check for [bold] markers)
			applyWhiteStyle(segmentShape);
		}

		// Set title for segment
		segmentShape.setTitle(`SEGMENT_${i + 1}`);
		segmentShape.bringForward();
	}

	// Create divider lines between segments (but not after the last one)
	for (let i = 0; i < segmentCount - 1; i++) {
		const lineTop = top + (i + 1) * segmentHeight;
		createHorizontalDividerLine(slide, left, lineTop, width, rotation);
	}

	console.log(
		`Created vertically split cell with ${segmentCount} segments and ${segmentCount - 1} divider lines`,
	);
}

/**
 * Post-processes child shapes to handle vertical splits after all normal shapes are created.
 * @param {Slide} slide - The slide containing the shapes
 * @param {Object} layout - The grid layout structure
 * @param {Array} childShapes - Array of created child shapes
 * @param {number} parentRotation - Parent shape rotation
 * @param {number} padding - Padding in points
 * @param {number} paddingTop - Top padding in points
 * @param {number} gap - Gap between shapes in points
 * @param {number} parentLeft - Parent left position
 * @param {number} parentTop - Parent top position
 * @param {number} parentWidth - Parent width
 * @param {number} parentHeight - Parent height
 */
function postProcessVerticalSplits(
	slide,
	layout,
	childShapes,
	parentRotation,
	padding,
	paddingTop,
	gap,
	parentLeft,
	parentTop,
	parentWidth,
	parentHeight,
) {
	let shapeIndex = 0;

	// Calculate available space
	const availableWidth = parentWidth - padding * 2;
	const availableHeight = parentHeight - paddingTop - padding;
	const rowHeight = (availableHeight - gap * (layout.rows - 1)) / layout.rows;

	// Iterate through each row and column to find vertical splits
	for (let rowIndex = 0; rowIndex < layout.rows; rowIndex++) {
		const rowInfo = layout.rowData[rowIndex];
		const row = rowInfo.cells || rowInfo;
		const columnsInRow = row.length;
		const columnWidth =
			(availableWidth - gap * (columnsInRow - 1)) / columnsInRow;

		for (let colIndex = 0; colIndex < columnsInRow; colIndex++) {
			const cellText = row[colIndex];
			const childShape = childShapes[shapeIndex];

			// Check if this cell needs vertical splitting
			if (cellText && cellText.includes("--")) {
				// Split the text by --
				const segments = cellText
					.split("--")
					.map((s) => s.trim())
					.filter((s) => s !== "");

				if (segments.length >= 2) {
					// Calculate position for this cell
					const columnLeft =
						parentLeft + padding + colIndex * (columnWidth + gap);
					const rowTop = parentTop + paddingTop + rowIndex * (rowHeight + gap);

					// Log the calculated positions for debugging
					console.log(`About to create vertical split for "${cellText}":`, {
						columnLeft,
						rowTop,
						columnWidth,
						rowHeight,
						parentRotation,
						segments,
					});

					// Create simple vertical split using the footer box pattern
					createSimpleVerticalSplit(
						slide,
						columnLeft,
						rowTop,
						columnWidth,
						rowHeight,
						segments,
						parentRotation,
					);

					// Clear the original shape text since we've replaced it with segments
					childShape.getText().setText("");

					console.log(
						`Post-processed vertical split for cell: "${cellText}" into ${segments.length} segments`,
					);
				}
			}

			shapeIndex++;
		}
	}
}

/**
 * Creates simple vertical split shapes following the footer box pattern.
 * @param {Slide} slide - The slide to add shapes to
 * @param {number} left - Left position of the cell
 * @param {number} top - Top position of the cell
 * @param {number} width - Width of the cell
 * @param {number} height - Height of the cell
 * @param {Array} segments - Array of text segments
 * @param {number} rotation - Rotation angle to apply
 */
function createSimpleVerticalSplit(
	slide,
	left,
	top,
	width,
	height,
	segments,
	rotation,
) {
	const segmentCount = segments.length;
	const segmentHeight = height / segmentCount;

	// Log all input parameters for debugging
	console.log("createSimpleVerticalSplit called with:", {
		left,
		top,
		width,
		height,
		segmentCount,
		segmentHeight,
		segments,
		rotation,
	});

	// Validate ALL dimensions to prevent Google Apps Script errors
	if (left < 0 || top < 0 || width <= 0 || height <= 0 || segmentHeight <= 0) {
		console.error("Invalid dimensions for vertical split:", {
			left,
			top,
			width,
			height,
			segmentCount,
			segmentHeight,
		});
		return;
	}

	// Create each segment as a simple text box (like footer boxes)
	for (let i = 0; i < segmentCount; i++) {
		const segmentText = segments[i];
		const segmentTop = top + i * segmentHeight;

		// Process this segment for footer box syntax (text)
		const footerData = processFooterBoxTextForSegment(segmentText);

		// Log the exact parameters being passed to insertShape
		console.log(`Creating segment ${i + 1} with:`, {
			type: "TEXT_BOX",
			left,
			segmentTop,
			width,
			segmentHeight,
			segmentText,
			footerData,
		});

		// Adjust segment height if footer box is needed
		let adjustedSegmentHeight = segmentHeight;
		if (footerData.hasFooter) {
			adjustedSegmentHeight = segmentHeight - FOOTER_BOX_HEIGHT;
		}

		// Create a text box for this segment (same pattern as footer box)
		const segmentShape = slide.insertShape(
			SlidesApp.ShapeType.TEXT_BOX,
			left,
			segmentTop,
			width,
			adjustedSegmentHeight,
		);

		// Apply rotation if needed (same as footer box)
		if (rotation !== 0) {
			segmentShape.setRotation(rotation);
		}

		// Set the main text content (without footer text)
		if (footerData.mainText) {
			segmentShape.getText().setText(footerData.mainText);
		}

		// Apply the same styling as regular cells
		applyWhiteStyle(segmentShape);

		// Set font size and alignment to match regular cells
		if (segmentShape.getText()) {
			const textRange = segmentShape.getText();
			const textStyle = textRange.getTextStyle();
			textStyle.setFontSize(label_font_size); // Use global font size variable

			// Set horizontal centering (paragraph alignment)
			textRange
				.getParagraphStyle()
				.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
		}

		// Set vertical centering (content alignment)
		segmentShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

		// Create footer box if needed
		if (footerData.hasFooter) {
			createFooterBox(
				slide,
				left,
				segmentTop + adjustedSegmentHeight, // Position at bottom of adjusted segment
				width,
				FOOTER_BOX_HEIGHT,
				footerData.footerText,
				rotation,
			);
		}

		// Bring shape forward (same as footer box)
		segmentShape.bringForward();

		// Set title for identification
		segmentShape.setTitle(`VSPLIT_${i + 1}`);
	}

	// Create divider lines between segments (but not after the last one)
	// TEMPORARILY DISABLED FOR DEBUGGING
	console.log("Skipping divider lines to debug the error");
	/*
	for (let i = 0; i < segmentCount - 1; i++) {
		const lineTop = top + (i + 1) * segmentHeight;
		createHorizontalDividerLine(slide, left, lineTop, width, rotation);
	}
	*/

	console.log(`Created simple vertical split with ${segmentCount} segments`);
}

/**
 * Creates a horizontal divider line with main_color.
 * @param {Slide} slide - The slide to add the line to
 * @param {number} left - Left position of the line
 * @param {number} top - Top position of the line
 * @param {number} width - Width of the line
 * @param {number} rotation - Rotation angle to apply
 * @return {Shape} The created line shape
 */
function createHorizontalDividerLine(slide, left, top, width, rotation) {
	// Create a thin rectangle to serve as a horizontal line
	const lineHeight = 1; // 1pt height for the line
	const adjustedTop = top - lineHeight / 2; // Center the line on the specified top position

	// Log parameters for debugging
	console.log("createHorizontalDividerLine called with:", {
		left,
		top,
		adjustedTop,
		width,
		lineHeight,
		rotation,
	});

	// Validate parameters before creating the line
	if (left < 0 || adjustedTop < 0 || width <= 0 || lineHeight <= 0) {
		console.error("Invalid line parameters:", {
			left,
			top,
			adjustedTop,
			width,
			lineHeight,
		});
		return null;
	}

	const line = slide.insertShape(
		SlidesApp.ShapeType.RECTANGLE,
		left,
		adjustedTop,
		width,
		lineHeight,
	);

	// Apply rotation if needed
	if (rotation !== 0) {
		line.setRotation(rotation);
	}

	// Set main_color fill and no border
	const fill = line.getFill();
	fill.setSolidFill(main_color);
	line.getBorder().setWeight(0);

	// Bring line forward
	line.bringForward();

	return line;
}
