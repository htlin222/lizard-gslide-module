/**
 * Shows a dialog to set gap between selected shapes.
 */
function showSetGapDialog() {
	const ui = SlidesApp.getUi();

	// Check if shapes are selected
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectedShapes = selection.getPageElementRange()
		? selection
				.getPageElementRange()
				.getPageElements()
				.filter(
					(element) =>
						element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
				)
		: [];

	if (selectedShapes.length < 2) {
		ui.alert(
			"Error",
			"Please select at least 2 shapes to adjust gaps.",
			ui.ButtonSet.OK,
		);
		return;
	}

	// Create and show the dialog
	const htmlOutput = HtmlService.createHtmlOutput(createSetGapDialogHtml())
		.setWidth(300)
		.setHeight(150);

	ui.showModalDialog(htmlOutput, "Set Gap Between Shapes");
}

/**
 * Creates the HTML content for the set gap dialog.
 * @return {string} The HTML content.
 */
function createSetGapDialogHtml() {
	return `
		<!DOCTYPE html>
		<html>
		<head>
			<base target="_top">
			<style>
				body { font-family: Arial, sans-serif; margin: 10px; font-size: 14px; }
				.form-group { margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; }
				label { font-weight: bold; flex: 1; margin-right: 10px; }
				input[type="number"] { width: 80px; padding: 4px 8px; text-align: center; border: 1px solid #ccc; border-radius: 3px; }
				.button-container { display: flex; justify-content: flex-end; margin-top: 20px; }
				button { padding: 8px 16px; background-color: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer; }
				button:hover { background-color: #2a75f3; }
				.input-suffix { font-size: 12px; color: #666; margin-left: 5px; }
			</style>
		</head>
		<body>
			<div class="form-group">
				<label for="targetGap">Target Gap:</label>
				<div>
					<input type="number" id="targetGap" min="0" value="7">
					<span class="input-suffix">pt</span>
				</div>
			</div>
			<div class="button-container">
				<button onclick="submitForm()">Set Gap</button>
			</div>
			
			<script>
				function submitForm() {
					const targetGap = parseInt(document.getElementById('targetGap').value);
					
					if (targetGap < 0 || isNaN(targetGap)) {
						alert('Please enter a valid gap value (0 or greater).');
						return;
					}
					
					google.script.run
						.withSuccessHandler(function() {
							google.script.host.close();
						})
						.withFailureHandler(function(error) {
							alert('Error: ' + error);
						})
						.setGapBetweenShapes(targetGap);
				}
			</script>
		</body>
		</html>
	`;
}

/**
 * Adjusts gaps between selected shapes to a target value by resizing and repositioning them.
 * @param {number} targetGap - The desired gap in points.
 */
function setGapBetweenShapes(targetGap) {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		const selectedShapes = selection.getPageElementRange()
			? selection
					.getPageElementRange()
					.getPageElements()
					.filter(
						(element) =>
							element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
					)
					.map((element) => element.asShape())
			: [];

		if (selectedShapes.length < 2) {
			throw new Error("Please select at least 2 shapes to adjust gaps.");
		}

		// Sort shapes by position (top-left to bottom-right)
		selectedShapes.sort((a, b) => {
			const aTop = a.getTop();
			const bTop = b.getTop();
			if (Math.abs(aTop - bTop) < 5) {
				// Same row (within 5pt tolerance)
				return a.getLeft() - b.getLeft(); // Sort by left position
			}
			return aTop - bTop; // Sort by top position
		});

		// Group shapes by rows (shapes with similar Y positions)
		const rows = [];
		let currentRow = [selectedShapes[0]];

		for (let i = 1; i < selectedShapes.length; i++) {
			const shape = selectedShapes[i];
			const lastShape = currentRow[currentRow.length - 1];

			// If shapes are on roughly the same horizontal line (within 5pt)
			if (Math.abs(shape.getTop() - lastShape.getTop()) < 5) {
				currentRow.push(shape);
			} else {
				rows.push(currentRow);
				currentRow = [shape];
			}
		}
		rows.push(currentRow);

		// Adjust horizontal gaps within each row
		for (const row of rows) {
			if (row.length < 2) continue;

			adjustHorizontalGaps(row, targetGap);
		}

		// Adjust vertical gaps between rows
		if (rows.length > 1) {
			adjustVerticalGaps(rows, targetGap);
		}

		console.log(
			`Successfully adjusted gaps to ${targetGap}pt for ${selectedShapes.length} shapes`,
		);
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			`An error occurred: ${error.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Adjusts horizontal gaps between shapes in a row.
 * Maintains the total group size - only redistributes space between shapes and gaps.
 * @param {Array} shapes - Array of shapes in the same row.
 * @param {number} targetGap - The target gap in points.
 */
function adjustHorizontalGaps(shapes, targetGap) {
	if (shapes.length < 2) return;

	// Calculate total group bounds (this stays constant)
	const leftmost = Math.min(...shapes.map((s) => s.getLeft()));
	const rightmost = Math.max(...shapes.map((s) => s.getLeft() + s.getWidth()));
	const totalGroupWidth = rightmost - leftmost; // This remains constant

	// Calculate total gap space needed
	const totalGapSpace = targetGap * (shapes.length - 1);

	// Calculate total width available for shapes (group size minus gaps)
	const totalShapeWidth = totalGroupWidth - totalGapSpace;

	// Get original widths for proportional distribution
	const originalWidths = shapes.map((s) => s.getWidth());
	const totalOriginalWidth = originalWidths.reduce((sum, w) => sum + w, 0);

	// Redistribute widths proportionally within the fixed group size
	let currentLeft = leftmost;

	for (let i = 0; i < shapes.length; i++) {
		const shape = shapes[i];
		// Calculate new width based on proportion of original width
		const widthProportion = originalWidths[i] / totalOriginalWidth;
		const newWidth = totalShapeWidth * widthProportion;

		// Set new position and width, keep original height
		shape.setLeft(currentLeft);
		shape.setWidth(newWidth);
		// Height remains unchanged

		currentLeft += newWidth + targetGap;
	}
}

/**
 * Adjusts vertical gaps between rows of shapes.
 * Maintains the total group size - only redistributes space between shapes and gaps.
 * @param {Array} rows - Array of shape rows.
 * @param {number} targetGap - The target gap in points.
 */
function adjustVerticalGaps(rows, targetGap) {
	if (rows.length < 2) return;

	// Calculate total group bounds (this stays constant)
	const topmost = Math.min(...rows[0].map((s) => s.getTop()));
	const bottommost = Math.max(
		...rows[rows.length - 1].map((s) => s.getTop() + s.getHeight()),
	);
	const totalGroupHeight = bottommost - topmost; // This remains constant

	// Calculate total gap space needed
	const totalGapSpace = targetGap * (rows.length - 1);

	// Calculate total height available for shapes (group size minus gaps)
	const totalShapeHeight = totalGroupHeight - totalGapSpace;

	// Get original row heights for proportional distribution
	const originalRowHeights = rows.map((row) =>
		Math.max(...row.map((s) => s.getHeight())),
	);
	const totalOriginalHeight = originalRowHeights.reduce((sum, h) => sum + h, 0);

	// Redistribute heights proportionally within the fixed group size
	let currentTop = topmost;

	for (let i = 0; i < rows.length; i++) {
		const row = rows[i];
		// Calculate new height based on proportion of original height
		const heightProportion = originalRowHeights[i] / totalOriginalHeight;
		const newRowHeight = totalShapeHeight * heightProportion;

		for (const shape of row) {
			// Adjust position and height, keep original width
			shape.setTop(currentTop);
			shape.setHeight(newRowHeight);
			// Width remains unchanged
		}

		currentTop += newRowHeight + targetGap;
	}
}

/**
 * Smart function to analyze parent-child relationships and reset gaps/padding uniformly.
 * Select multiple children + 1 parent, and this will reverse-engineer and reset the layout.
 */
function showSmartGapResetDialog() {
	const ui = SlidesApp.getUi();

	// Check if shapes are selected
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectedShapes = selection.getPageElementRange()
		? selection
				.getPageElementRange()
				.getPageElements()
				.filter(
					(element) =>
						element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
				)
				.map((element) => element.asShape())
		: [];

	if (selectedShapes.length < 3) {
		ui.alert(
			"Error",
			"Please select at least 3 shapes (1 parent + 2+ children) for smart gap reset.",
			ui.ButtonSet.OK,
		);
		return;
	}

	try {
		// Analyze the selection to identify parent and children
		const analysis = analyzeParentChildLayout(selectedShapes);

		if (!analysis.success) {
			ui.alert(
				"Error",
				analysis.error || "Could not identify parent-child relationships.",
				ui.ButtonSet.OK,
			);
			return;
		}

		// Create and show the dialog with detected values
		const htmlOutput = HtmlService.createHtmlOutput(
			createSmartGapResetDialogHtml(analysis),
		)
			.setWidth(400)
			.setHeight(300);

		ui.showModalDialog(htmlOutput, "Smart Gap & Padding Reset");
	} catch (error) {
		ui.alert("Error", `Analysis failed: ${error.message}`, ui.ButtonSet.OK);
	}
}

/**
 * Analyzes selected shapes to identify parent-child relationships and current spacing.
 * @param {Array} shapes - Array of selected shapes
 * @return {Object} Analysis result with parent, children, and current spacing values
 */
function analyzeParentChildLayout(shapes) {
	try {
		// Step 1: Identify the parent (largest shape that contains others)
		const parent = identifyParentShape(shapes);
		if (!parent) {
			return {
				success: false,
				error:
					"Could not identify parent shape. Select the largest containing shape.",
			};
		}

		// Step 2: Identify children (shapes inside parent)
		const children = shapes.filter(
			(shape) => shape.getObjectId() !== parent.getObjectId(),
		);
		if (children.length < 2) {
			return {
				success: false,
				error: "Need at least 2 child shapes inside parent.",
			};
		}

		// Step 3: Analyze current layout to extract spacing values
		const layoutAnalysis = analyzeCurrentSpacing(parent, children);

		return {
			success: true,
			parent: parent,
			children: children,
			currentSpacing: layoutAnalysis,
			childCount: children.length,
		};
	} catch (error) {
		return { success: false, error: error.message };
	}
}

/**
 * Identifies the parent shape (largest that encompasses others).
 * @param {Array} shapes - Array of shapes
 * @return {Shape|null} The parent shape or null if not found
 */
function identifyParentShape(shapes) {
	// Find the shape with largest area that contains most other shapes
	let bestParent = null;
	let maxContainedCount = 0;

	for (const candidate of shapes) {
		const containedCount = shapes.filter(
			(other) =>
				other.getObjectId() !== candidate.getObjectId() &&
				isShapeContainedIn(other, candidate),
		).length;

		if (containedCount > maxContainedCount) {
			maxContainedCount = containedCount;
			bestParent = candidate;
		}
	}

	// Must contain at least 2 other shapes to be considered parent
	return maxContainedCount >= 2 ? bestParent : null;
}

/**
 * Checks if shape A is contained within shape B.
 * @param {Shape} shapeA - The shape to check if contained
 * @param {Shape} shapeB - The potential container shape
 * @return {boolean} True if A is inside B
 */
function isShapeContainedIn(shapeA, shapeB) {
	const aLeft = shapeA.getLeft();
	const aTop = shapeA.getTop();
	const aRight = aLeft + shapeA.getWidth();
	const aBottom = aTop + shapeA.getHeight();

	const bLeft = shapeB.getLeft();
	const bTop = shapeB.getTop();
	const bRight = bLeft + shapeB.getWidth();
	const bBottom = bTop + shapeB.getHeight();

	// A is contained in B if all A's edges are inside B's bounds
	return (
		aLeft >= bLeft && aTop >= bTop && aRight <= bRight && aBottom <= bBottom
	);
}

/**
 * Analyzes current spacing between parent and children to extract gap/padding values.
 * @param {Shape} parent - The parent shape
 * @param {Array} children - Array of child shapes
 * @return {Object} Current spacing analysis
 */
function analyzeCurrentSpacing(parent, children) {
	const parentLeft = parent.getLeft();
	const parentTop = parent.getTop();
	const parentRight = parentLeft + parent.getWidth();
	const parentBottom = parentTop + parent.getHeight();

	// Group children by rows (similar Y positions)
	const rows = groupChildrenByRows(children);

	// Calculate current padding (distance from parent edges)
	const leftPadding = Math.min(
		...children.map((child) => child.getLeft() - parentLeft),
	);
	const rightPadding = Math.min(
		...children.map(
			(child) => parentRight - (child.getLeft() + child.getWidth()),
		),
	);
	const topPadding = Math.min(
		...children.map((child) => child.getTop() - parentTop),
	);
	const bottomPadding = Math.min(
		...children.map(
			(child) => parentBottom - (child.getTop() + child.getHeight()),
		),
	);

	// Calculate current gaps between children
	const horizontalGaps = [];
	const verticalGaps = [];

	// Horizontal gaps within rows
	for (const row of rows) {
		if (row.length > 1) {
			const sortedRow = row.sort((a, b) => a.getLeft() - b.getLeft());
			for (let i = 0; i < sortedRow.length - 1; i++) {
				const gap =
					sortedRow[i + 1].getLeft() -
					(sortedRow[i].getLeft() + sortedRow[i].getWidth());
				horizontalGaps.push(Math.round(gap));
			}
		}
	}

	// Vertical gaps between rows
	if (rows.length > 1) {
		const sortedRows = rows.sort(
			(a, b) =>
				Math.min(...a.map((s) => s.getTop())) -
				Math.min(...b.map((s) => s.getTop())),
		);
		for (let i = 0; i < sortedRows.length - 1; i++) {
			const row1Bottom = Math.max(
				...sortedRows[i].map((s) => s.getTop() + s.getHeight()),
			);
			const row2Top = Math.min(...sortedRows[i + 1].map((s) => s.getTop()));
			const gap = row2Top - row1Bottom;
			verticalGaps.push(Math.round(gap));
		}
	}

	return {
		padding: Math.round(Math.min(leftPadding, rightPadding)),
		paddingTop: Math.round(topPadding),
		paddingBottom: Math.round(bottomPadding),
		horizontalGap:
			horizontalGaps.length > 0
				? Math.round(
						horizontalGaps.reduce((sum, gap) => sum + gap, 0) /
							horizontalGaps.length,
					)
				: 7,
		verticalGap:
			verticalGaps.length > 0
				? Math.round(
						verticalGaps.reduce((sum, gap) => sum + gap, 0) /
							verticalGaps.length,
					)
				: 7,
		rows: rows.length,
		maxColumns: Math.max(...rows.map((row) => row.length)),
	};
}

/**
 * Groups child shapes by rows based on similar Y positions.
 * @param {Array} children - Array of child shapes
 * @return {Array} Array of row arrays
 */
function groupChildrenByRows(children) {
	if (children.length === 0) return [];

	// Sort by Y position first
	const sorted = children.sort((a, b) => a.getTop() - b.getTop());
	const tolerance = 10; // pt tolerance for same row

	const rows = [];
	let currentRow = [sorted[0]];

	for (let i = 1; i < sorted.length; i++) {
		const shape = sorted[i];
		const lastShape = currentRow[currentRow.length - 1];

		// If shapes are on roughly the same horizontal line
		if (Math.abs(shape.getTop() - lastShape.getTop()) < tolerance) {
			currentRow.push(shape);
		} else {
			rows.push(currentRow);
			currentRow = [shape];
		}
	}
	rows.push(currentRow);

	return rows;
}

/**
 * Creates the HTML content for the smart gap reset dialog.
 * @param {Object} analysis - The analysis result from analyzeParentChildLayout
 * @return {string} The HTML content
 */
function createSmartGapResetDialogHtml(analysis) {
	const spacing = analysis.currentSpacing;

	return `
		<!DOCTYPE html>
		<html>
		<head>
			<base target="_top">
			<style>
				body { font-family: Arial, sans-serif; margin: 15px; font-size: 14px; }
				.header { margin-bottom: 20px; padding: 10px; background-color: #f5f5f5; border-radius: 5px; }
				.header h3 { margin: 0 0 5px 0; color: #333; }
				.header p { margin: 0; color: #666; font-size: 12px; }
				.form-section { margin-bottom: 20px; }
				.form-section h4 { margin: 0 0 10px 0; color: #444; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
				.form-group { margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; }
				label { font-weight: bold; flex: 1; margin-right: 10px; }
				input[type="number"] { width: 80px; padding: 4px 8px; text-align: center; border: 1px solid #ccc; border-radius: 3px; }
				.current-value { font-size: 12px; color: #666; margin-left: 5px; }
				.button-container { display: flex; justify-content: flex-end; margin-top: 20px; gap: 10px; }
				button { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; }
				.btn-primary { background-color: #4285f4; color: white; }
				.btn-primary:hover { background-color: #2a75f3; }
				.btn-secondary { background-color: #f0f0f0; color: #333; }
				.btn-secondary:hover { background-color: #e0e0e0; }
				.input-suffix { font-size: 12px; color: #666; margin-left: 5px; }
				.detected-info { background-color: #e8f5e8; padding: 8px; border-radius: 3px; font-size: 12px; }
			</style>
		</head>
		<body>
			<div class="header">
				<h3>Smart Gap & Padding Reset</h3>
				<p>Detected: 1 parent + ${analysis.childCount} children (${spacing.rows} rows, max ${spacing.maxColumns} columns)</p>
			</div>
			
			<div class="form-section">
				<h4>Current Detected Values</h4>
				<div class="detected-info">
					Padding: ${spacing.padding}pt • Top: ${spacing.paddingTop}pt • H-Gap: ${spacing.horizontalGap}pt • V-Gap: ${spacing.verticalGap}pt
				</div>
			</div>
			
			<div class="form-section">
				<h4>New Values</h4>
				<div class="form-group">
					<label for="newPadding">Padding (L/R/B):</label>
					<div>
						<input type="number" id="newPadding" min="0" value="${spacing.padding}">
						<span class="input-suffix">pt</span>
						<span class="current-value">(current: ${spacing.padding}pt)</span>
					</div>
				</div>
				<div class="form-group">
					<label for="newPaddingTop">Top Padding:</label>
					<div>
						<input type="number" id="newPaddingTop" min="0" value="${spacing.paddingTop}">
						<span class="input-suffix">pt</span>
						<span class="current-value">(current: ${spacing.paddingTop}pt)</span>
					</div>
				</div>
				<div class="form-group">
					<label for="newHorizontalGap">Horizontal Gap:</label>
					<div>
						<input type="number" id="newHorizontalGap" min="0" value="7">
						<span class="input-suffix">pt</span>
						<span class="current-value">(current: ${spacing.horizontalGap}pt)</span>
					</div>
				</div>
				<div class="form-group">
					<label for="newVerticalGap">Vertical Gap:</label>
					<div>
						<input type="number" id="newVerticalGap" min="0" value="7">
						<span class="input-suffix">pt</span>
						<span class="current-value">(current: ${spacing.verticalGap}pt)</span>
					</div>
				</div>
			</div>
			
			<div class="button-container">
				<button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
				<button class="btn-primary" onclick="submitReset()">Apply Reset</button>
			</div>
			
			<script>
				function submitReset() {
					const newPadding = parseInt(document.getElementById('newPadding').value);
					const newPaddingTop = parseInt(document.getElementById('newPaddingTop').value);
					const newHorizontalGap = parseInt(document.getElementById('newHorizontalGap').value);
					const newVerticalGap = parseInt(document.getElementById('newVerticalGap').value);
					
					if (newPadding < 0 || newPaddingTop < 0 || newHorizontalGap < 0 || newVerticalGap < 0 ||
						isNaN(newPadding) || isNaN(newPaddingTop) || isNaN(newHorizontalGap) || isNaN(newVerticalGap)) {
						alert('Please enter valid values (0 or greater).');
						return;
					}
					
					google.script.run
						.withSuccessHandler(function() {
							google.script.host.close();
						})
						.withFailureHandler(function(error) {
							alert('Error: ' + error);
						})
						.applySmartGapReset(newPadding, newPaddingTop, newHorizontalGap, newVerticalGap);
				}
			</script>
		</body>
		</html>
	`;
}

/**
 * Applies the smart gap reset with new spacing values.
 * @param {number} newPadding - New padding value in points (left/right/bottom)
 * @param {number} newPaddingTop - New top padding value in points
 * @param {number} newHorizontalGap - New horizontal gap value in points
 * @param {number} newVerticalGap - New vertical gap value in points
 */
function applySmartGapReset(
	newPadding,
	newPaddingTop,
	newHorizontalGap,
	newVerticalGap,
) {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const selectedShapes = selection.getPageElementRange()
			? selection
					.getPageElementRange()
					.getPageElements()
					.filter(
						(element) =>
							element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
					)
					.map((element) => element.asShape())
			: [];

		if (selectedShapes.length < 3) {
			throw new Error(
				"Please select at least 3 shapes (1 parent + 2+ children) for smart gap reset.",
			);
		}

		// Re-analyze the selection to get current layout
		const analysis = analyzeParentChildLayout(selectedShapes);
		if (!analysis.success) {
			throw new Error(
				analysis.error || "Could not identify parent-child relationships.",
			);
		}

		const parent = analysis.parent;
		const children = analysis.children;

		// Get parent properties
		const parentLeft = parent.getLeft();
		const parentTop = parent.getTop();
		const parentWidth = parent.getWidth();
		const parentHeight = parent.getHeight();

		// Group children by rows
		const rows = groupChildrenByRows(children);

		// Calculate new layout dimensions
		const availableWidth = parentWidth - newPadding * 2;
		const availableHeight = parentHeight - newPaddingTop - newPadding;

		// Calculate new row height
		const newRowHeight =
			(availableHeight - newVerticalGap * (rows.length - 1)) / rows.length;

		if (newRowHeight <= 0) {
			throw new Error(
				"New spacing values are too large for the parent shape size.",
			);
		}

		// Reposition and resize all child shapes
		for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
			const row = rows[rowIndex];
			const columnsInRow = row.length;

			// Calculate new column width for this row
			const newColumnWidth =
				(availableWidth - newHorizontalGap * (columnsInRow - 1)) / columnsInRow;

			if (newColumnWidth <= 0) {
				console.warn(
					`Row ${rowIndex + 1} has too many columns for the new spacing`,
				);
				continue;
			}

			// Calculate the starting Y position for this row
			const newRowTop =
				parentTop + newPaddingTop + rowIndex * (newRowHeight + newVerticalGap);

			// Reposition shapes in this row
			for (let colIndex = 0; colIndex < columnsInRow; colIndex++) {
				const shape = row[colIndex];
				const newColumnLeft =
					parentLeft +
					newPadding +
					colIndex * (newColumnWidth + newHorizontalGap);

				// Apply new position and size
				shape.setLeft(newColumnLeft);
				shape.setTop(newRowTop);
				shape.setWidth(newColumnWidth);
				shape.setHeight(newRowHeight);
			}
		}

		console.log(
			`Successfully applied smart gap reset: padding=${newPadding}pt, top=${newPaddingTop}pt, h-gap=${newHorizontalGap}pt, v-gap=${newVerticalGap}pt`,
		);
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			`Smart gap reset failed: ${error.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}
