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
