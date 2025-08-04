/**
 * Dialog Manager for Child Shapes Creation
 * Handles all user interface interactions for creating child shapes
 */

/**
 * Shows a dialog to input parameters for creating child shapes inside a parent shape.
 */
function showCreateChildShapesDialog() {
	const ui = SlidesApp.getUi();

	// Check if a shape is selected
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

	if (selectedShapes.length !== 1) {
		ui.alert(
			"Error",
			"Please select exactly one shape to create child shapes in.",
			ui.ButtonSet.OK,
		);
		return;
	}

	// Create and show the dialog
	const htmlOutput = HtmlService.createHtmlOutputFromFile(
		"src/components/create-child-shapes-dialog.html",
	)
		.setWidth(350)
		.setHeight(280);

	ui.showModalDialog(htmlOutput, "Create Child Shapes");
}

/**
 * Applies bold styling to selected shape if text is wrapped in **asterisks**.
 * This is a standalone function for testing the bold styling feature.
 */
function applyBoldStyleToSelectedShape() {
	try {
		// Get the active presentation and selection
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		// Check if a shape is selected
		const selectedElements = selection.getPageElementRange()
			? selection.getPageElementRange().getPageElements()
			: [];

		const selectedShapes = selectedElements.filter(
			(element) =>
				element.getPageElementType() === SlidesApp.PageElementType.SHAPE,
		);

		if (selectedShapes.length !== 1) {
			SlidesApp.getUi().alert(
				"Error",
				"Please select exactly one shape to apply bold styling.",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		const shape = selectedShapes[0].asShape();

		// Apply the bold styling transformation
		const result = applyBoldStyleTransformation(shape);

		if (result.applied) {
			SlidesApp.getUi().alert(
				"Success",
				`Bold styling applied. Text changed from "${result.originalText}" to "${result.newText}"`,
				SlidesApp.getUi().ButtonSet.OK,
			);
		} else {
			// Create debug message
			let debugMessage = "No bold styling applied.\n\n";
			debugMessage += `Original text: "${result.originalText}"\n`;
			debugMessage += `Text length: ${result.debug.textLength}\n`;
			debugMessage += `Starts with **: ${result.debug.startsWithAsterisk}\n`;
			debugMessage += `Ends with **: ${result.debug.endsWithAsterisk}\n`;

			if (result.debug.error) {
				debugMessage += `\nError: ${result.debug.error}`;
			} else if (result.debug.reason) {
				debugMessage += `\nReason: ${result.debug.reason}`;
			} else {
				debugMessage +=
					"\nText must be wrapped in **asterisks** (e.g., **text**).";
			}

			SlidesApp.getUi().alert(
				"Debug Info",
				debugMessage,
				SlidesApp.getUi().ButtonSet.OK,
			);
		}
	} catch (error) {
		SlidesApp.getUi().alert(
			"Error",
			`An error occurred: ${error.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}
