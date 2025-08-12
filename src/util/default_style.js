/**
 * util_default_style.js
 *
 * Provides utility functions for applying predefined styles to shapes and textboxes
 * in Google Slides presentations.
 */

/**
 * Apply a default style to the selected shape or textbox
 * @param {number} styleNumber - The style number to apply (1, 2, or 3)
 */
function applyDefaultStyle(styleNumber) {
	const presentation = SlidesApp.getActivePresentation();
	const selection = presentation.getSelection();
	const selectedElements = selection.getPageElementRange()
		? selection.getPageElementRange().getPageElements()
		: [];

	// Define the styles
	const styles = {
		1: {
			// Base color fill, main color border and text
			borderColor: main_color,
			borderWidth: 1,
			fillColor: base_color,
			textColor: main_color,
		},
		2: {
			// Main color fill, text color border and base color text
			borderColor: main_color,
			borderWidth: 1,
			fillColor: main_color,
			textColor: base_color,
		},
		3: {
			// Sub1 color fill, main color border and base color text
			borderColor: main_color,
			borderWidth: 1,
			fillColor: sub1_color,
			textColor: main_color,
		},
		4: {
			// Base color fill, accent color border and accent color text
			borderColor: accent_color,
			borderWidth: 1,
			fillColor: base_color,
			textColor: accent_color,
		},
		5: {
			// Accent color fill, base color border and base color text
			borderColor: accent_color,
			borderWidth: 1,
			fillColor: accent_color,
			textColor: accent_color,
		},
		6: {
			// Base color fill, base color border and main color text
			borderColor: base_color,
			borderWidth: 1,
			fillColor: base_color,
			textColor: main_color,
		},
	};

	// Get the selected style
	const style = styles[styleNumber];
	if (!style) {
		Logger.log("Invalid style number: " + styleNumber);
		return;
	}

	// Apply the style to each selected element
	for (let i = 0; i < selectedElements.length; i++) {
		const element = selectedElements[i];

		// Check if the element is a shape
		if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const shape = element.asShape();

			try {
				// First, ensure the shape has text to avoid 'has no text' error
				// This must be done before any text styling operations
				if (!shape.getText().asString()) {
					shape.getText().setText("TEXT_HERE");
				}

				// Apply border
				shape.getBorder().setWeight(style.borderWidth);
				shape.getBorder().getLineFill().setSolidFill(style.borderColor);

				// Apply fill
				shape.getFill().setSolidFill(style.fillColor);

				// Apply text color
				shape.getText().getTextStyle().setForegroundColor(style.textColor);
			} catch (e) {
				Logger.log("Error applying style to shape: " + e.toString());
			}
		}
		// Check if the element is a text box (which is also a shape in Google Slides)
		else if (
			element.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
			element.asShape().getShapeType() === SlidesApp.ShapeType.TEXT_BOX
		) {
			const textBox = element.asShape();

			try {
				// First, ensure the textbox has text to avoid 'has no text' error
				// This must be done before any text styling operations
				if (!textBox.getText().asString()) {
					textBox.getText().setText("TEXT_HERE");
				}

				// Apply border
				textBox.getBorder().setWeight(style.borderWidth);
				textBox.getBorder().getLineFill().setSolidFill(style.borderColor);

				// Apply fill
				textBox.getFill().setSolidFill(style.fillColor);

				// Apply text color
				textBox.getText().getTextStyle().setForegroundColor(style.textColor);
			} catch (e) {
				Logger.log("Error applying style to textbox: " + e.toString());
			}
		}
	}
}

/**
 * Apply style 1: White fill, main color border and text
 * This function is exposed to the sidebar HTML
 */
function applyStyle1() {
	applyDefaultStyle(1);
	return true; // Return a value to confirm execution to the sidebar
}

/**
 * Apply style 2: Main color fill, white text
 * This function is exposed to the sidebar HTML
 */
function applyStyle2() {
	applyDefaultStyle(2);
	return true; // Return a value to confirm execution to the sidebar
}

/**
 * Apply style 3: Sub1 color fill, main color border and base color text
 * This function is exposed to the sidebar HTML
 */
function applyStyle3() {
	applyDefaultStyle(3);
	return true; // Return a value to confirm execution to the sidebar
}

/**
 * Apply style 4: Base color fill, accent color border and accent color text
 * This function is exposed to the sidebar HTML
 */
function applyStyle4() {
	applyDefaultStyle(4);
	return true; // Return a value to confirm execution to the sidebar
}

/**
 * Apply style 5: Accent color fill, base color border and base color text
 * This function is exposed to the sidebar HTML
 */
function applyStyle5() {
	applyDefaultStyle(5);
	return true; // Return a value to confirm execution to the sidebar
}

/**
 * Apply style 6: Base color fill, base color border and main color text
 * This function is exposed to the sidebar HTML
 */
function applyStyle6() {
	applyDefaultStyle(6);
	return true; // Return a value to confirm execution to the sidebar
}
