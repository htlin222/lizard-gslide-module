/**
 * Shape Styler for Child Shapes
 * Handles all shape styling and formatting logic
 */

/**
 * Applies bold style transformation to a shape if its text is wrapped in [brackets].
 * @param {Shape} shape - The shape to apply bold style to
 * @return {Object} Result object with applied status, original text, and new text
 */
function applyBoldStyleTransformation(shape) {
	const result = {
		applied: false,
		originalText: "",
		newText: "",
		debug: {},
	};

	try {
		// Check if the shape has text
		const textRange = shape.getText();
		if (!textRange) {
			result.debug.hasText = false;
			return result;
		}

		// Get the text content, remove line breaks and normalize whitespace
		const textContent = textRange
			.asString()
			.replace(/\r?\n|\r/g, " ")
			.replace(/\s+/g, " ")
			.trim();
		result.originalText = textContent;
		result.debug.hasText = true;
		result.debug.textLength = textContent.length;
		result.debug.startsWithBracket = textContent.startsWith("[");
		result.debug.endsWithBracket = textContent.endsWith("]");

		// Check if text is wrapped in [] and has content between them
		if (
			textContent.startsWith("[") &&
			textContent.endsWith("]") &&
			textContent.length > 2 // Must have at least 1 character between the []
		) {
			// Apply special formatting for [text]
			// Set border with main_color and 1pt weight
			shape.getBorder().setWeight(1);
			shape.getBorder().getLineFill().setSolidFill(main_color);

			// Remove the [] markers and set text
			const cleanText = textContent.substring(1, textContent.length - 1);
			textRange.setText(cleanText);

			// Set text to main_color and bold
			const textStyle = textRange.getTextStyle();
			textStyle.setForegroundColor(main_color);
			textStyle.setBold(true);

			result.newText = cleanText;
			result.applied = true;
			result.debug.transformationApplied = true;
		} else {
			result.debug.transformationApplied = false;
			result.debug.reason = "Text not wrapped in [] or too short";
		}
	} catch (error) {
		console.log(`Error in applyBoldStyleTransformation: ${error.message}`);
		result.debug.error = error.message;
	}

	return result;
}

/**
 * Applies white fill and white stroke to a shape.
 * If the text is wrapped in **asterisks**, applies special formatting.
 * @param {Shape} shape - The shape to apply white style to.
 */
function applyWhiteStyle(shape) {
	try {
		// Always set white fill first
		const fill = shape.getFill();
		fill.setSolidFill("#FFFFFF");

		// Check if we should apply bold style transformation
		const result = applyBoldStyleTransformation(shape);

		if (!result.applied) {
			// Apply normal white style if no bold transformation was applied
			// Set white border with 0.1pt weight
			shape.getBorder().setWeight(0.1);
			shape.getBorder().getLineFill().setSolidFill("#FFFFFF");

			// Optionally set text color to black for visibility on white background
			if (shape.getText()) {
				const textStyle = shape.getText().getTextStyle();
				textStyle.setForegroundColor("#000000");
			}
		}
	} catch (error) {
		// Log error but continue execution
		console.log(`Error applying white style: ${error.message}`);
	}
}
