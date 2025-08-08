/**
 * Utility functions for applying Markdown-style formatting to text in Google Slides
 */

/**
 * Applies bold formatting to text enclosed in double asterisks (**text**)
 * in the selected text elements.
 *
 * @return {boolean} True if the operation was successful
 */
function applyMarkdownBoldFormatting() {
	const selection = SlidesApp.getActivePresentation().getSelection();
	const pageElementRange = selection.getPageElementRange();

	if (!pageElementRange) {
		Logger.log("No element selected.");
		return false;
	}

	const elements = pageElementRange.getPageElements();

	elements.forEach((element) => {
		if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const shape = element.asShape();
			const textRange = shape.getText();
			const originalText = textRange.asString();

			// Find all **text** format matches
			const matches = [...originalText.matchAll(/\*\*(.+?)\*\*/g)];

			if (matches.length === 0) {
				return; // No markdown bold formatting found in this element
			}

			let newText = "";
			let lastIndex = 0;
			const formattingRanges = [];

			matches.forEach((match) => {
				const matchStart = match.index;
				const matchEnd = match.index + match[0].length;
				const content = match[1];

				// Add text before the match
				newText += originalText.substring(lastIndex, matchStart);

				// Record the position of formatted text in the new text
				const formatStart = newText.length;
				newText += content;
				const formatEnd = newText.length;

				// Store range (end is exclusive)
				formattingRanges.push({ start: formatStart, end: formatEnd });

				lastIndex = matchEnd;
			});

			// Add remaining original text
			newText += originalText.substring(lastIndex);

			// Replace text
			shape.getText().setText(newText);

			// Get updated textRange, as setText() resets it
			const updatedTextRange = shape.getText();

			// Apply formatting
			formattingRanges.forEach(({ start, end }) => {
				const range = updatedTextRange.getRange(start, end);
				range.getTextStyle().setBold(true);

				// Check if main_color is defined in the global scope
				try {
					if (typeof main_color !== "undefined") {
						range.getTextStyle().setForegroundColor(main_color);
					}
				} catch (e) {
					// If main_color is not defined, we just skip setting the color
					Logger.log("Note: main_color not defined, skipping color formatting");
				}
			});
		}
	});

	return true;
}

/**
 * Exposes the function for use in the menu
 */
function runApplyMarkdownBoldFormatting() {
	return applyMarkdownBoldFormatting();
}

/**
 * Calculates the optimal font size based on text length
 * Uses a formula to adjust font size to fit approximately 8 lines × 36 characters per line
 *
 * @param {string} text - The text to calculate font size for
 * @return {number} The calculated font size (between 8 and 18)
 */
function getFontSize(text) {
	const N = text.length; // 字數
	const baseFont = 18; // 基準字級
	const baseCapacity = 600; // Adjusted for better visual fit

	// Use a gentler reduction curve
	return Math.max(
		12,
		Math.min(baseFont, Math.floor(baseFont * Math.pow(baseCapacity / N, 0.3))),
	);
}
