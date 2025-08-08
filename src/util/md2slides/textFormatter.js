/**
 * Text Formatting Module
 *
 * Handles markdown text formatting: **bold**, ~strikethrough~, _italic_
 */

/**
 * Applies all markdown text formatting to slides
 * @param {Array} slides - Array of slide objects
 * @return {boolean} Success status
 */
function applyMarkdownFormattingToSlides(slides) {
	try {
		for (let i = 0; i < slides.length; i++) {
			const slide = slides[i];
			applyMarkdownFormattingToSlide(slide);
		}
		return true;
	} catch (error) {
		Logger.log(
			`Error applying markdown formatting to slides: ${error.message}`,
		);
		return false;
	}
}

/**
 * Applies markdown text formatting to a single slide
 * @param {Slide} slide - The slide to format
 * @return {boolean} Success status
 */
function applyMarkdownFormattingToSlide(slide) {
	try {
		const shapes = slide.getShapes();

		// Process each shape in the slide
		for (let j = 0; j < shapes.length; j++) {
			const shape = shapes[j];

			try {
				// Only process text boxes and placeholders
				if (
					shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX ||
					shape.getPlaceholderType
				) {
					applyMarkdownFormattingToTextRange(shape.getText());
				}
			} catch (shapeError) {
				Logger.log(`Error processing shape: ${shapeError.message}`);
			}
		}
		return true;
	} catch (error) {
		Logger.log(`Error applying markdown formatting to slide: ${error.message}`);
		return false;
	}
}

/**
 * Applies markdown text formatting to a text range
 * @param {TextRange} textRange - The text range to format
 * @return {boolean} Success status
 */
function applyMarkdownFormattingToTextRange(textRange) {
	try {
		const originalText = textRange.asString();

		// Define markdown patterns
		const patterns = [
			{ regex: /\*\*(.+?)\*\*/g, style: "bold" }, // **bold**
			{ regex: /~(.+?)~/g, style: "strikethrough" }, // ~strikethrough~
			{ regex: /_(.+?)_/g, style: "italic" }, // _italic_
		];

		// Find all formatting matches
		const allMatches = [];
		for (const pattern of patterns) {
			const matches = [...originalText.matchAll(pattern.regex)];
			for (const match of matches) {
				allMatches.push({
					start: match.index,
					end: match.index + match[0].length,
					content: match[1],
					style: pattern.style,
				});
			}
		}

		// If no matches, nothing to format
		if (allMatches.length === 0) {
			return true;
		}

		// Sort matches by position
		allMatches.sort((a, b) => a.start - b.start);

		// Build new text and collect formatting ranges
		let newText = "";
		let lastIndex = 0;
		const formattingRanges = [];

		for (const match of allMatches) {
			// Add text before the match
			newText += originalText.substring(lastIndex, match.start);

			// Record the position of formatted text in the new text
			const formatStart = newText.length;
			newText += match.content;
			const formatEnd = newText.length;

			// Store range with style info (end is exclusive)
			formattingRanges.push({
				start: formatStart,
				end: formatEnd,
				style: match.style,
			});

			lastIndex = match.end;
		}

		// Add remaining original text
		newText += originalText.substring(lastIndex);

		// Replace text
		textRange.setText(newText);

		// Apply formatting
		for (const { start, end, style } of formattingRanges) {
			try {
				const range = textRange.getRange(start, end);
				const textStyle = range.getTextStyle();

				// Apply the appropriate style
				switch (style) {
					case "bold":
						textStyle.setBold(true);
						// Apply color if main_color is available
						try {
							if (typeof main_color !== "undefined") {
								textStyle.setForegroundColor(main_color);
							}
						} catch (e) {
							// Skip color formatting if main_color not available
						}
						break;
					case "italic":
						textStyle.setItalic(true);
						break;
					case "strikethrough":
						textStyle.setStrikethrough(true);
						break;
				}
			} catch (styleError) {
				Logger.log(`Error applying ${style} formatting: ${styleError.message}`);
			}
		}

		return true;
	} catch (error) {
		Logger.log(`Error in applyMarkdownFormattingToTextRange: ${error.message}`);
		return false;
	}
}

/**
 * Processes markdown formatting for a specific text string
 * @param {string} text - The text to process
 * @return {Object} Object with cleanText and formattingInfo
 */
function processMarkdownText(text) {
	try {
		const patterns = [
			{ regex: /\*\*(.+?)\*\*/g, style: "bold" },
			{ regex: /~(.+?)~/g, style: "strikethrough" },
			{ regex: /_(.+?)_/g, style: "italic" },
		];

		const allMatches = [];
		for (const pattern of patterns) {
			const matches = [...text.matchAll(pattern.regex)];
			for (const match of matches) {
				allMatches.push({
					start: match.index,
					end: match.index + match[0].length,
					content: match[1],
					style: pattern.style,
				});
			}
		}

		// If no matches, return original text
		if (allMatches.length === 0) {
			return { cleanText: text, formattingInfo: [] };
		}

		// Sort matches by position
		allMatches.sort((a, b) => a.start - b.start);

		// Build clean text and collect formatting info
		let cleanText = "";
		let lastIndex = 0;
		const formattingInfo = [];

		for (const match of allMatches) {
			// Add text before the match
			cleanText += text.substring(lastIndex, match.start);

			// Record the position of formatted text in clean text
			const formatStart = cleanText.length;
			cleanText += match.content;
			const formatEnd = cleanText.length;

			// Store formatting info
			formattingInfo.push({
				start: formatStart,
				end: formatEnd,
				style: match.style,
			});

			lastIndex = match.end;
		}

		// Add remaining text
		cleanText += text.substring(lastIndex);

		return { cleanText, formattingInfo };
	} catch (error) {
		Logger.log(`Error processing markdown text: ${error.message}`);
		return { cleanText: text, formattingInfo: [] };
	}
}

/**
 * Applies formatting info to a text range
 * @param {TextRange} textRange - The text range to format
 * @param {Array} formattingInfo - Array of formatting objects
 * @return {boolean} Success status
 */
function applyFormattingToTextRange(textRange, formattingInfo) {
	try {
		for (const { start, end, style } of formattingInfo) {
			try {
				const range = textRange.getRange(start, end);
				const textStyle = range.getTextStyle();

				switch (style) {
					case "bold":
						textStyle.setBold(true);
						try {
							if (typeof main_color !== "undefined") {
								textStyle.setForegroundColor(main_color);
							}
						} catch (e) {
							// Skip color formatting if main_color not available
						}
						break;
					case "italic":
						textStyle.setItalic(true);
						break;
					case "strikethrough":
						textStyle.setStrikethrough(true);
						break;
				}
			} catch (styleError) {
				Logger.log(`Error applying ${style} formatting: ${styleError.message}`);
			}
		}
		return true;
	} catch (error) {
		Logger.log(`Error applying formatting to text range: ${error.message}`);
		return false;
	}
}
