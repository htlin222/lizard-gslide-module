/**
 * Markdown to Slides Converter Utility
 *
 * This utility converts markdown text to Google Slides with the following rules:
 * - H1 headings become SECTION_HEADER slides
 * - H2 headings become TITLE_AND_BODY slides
 * - Text below H2 headings becomes bullet points in the body
 *
 * The approach is modular:
 * 1. Parse markdown into a structured format
 * 2. Create all slides based on the parsed structure
 * 3. Add content to each slide
 * 4. Apply formatting (like bullet points) to the content
 * 5. Apply markdown bold formatting (**text**) to the content
 */

/**
 * Shows a dialog for the user to paste markdown content
 */
function showMarkdownDialog() {
	const html = HtmlService.createTemplateFromFile(
		"src/components/md2slides-dialog",
	)
		.evaluate()
		.setWidth(600)
		.setHeight(500)
		.setTitle("Markdown to Slides Converter");

	SlidesApp.getUi().showModalDialog(html, "Markdown to Slides Converter");
}

/**
 * Converts markdown text to slides using the modular approach
 * @param {string} markdownText - The markdown text to convert
 * @return {boolean} - Success status
 */
function convertMarkdownToSlides(markdownText) {
	try {
		// Step 1: Clean and parse the markdown into a structured format
		const cleanedText = cleanMarkdownText(markdownText);
		const slideStructure = parseMarkdownToStructure(cleanedText);

		if (slideStructure.length === 0) {
			debugLog(
				"md2slides",
				"convertMarkdownToSlides",
				"No slides to create from markdown",
			);
			return false;
		}

		// Step 2: Determine where to insert the slides
		const presentation = SlidesApp.getActivePresentation();
		const insertIndex = getInsertIndex(presentation);

		// Step 3: Create all slides first
		const createdSlides = createSlidesFromStructure(
			slideStructure,
			presentation,
			insertIndex,
		);

		if (createdSlides.length === 0) {
			debugLog(
				"md2slides",
				"convertMarkdownToSlides",
				"Failed to create slides",
			);
			return false;
		}

		// Step 4: Add content to all slides
		const contentSuccess = addContentToSlides(createdSlides);
		if (!contentSuccess) {
			debugLog(
				"md2slides",
				"convertMarkdownToSlides",
				"Failed to add content to slides",
			);
		}

		// Step 5: Apply list formatting to all TITLE_AND_BODY slides
		const listFormattingSuccess = applyListFormattingToSlides(createdSlides);
		if (!listFormattingSuccess) {
			debugLog(
				"md2slides",
				"convertMarkdownToSlides",
				"Failed to apply list formatting",
			);
		}

		// Step 6: Apply markdown bold formatting to all slides
		applyMarkdownBoldToSlides(createdSlides.map((obj) => obj.slide));

		debugLog(
			"md2slides",
			"convertMarkdownToSlides",
			`Successfully created ${createdSlides.length} slides`,
		);
		return true;
	} catch (error) {
		const errorObj = createMd2SlidesError(
			"md2slides",
			"convertMarkdownToSlides",
			"Failed to convert markdown to slides",
			error,
		);
		console.error(
			`Error converting markdown to slides: ${JSON.stringify(errorObj)}`,
		);
		return false;
	}
}

// Note: parseMarkdownToStructure is now handled by parser.js module

// Note: getInsertIndex is now handled by slideCreator.js module

// Note: Content management is now handled by contentManager.js module

// Note: List formatting is now handled by listFormatter.js module

/**
 * Apply markdown bold formatting to text enclosed in double asterisks (**text**)
 * in all slides
 * @param {Array} slides - Array of slides to apply formatting to
 */
function applyMarkdownBoldToSlides(slides) {
	for (let i = 0; i < slides.length; i++) {
		try {
			const slide = slides[i];
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
						const textRange = shape.getText();
						const originalText = textRange.asString();

						// Find all **text** format matches
						const matches = [...originalText.matchAll(/\*\*(.+?)\*\*/g)];

						if (matches.length === 0) {
							continue; // No markdown bold formatting found in this element
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
						textRange.setText(newText);

						// Apply formatting
						formattingRanges.forEach(({ start, end }) => {
							const range = textRange.getRange(start, end);
							range.getTextStyle().setBold(true);

							// Check if main_color is defined in the global scope
							try {
								if (typeof main_color !== "undefined") {
									range.getTextStyle().setForegroundColor(main_color);
								}
							} catch (e) {
								// If main_color is not defined, we just skip setting the color
								Logger.log(
									"Note: main_color not defined, skipping color formatting",
								);
							}
						});
					}
				} catch (e) {
					Logger.log(
						"Error processing shape for bold formatting: " + e.message,
					);
					// Continue with next shape
				}
			}

			Logger.log("Applied markdown bold formatting to slide " + (i + 1));
		} catch (error) {
			Logger.log(
				"Error applying markdown bold formatting to slide " +
					(i + 1) +
					": " +
					error.message,
			);
			// Continue with next slide
		}
	}
}

/**
 * Registers the md2slides utility in the menu
 */
function registerMd2SlidesMenu() {
	const ui = SlidesApp.getUi();
	ui.createMenu("Lizard Utilities")
		.addItem("Markdown to Slides", "showMarkdownDialog")
		.addToUi();
}
