/**
 * Speaker Note Generator - Generates speaker notes using OpenAI API based on slide content
 */

/**
 * Shows the Speaker Note sidebar
 */
function showSpeakerNoteSidebar() {
	try {
		const sidebar = createSpeakerNoteSidebar();
		SlidesApp.getUi().showSidebar(sidebar);
	} catch (e) {
		console.error("Error showing Speaker Note sidebar: " + e.message);
		SlidesApp.getUi().alert(
			"Error",
			"Could not open the Speaker Note sidebar: " + e.message,
		);
	}
}

/**
 * Creates the Speaker Note sidebar HTML
 */
function createSpeakerNoteSidebar() {
	return HtmlService.createHtmlOutputFromFile(
		"src/components/speakerNoteSidebar",
	)
		.setTitle("AI Speaker Notes Generator")
		.setWidth(400);
}

/**
 * Gets all text content from the current slide
 * @return {Object} Object containing slide text and metadata
 */
function getCurrentSlideContent() {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const currentPage = selection.getCurrentPage();

		if (
			!currentPage ||
			currentPage.getPageType() !== SlidesApp.PageType.SLIDE
		) {
			throw new Error("Please select a slide first");
		}

		// Convert currentPage to a Slide object
		const currentSlide = currentPage.asSlide();
		const slides = presentation.getSlides();
		const slideNumber = slides.indexOf(currentSlide) + 1;
		const totalSlides = slides.length;

		// Get all shapes and their text content
		const shapes = currentSlide.getShapes();
		const textContent = [];

		shapes.forEach((shape) => {
			try {
				// Check if shape has text
				if (shape.getText && shape.getText().asString().trim() !== "") {
					const text = shape.getText().asString().trim();
					const shapeType = shape.getShapeType().toString();

					textContent.push({
						type: shapeType,
						text: text,
						length: text.length,
					});
				}
			} catch (e) {
				// Skip shapes that don't have text or cause errors
				console.log("Skipping shape due to error: " + e.message);
			}
		});

		// Get tables if any
		const tables = currentSlide.getTables();
		tables.forEach((table, tableIndex) => {
			try {
				const numRows = table.getNumRows();
				const numCols = table.getNumColumns();
				let tableText = `Table ${tableIndex + 1} (${numRows}x${numCols}): `;

				for (let row = 0; row < numRows; row++) {
					for (let col = 0; col < numCols; col++) {
						const cell = table.getCell(row, col);
						const cellText = cell.getText().asString().trim();
						if (cellText) {
							tableText += cellText + " | ";
						}
					}
					tableText += "\n";
				}

				if (tableText.length > 50) {
					// Only add if table has content
					textContent.push({
						type: "TABLE",
						text: tableText.trim(),
						length: tableText.length,
					});
				}
			} catch (e) {
				console.log("Error reading table: " + e.message);
			}
		});

		// Combine all text content
		const allText = textContent.map((item) => item.text).join("\n\n");

		return {
			slideNumber: slideNumber,
			totalSlides: totalSlides,
			textContent: textContent,
			allText: allText,
			hasContent: textContent.length > 0,
			title: currentSlide.getTitle
				? currentSlide.getTitle()
				: "Slide " + slideNumber,
		};
	} catch (e) {
		console.error("Error getting slide content: " + e.message);
		throw new Error("Failed to get slide content: " + e.message);
	}
}

/**
 * Gets the current speaker notes for the slide
 * @return {string} Current speaker notes
 */
function getCurrentSpeakerNotes() {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const currentPage = selection.getCurrentPage();

		if (
			!currentPage ||
			currentPage.getPageType() !== SlidesApp.PageType.SLIDE
		) {
			return "";
		}

		// Convert currentPage to a Slide object
		const slide = currentPage.asSlide();
		const notesPage = slide.getNotesPage();
		const shape = notesPage.getSpeakerNotesShape();

		const notesText = shape ? shape.getText().asString() : "";
		return notesText.trim();
	} catch (e) {
		console.error(`Error getting speaker notes: ${e.message}`);
		return "";
	}
}

/**
 * Generates speaker notes via Groq for the given slide content.
 *
 * SECURITY: the API key is NOT a parameter. It is read server-side from the
 * current user's UserProperties inside callGroq_() (see src/util/aiKey.js).
 * The client never sends the key.
 *
 * @param {string} slideContent - Content from the slide.
 * @param {string} prompt - User prompt/instructions.
 * @return {Object} API response with generated content.
 */
function generateSpeakerNotesAI(slideContent, prompt) {
	if (!slideContent || slideContent.trim() === "") {
		return {
			success: false,
			error: "No slide content found to process",
			generatedText: "",
		};
	}

	const systemMessage =
		"You are a helpful assistant that creates speaker notes for presentations. Generate clear, concise, and helpful speaker notes based on the slide content provided.";

	const userMessage = `Please create speaker notes for this slide content:

SLIDE CONTENT:
${slideContent}

USER INSTRUCTIONS:
${prompt || "Create professional speaker notes that help explain and expand on the slide content."}

Please provide speaker notes that:
1. Explain the key points in more detail
2. Provide context and background information
3. Suggest transitions and speaking points
4. Are natural to read aloud
5. Help the presenter engage with the audience`;

	return callGroq_(systemMessage, userMessage, {
		maxTokens: 1000,
		temperature: 0.7,
	});
}

/**
 * Appends generated text to the speaker notes of the current slide
 * @param {string} textToAppend - Text to append to speaker notes
 * @return {Object} Success status and message
 */
function appendToSpeakerNotes(textToAppend) {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const currentPage = selection.getCurrentPage();

		if (
			!currentPage ||
			currentPage.getPageType() !== SlidesApp.PageType.SLIDE
		) {
			throw new Error("Please select a slide first");
		}

		// Convert currentPage to a Slide object
		const slide = currentPage.asSlide();
		const notesPage = slide.getNotesPage();
		const shape = notesPage.getSpeakerNotesShape();

		if (!shape) {
			throw new Error("Could not access speaker notes shape");
		}

		// Get current notes and append new text
		const currentText = shape.getText().asString();

		// Create new text with proper spacing
		const newText = currentText.trim()
			? `${currentText.trim()}\n\n${textToAppend}`
			: textToAppend;

		// Update the speaker notes
		shape.getText().setText(newText);

		return {
			success: true,
			message: "Speaker notes updated successfully",
		};
	} catch (e) {
		console.error(`Error appending to speaker notes: ${e.message}`);
		return {
			success: false,
			message: `Failed to update speaker notes: ${e.message}`,
		};
	}
}

/**
 * Replaces the speaker notes of the current slide with new text
 * @param {string} newText - New text for speaker notes
 * @return {Object} Success status and message
 */
function replaceSpeakerNotes(newText) {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		const currentPage = selection.getCurrentPage();

		if (
			!currentPage ||
			currentPage.getPageType() !== SlidesApp.PageType.SLIDE
		) {
			throw new Error("Please select a slide first");
		}

		// Convert currentPage to a Slide object
		const slide = currentPage.asSlide();
		const notesPage = slide.getNotesPage();
		const shape = notesPage.getSpeakerNotesShape();

		if (!shape) {
			throw new Error("Could not access speaker notes shape");
		}

		// Replace the speaker notes with new text
		shape.getText().setText(newText || "");

		return {
			success: true,
			message: "Speaker notes replaced successfully",
		};
	} catch (e) {
		console.error(`Error replacing speaker notes: ${e.message}`);
		return {
			success: false,
			message: `Failed to replace speaker notes: ${e.message}`,
		};
	}
}
