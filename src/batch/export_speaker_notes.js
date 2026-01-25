// 🎤 EXPORT SPEAKER NOTES MODULE - Extract all speaker notes as JSON
/**
 * Exports all speaker notes from the presentation as a JSON dictionary.
 * Format: { "1": "note text", "2": "note text", ... }
 * where keys are slide numbers (1-indexed).
 */

/**
 * Extract speaker notes from a single slide
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to extract notes from
 * @returns {string} The speaker notes text, or empty string if none
 */
function extractSpeakerNotes(slide) {
	const notesPage = slide.getNotesPage();
	if (!notesPage) return "";

	const notesShapes = notesPage.getShapes();
	for (const shape of notesShapes) {
		const placeholderType = shape.getPlaceholderType();
		if (placeholderType === SlidesApp.PlaceholderType.BODY) {
			const notesText = shape.getText().asString().trim();
			if (notesText) {
				return notesText;
			}
		}
	}
	return "";
}

/**
 * Export all speaker notes as a JSON dictionary
 * @returns {Object} Object with slide numbers as keys and notes as values
 */
function exportAllSpeakerNotes() {
	const presentation = SlidesApp.getActivePresentation();
	const slides = presentation.getSlides();
	const speakerNotes = {};

	for (let i = 0; i < slides.length; i++) {
		const slideNumber = i + 1; // 1-indexed slide number
		const notes = extractSpeakerNotes(slides[i]);
		// Include all slides, even those without notes (empty string)
		speakerNotes[slideNumber] = notes;
	}

	return speakerNotes;
}

/**
 * Export speaker notes with additional metadata
 * @returns {Object} Object with presentation info and notes
 */
function exportSpeakerNotesWithMetadata() {
	const presentation = SlidesApp.getActivePresentation();
	const slides = presentation.getSlides();

	const result = {
		presentationId: presentation.getId(),
		presentationName: presentation.getName(),
		totalSlides: slides.length,
		exportedAt: new Date().toISOString(),
		notes: {},
	};

	for (let i = 0; i < slides.length; i++) {
		const slideNumber = i + 1;
		const slide = slides[i];
		const notes = extractSpeakerNotes(slide);

		// Get slide title for reference
		let slideTitle = "";
		const pageElements = slide.getPageElements();
		for (const element of pageElements) {
			if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
				const shape = element.asShape();
				const placeholderType = shape.getPlaceholderType();
				if (
					placeholderType === SlidesApp.PlaceholderType.TITLE ||
					placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE
				) {
					slideTitle = shape.getText().asString().trim();
					break;
				}
			}
		}

		result.notes[slideNumber] = {
			title: slideTitle,
			notes: notes,
		};
	}

	return result;
}

/**
 * Show dialog with exported speaker notes as JSON
 */
function showExportSpeakerNotesDialog() {
	const speakerNotes = exportAllSpeakerNotes();
	const jsonString = JSON.stringify(speakerNotes, null, 2);

	const htmlTemplate = HtmlService.createTemplateFromFile(
		"src/components/export-speaker-notes-dialog.html",
	);
	htmlTemplate.jsonContent = jsonString;

	const html = htmlTemplate
		.evaluate()
		.setWidth(700)
		.setHeight(500)
		.setTitle("Export Speaker Notes");

	SlidesApp.getUi().showModalDialog(html, "🎤 匯出演講者備註");
}

/**
 * Show dialog with speaker notes including metadata
 */
function showExportSpeakerNotesWithMetadataDialog() {
	const speakerNotes = exportSpeakerNotesWithMetadata();
	const jsonString = JSON.stringify(speakerNotes, null, 2);

	const htmlTemplate = HtmlService.createTemplateFromFile(
		"src/components/export-speaker-notes-dialog.html",
	);
	htmlTemplate.jsonContent = jsonString;

	const html = htmlTemplate
		.evaluate()
		.setWidth(700)
		.setHeight(500)
		.setTitle("Export Speaker Notes (with metadata)");

	SlidesApp.getUi().showModalDialog(html, "🎤 匯出演講者備註 (含詳細資訊)");
}
