/**
 * Slide Creator Module
 *
 * Handles creating slides in Google Slides based on structured data
 */

/**
 * Creates slides based on the parsed markdown structure
 * @param {Array} slideStructure - Array of slide objects from parser
 * @param {Presentation} presentation - The Google Slides presentation
 * @param {number} insertIndex - Where to insert the slides
 * @return {Array} Array of created slide objects with slide reference and info
 */
function createSlidesFromStructure(slideStructure, presentation, insertIndex) {
	const createdSlides = [];
	let currentIndex = insertIndex;

	for (let i = 0; i < slideStructure.length; i++) {
		const slideInfo = slideStructure[i];

		let slide;
		if (slideInfo.layout === "SECTION_HEADER") {
			slide = presentation.insertSlide(
				currentIndex,
				SlidesApp.PredefinedLayout.SECTION_HEADER,
			);
		} else if (slideInfo.layout === "TITLE_AND_BODY") {
			slide = presentation.insertSlide(
				currentIndex,
				SlidesApp.PredefinedLayout.TITLE_AND_BODY,
			);
		}

		createdSlides.push({
			slide: slide,
			info: slideInfo,
		});

		currentIndex++;
	}

	return createdSlides;
}

/**
 * Determines the index where new slides should be inserted
 * @param {Presentation} presentation - The active presentation
 * @return {number} - The index to insert slides at
 */
function getInsertIndex(presentation) {
	try {
		const selection = presentation.getSelection();

		if (selection) {
			const currentPage = selection.getCurrentPage();
			if (currentPage) {
				// Find the index of the current slide
				const slides = presentation.getSlides();
				for (let i = 0; i < slides.length; i++) {
					if (slides[i].getObjectId() === currentPage.getObjectId()) {
						// Insert after the current slide
						return i + 1;
					}
				}
			}
		}

		// Default to the end of the presentation if we can't determine the current slide
		return presentation.getSlides().length;
	} catch (error) {
		// Default to the end of the presentation
		return presentation.getSlides().length;
	}
}

/**
 * Creates a slide with a specific layout
 * @param {Presentation} presentation - The presentation to add to
 * @param {number} index - The index to insert at
 * @param {string} layout - The layout type ("SECTION_HEADER" or "TITLE_AND_BODY")
 * @return {Slide} The created slide
 */
function createSlideWithLayout(presentation, index, layout) {
	if (layout === "SECTION_HEADER") {
		return presentation.insertSlide(
			index,
			SlidesApp.PredefinedLayout.SECTION_HEADER,
		);
	}
	if (layout === "TITLE_AND_BODY") {
		return presentation.insertSlide(
			index,
			SlidesApp.PredefinedLayout.TITLE_AND_BODY,
		);
	}
	throw new Error(`Unknown layout type: ${layout}`);
}
