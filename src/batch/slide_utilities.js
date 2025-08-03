// 🚀 SLIDE UTILITIES MODULE - Slide analysis and manipulation helpers
/**
 * Utility functions for slide analysis and content extraction
 * - Section detection
 * - Title extraction
 * - Element deletion patterns
 */

/**
 * Ultra-efficient section header detection
 */
function getSectionHeadersUltra(slides) {
	const sections = [];
	for (let i = 0; i < slides.length; i++) {
		const slide = slides[i];
		if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') {
			// Early exit on first text found
			const shapes = slide.getShapes();
			for (const shape of shapes) {
				if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
					const text = shape.getText().asString().trim();
					if (text) {
						sections.push({ title: text, index: i, slideId: slide.getObjectId() });
						break; // Found text, stop searching this slide
					}
				}
			}
		}
	}
	return sections;
}

/**
 * Extract main title from first slide
 */
function getMainTitleFromFirstSlide(slide) {
	const elements = slide.getPageElements();
	for (let el of elements) {
		if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const txt = el.asShape().getText().asString().trim();
			if (txt) return txt;
		}
	}
	return '';
}

/**
 * Get title from outline slide (second slide if it exists)
 */
function getOutlineSlideTitle(slide) {
	let title = '';
	const placeholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
	if (placeholder && placeholder.asShape) {
		title = placeholder.asShape().getText().asString().trim();
	} else {
		// Quick text search
		const shapes = slide.getShapes();
		for (const shape of shapes) {
			if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
				const txt = shape.getText().asString().trim();
				if (txt) { 
					title = txt; 
					break; 
				}
			}
		}
	}
	return title;
}

/**
 * Ultra-efficient batch delete of old elements
 */
function batchDeleteAllElements(slides, requests) {
	const deletePatterns = ['tab_', 'progress_', 'before_', 'after_', 'label_', 'outline_', 'obj_', 'page_num_'];
	const deleteTargets = ['PROGRESS', 'PROGRESS_BG', 'MAIN_TITLE'];

	for (let i = 1; i < slides.length; i++) { // Skip first slide
		const shapes = slides[i].getShapes();
		for (const shape of shapes) {
			const id = shape.getObjectId();
			const shouldDelete = deletePatterns.some(p => id.startsWith(p)) ||
				(shape.getTitle && deleteTargets.includes(shape.getTitle()));
			
			if (shouldDelete) {
				requests.push({ deleteObject: { objectId: id } });
			}
		}
	}
}

/**
 * Convert hex color to RGB format for Google Slides API
 */
function hexToRgb(hex) {
	const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
	return m ? {
		red: parseInt(m[1], 16) / 255,
		green: parseInt(m[2], 16) / 255,
		blue: parseInt(m[3], 16) / 255
	} : { red: 0, green: 0, blue: 0 };
}