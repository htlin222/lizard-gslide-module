// 🚀 SLIDE UTILITIES MODULE - Slide analysis and manipulation helpers
/**
 * Utility functions for slide analysis and content extraction
 * - Section detection
 * - Title extraction
 * - Element deletion patterns
 */

// Layout display names treated as a "section header" slide. getLayoutName() is
// localized by the editor's UI language and also depends on the deck's import
// source, so match a list rather than one hardcoded string. Extend for your locale.
var LZ_SECTION_LAYOUT_NAMES = [
	"SECTION_HEADER", // Google canonical / Google-exported PPTX
	"Section Header", // python-pptx default template
	"區段標題", // zh-Hant
	"区段标题", // zh-Hans
	"セクションの見出し", // ja
	"섹션 헤더", // ko
];

/**
 * Ultra-efficient section header detection
 */
function getSectionHeadersUltra(slides) {
	const sections = [];
	for (let i = 0; i < slides.length; i++) {
		const slide = slides[i];
		const shapes = slide.getShapes();

		// A slide is a section boundary when EITHER its layout is a known
		// section-header layout OR it carries an LZ-Protocol SECTION marker. The
		// marker is the reliable, language-independent path — layout NAMES vary by
		// UI locale (zh: 「區段標題」) and by import source (python-pptx: "Section
		// Header"), so the list below is a best-effort fallback for decks without a
		// marker. Add your locale's name here if needed. See LZ-PROTOCOL.md.
		const byLayout =
			LZ_SECTION_LAYOUT_NAMES.indexOf(slide.getLayout().getLayoutName()) >= 0;
		let markerTitle = "";
		for (const shape of shapes) {
			if (lzIsSectionMarker(shape)) {
				// Prefer the marker's clean JSON title, else the shape's text.
				markerTitle =
					lzMarkerTitle(shape) ||
					(shape.getText ? shape.getText().asString().trim() : "");
				break;
			}
		}
		if (!byLayout && !markerTitle) continue;

		let title = markerTitle;
		if (!title) {
			// Early exit on first text found
			for (const shape of shapes) {
				if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
					const text = shape.getText().asString().trim();
					if (text) {
						title = text;
						break;
					}
				}
			}
		}
		if (title) {
			sections.push({ title: title, index: i, slideId: slide.getObjectId() });
		}
	}
	return sections;
}

/**
 * Extract main title from first slide
 * Selection priority:
 * 1. Title placeholder
 * 2. Largest text (by font size)
 * 3. Text in upper half of slide
 * 4. First non-empty text (fallback)
 */
function getMainTitleFromFirstSlide(slide) {
	// Try to get title placeholder first
	const titlePlaceholder = slide.getPlaceholder(
		SlidesApp.PlaceholderType.TITLE,
	);
	if (titlePlaceholder && titlePlaceholder.asShape) {
		const titleText = titlePlaceholder.asShape().getText().asString().trim();
		if (titleText) return titleText;
	}

	// Collect all text elements with their properties
	const textElements = [];
	const elements = slide.getPageElements();

	for (const el of elements) {
		if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const shape = el.asShape();
			const txt = shape.getText().asString().trim();
			if (txt) {
				// Get position and size info
				const transform = el.getTransform();
				const height = el.getHeight();
				const translateY = transform.getTranslateY();

				// Try to get font size (default to 0 if unable to get)
				let fontSize = 0;
				try {
					const textRange = shape.getText();
					if (textRange && textRange.getTextStyle) {
						const style = textRange.getTextStyle();
						if (style && style.getFontSize) {
							fontSize = style.getFontSize() || 0;
						}
					}
				} catch (e) {
					// Ignore errors when getting font size
				}

				textElements.push({
					text: txt,
					fontSize: fontSize,
					yPosition: translateY,
					height: height,
					element: el,
				});
			}
		}
	}

	if (textElements.length === 0) return "";

	// Sort by criteria:
	// 1. Larger font size first
	// 2. Higher position (smaller Y) first
	// 3. Larger height first
	textElements.sort((a, b) => {
		// Compare font size (larger is better)
		if (a.fontSize !== b.fontSize) {
			return b.fontSize - a.fontSize;
		}
		// Compare Y position (smaller Y = higher on slide = better)
		if (Math.abs(a.yPosition - b.yPosition) > 50) {
			// 50pt threshold
			return a.yPosition - b.yPosition;
		}
		// Compare height (larger is better)
		return b.height - a.height;
	});

	// Return the best candidate
	return textElements[0].text;
}

/**
 * Get title from outline slide (second slide if it exists)
 */
function getOutlineSlideTitle(slide) {
	let title = "";
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
	const deletePatterns = [
		"tab_",
		"progress_",
		"sections_",
		"label_",
		"outline_",
		"obj_",
		"page_num_",
	];
	const deleteTargets = ["PROGRESS", "PROGRESS_BG", "MAIN_TITLE"];

	for (let i = 1; i < slides.length; i++) {
		// Skip first slide
		const shapes = slides[i].getShapes();
		for (const shape of shapes) {
			const id = shape.getObjectId();
			// Enhanced deletion check - also check for malformed IDs (in case newGuid failed)
			const shouldDelete =
				deletePatterns.some((p) => id.startsWith(p)) ||
				(id.includes("page_num") && id.includes("undefined")) || // Handle broken IDs from undefined newGuid
				(shape.getTitle && deleteTargets.includes(shape.getTitle())) ||
				lzIsManaged(shape); // LZ-Protocol: reclaim tag-carrying chrome whose
			// objectId was regenerated by the PPTX→Slides import (see LZ-PROTOCOL.md)

			if (shouldDelete) {
				requests.push({ deleteObject: { objectId: id } });
			}
		}
	}
}

// hexToRgb(hex) now lives in shared/color_utils.js (single source of truth).
// The duplicate definition that used to be here was removed to avoid a global
// name collision; the canonical version keeps the same { red, green, blue }
// 0–1 shape and the same black fallback on malformed input.
