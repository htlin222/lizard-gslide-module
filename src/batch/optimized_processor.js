// Optimized batch processor for Google Slides
/**
 * High-performance batch processor that eliminates API call bottlenecks
 * by caching slide properties and minimizing client-side operations
 */

/**
 * Optimized batch processor with caching and minimal API calls
 */
function runOptimizedRequestProcessors(...processors) {
	const presentation = SlidesApp.getActivePresentation();
	const presentationId = presentation.getId();
	const slides = presentation.getSlides();
	const requests = [];

	// 🚀 OPTIMIZATION 1: Cache expensive slide properties once
	const slideCache = createSlideCache(presentation, slides);

	// 🚀 OPTIMIZATION 2: Pre-analyze sections once for all processors
	const sectionsCache = getSectionHeadersOptimized(slides);

	// 🚀 OPTIMIZATION 3: Batch all deletions first, then create
	batchDeleteOldElements(slides, requests);

	// Execute all processors with cached data
	processors.forEach((fn) => fn(slides, requests, slideCache, sectionsCache));

	// Single batch update for all operations
	if (requests.length) {
		Slides.Presentations.batchUpdate({ requests }, presentationId);
	}
}

/**
 * Create cached slide properties to avoid repeated API calls
 */
function createSlideCache(presentation, slides) {
	const slideWidth = presentation.getPageWidth();
	const slideHeight = presentation.getPageHeight();
	
	// Cache slide IDs and types to minimize API calls
	const slideData = slides.map(slide => ({
		id: slide.getObjectId(),
		layoutName: slide.getLayout().getLayoutName(),
		slide: slide // Keep reference for necessary operations
	}));

	return {
		width: slideWidth,
		height: slideHeight,
		totalSlides: slides.length,
		slideData: slideData,
		// Commonly used calculations
		progressBarY: slideHeight - progressBarHeight,
		maxProgressWidth: slideWidth,
		rightFooterX: slideWidth,
		centerX: slideWidth / 2
	};
}

/**
 * Optimized section header detection - single pass through slides
 */
function getSectionHeadersOptimized(slides) {
	const sections = [];
	
	for (let i = 0; i < slides.length; i++) {
		const slide = slides[i];
		const slideData = slide.getLayout().getLayoutName();
		
		if (slideData === 'SECTION_HEADER') {
			const title = getFirstTextboxTextOptimized(slide);
			if (title) {
				sections.push({
					title,
					index: i,
					slideId: slide.getObjectId()
				});
			}
		}
	}
	
	return sections;
}

/**
 * Optimized text extraction with early exit
 */
function getFirstTextboxTextOptimized(slide) {
	const shapes = slide.getShapes();
	for (const shape of shapes) {
		if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
			const text = shape.getText().asString().trim();
			if (text) return text;
		}
	}
	return '';
}

/**
 * 🚀 OPTIMIZATION 4: Batch delete all old elements first
 * This eliminates the expensive client-side shape iteration per function
 */
function batchDeleteOldElements(slides, requests) {
	const deletePatterns = [
		'tab_', 'tab_bg_', 'tab_line_', 'page_num_',  // Tab list elements
		'progress_', 'progress_bg_',                    // Progress bars
		'before_', 'after_', 'label_', 'outline_',     // Section boxes
		'obj_'                                          // Footer elements
	];

	// Single pass through all slides and shapes
	slides.forEach((slide, idx) => {
		if (idx === 0) return; // Skip first slide

		const shapes = slide.getShapes();
		shapes.forEach(shape => {
			const id = shape.getObjectId();
			const shouldDelete = deletePatterns.some(pattern => id.startsWith(pattern)) ||
				(shape.getTitle && (shape.getTitle() === 'PROGRESS' || 
					shape.getTitle() === 'PROGRESS_BG' || 
					shape.getTitle() === 'MAIN_TITLE'));

			if (shouldDelete) {
				requests.push({
					deleteObject: { objectId: id }
				});
			}
		});
	});
}