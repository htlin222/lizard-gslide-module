// 🚀 OPTIMIZED INDEX GENERATOR - Ultra-fast batch processing
/**
 * Optimized index slide generation using batch requests
 * Performance: ~100 individual calls → ~4-8 batch requests per item in 1 API call
 * Expected: 3-5x faster for typical presentations
 */
function generateIndexSlide() {
	const presentation = SlidesApp.getActivePresentation();
	const presentationId = presentation.getId();
	let slides = presentation.getSlides();

	Logger.log("🚀 Starting optimized index generation...");
	const startTime = Date.now();

	// Step 0: Batch delete old index slides
	const deleteRequests = [];
	for (let i = slides.length - 1; i >= 0; i--) {
		const slide = slides[i];
		const layout = slide.getLayout();
		const layoutName = layout ? layout.getLayoutName() : "";

		if (layoutName === "TITLE_ONLY") {
			const shapes = slide.getShapes();
			if (shapes.length > 0) {
				const titleText = shapes[0].getText().asString().trim();
				if (titleText === "Index") {
					deleteRequests.push({
						deleteObject: { objectId: slide.getObjectId() },
					});
				}
			}
		}
	}

	// Execute deletion batch if needed
	if (deleteRequests.length > 0) {
		Logger.log(`Deleting ${deleteRequests.length} old index slides...`);
		Slides.Presentations.batchUpdate(
			{ requests: deleteRequests },
			presentationId,
		);
		slides = presentation.getSlides(); // Refresh slides list
	}

	// Step 1: Collect index items efficiently
	const indexItems = generateIndexItemsUltra(slides);

	if (indexItems.length === 0) {
		Logger.log("No slides found for index generation");
		return;
	}

	// Step 2: Create new index slide
	const newSlide = presentation.appendSlide(
		SlidesApp.PredefinedLayout.TITLE_ONLY,
	);
	const indexSlideId = newSlide.getObjectId();

	// Update title immediately (single operation)
	newSlide.getShapes()[0].getText().setText("Index");

	// Step 3: Generate all index items in single batch
	generateIndexItemsBatch(indexItems, indexSlideId, presentationId);

	const elapsed = Date.now() - startTime;
	Logger.log(
		`✅ Ultra-optimized index created in ${elapsed}ms with ${indexItems.length} items`,
	);
}

/**
 * 🚀 Ultra-efficient index items collection
 */
function generateIndexItemsUltra(slides) {
	const indexItems = [];

	slides.forEach((slide, index) => {
		const layout = slide.getLayout();
		const layoutName = layout ? layout.getLayoutName() : "";
		if (layoutName === "SECTION_HEADER") return;

		// Quick title extraction
		const titleText = extractSlideTitle(slide);

		indexItems.push({
			text: `${titleText || "[No title]"}, p${index + 1}`,
			slideId: slide.getObjectId(), // Store ID instead of slide object
			index: index,
		});
	});

	return indexItems;
}

/**
 * Extract slide title efficiently
 */
function extractSlideTitle(slide) {
	const pageElements = slide.getPageElements();

	for (const element of pageElements) {
		if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const shape = element.asShape();
			const text = shape.getText().asString().trim();
			if (text.length > 0) {
				return text;
			}
		}
	}
	return "";
}

/**
 * 🚀 Ultra-efficient batch creation of index items
 * Consolidates ~100 individual API calls into 1 batch call
 */
function generateIndexItemsBatch(indexItems, indexSlideId, presentationId) {
	const requests = [];

	// Layout configuration
	const CONFIG = {
		initialLeft: 40,
		initialTop: 90,
		width: 230,
		height: 20,
		spacing: 0,
		itemsPerColumn: 12,
		fontSize: 9,
	};

	// Initialize cache for optimization
	const cache = initializeUltraCache();

	Logger.log(`Generating ${indexItems.length} index items in batch...`);

	indexItems.forEach((item, i) => {
		const column = Math.floor(i / CONFIG.itemsPerColumn);
		const row = i % CONFIG.itemsPerColumn;
		const left = CONFIG.initialLeft + column * CONFIG.width;
		const top = CONFIG.initialTop + row * (CONFIG.height + CONFIG.spacing);

		createIndexItemUltra(
			item,
			i,
			left,
			top,
			CONFIG,
			indexSlideId,
			requests,
			cache,
		);
	});

	// Single mega batch update
	if (requests.length > 0) {
		Logger.log(
			`Executing ultra batch: ${requests.length} operations in 1 API call`,
		);
		const batchStart = Date.now();
		Slides.Presentations.batchUpdate({ requests }, presentationId);
		Logger.log(`Batch completed in ${Date.now() - batchStart}ms`);
	}
}

/**
 * Create single index item with optimal batch requests
 * Consolidates 5+ individual calls into 4 batch requests
 */
function createIndexItemUltra(
	item,
	index,
	left,
	top,
	config,
	slideId,
	requests,
	cache,
) {
	const shapeId = `index_${index}_${getNextGuid()}`;

	// 1. Create text box shape
	requests.push({
		createShape: {
			objectId: shapeId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: slideId,
				size: {
					width: { magnitude: config.width, unit: "PT" },
					height: { magnitude: config.height, unit: "PT" },
				},
				transform: {
					...cache.transforms.identity,
					translateX: left,
					translateY: top,
				},
			},
		},
	});

	// 2. Insert text content
	requests.push({
		insertText: {
			objectId: shapeId,
			text: item.text,
		},
	});

	// 3. Apply text styling (consolidated)
	requests.push({
		updateTextStyle: {
			objectId: shapeId,
			textRange: { type: "ALL" },
			style: {
				fontSize: { magnitude: config.fontSize, unit: "PT" },
				fontFamily: main_font_family || "Arial",
				foregroundColor: { opaqueColor: { rgbColor: cache.colors.main } },
				link: { pageObjectId: item.slideId }, // Add link in text style
			},
			fields: "fontSize,fontFamily,foregroundColor,link",
		},
	});

	// 4. Apply shape properties (alignment)
	requests.push({
		updateShapeProperties: {
			objectId: shapeId,
			shapeProperties: {
				contentAlignment: "MIDDLE",
			},
			fields: "contentAlignment",
		},
	});
}
