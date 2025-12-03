// 🚀 SECTION ELEMENTS MODULE - Section-specific element generation
/**
 * Specialized generators for section slide elements
 * - Before/after section boxes
 * - Section labels
 * - Outline generation
 */

/**
 * Add section-specific elements (before/after boxes, labels, outline)
 */
function addSectionElementsUltra(slides, sectionsCache, requests, cache) {
	if (!sectionsCache.length) return;

	const BOX_CONFIG = {
		x: (720 - 600) / 2,
		yBefore: 30,
		yAfter: 230,
		width: 600,
		boxHeight: 150,
		fontSize: 20,
		textColor: cache.colors.inactive,
	};

	sectionsCache.forEach((sec, idx) => {
		const slideId = sec.slideId;
		const beforeTitles = sectionsCache.slice(0, idx).map((s) => s.title);
		const afterTitles = sectionsCache.slice(idx + 1).map((s) => s.title);

		// Before titles box
		if (beforeTitles.length) {
			addSectionBox(
				slideId,
				beforeTitles,
				BOX_CONFIG,
				"before",
				"BOTTOM",
				requests,
				cache,
			);
		}

		// After titles box
		if (afterTitles.length) {
			addSectionBox(
				slideId,
				afterTitles,
				BOX_CONFIG,
				"after",
				"TOP",
				requests,
				cache,
			);
		}

		// Section label
		addSectionLabel(slideId, idx + 1, requests, cache);
	});

	// Add outline to second slide if applicable
	addOutlineToSecondSlide(slides, sectionsCache, requests, cache);
}

/**
 * Add a section box (before or after)
 */
function addSectionBox(
	slideId,
	titles,
	config,
	type,
	alignment,
	requests,
	cache,
) {
	const boxId = `${type}_${slideId}_${getNextGuid()}`;
	const yPos = type === "before" ? config.yBefore : config.yAfter;

	requests.push(
		{
			createShape: {
				objectId: boxId,
				shapeType: "TEXT_BOX",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						width: { magnitude: config.width, unit: "PT" },
						height: { magnitude: config.boxHeight, unit: "PT" },
					},
					transform: {
						...cache.transforms.identity,
						translateX: config.x,
						translateY: yPos,
					},
				},
			},
		},
		{ insertText: { objectId: boxId, text: titles.join("\n") } },
		{
			updateShapeProperties: {
				objectId: boxId,
				shapeProperties: { contentAlignment: alignment },
				fields: "contentAlignment",
			},
		},
		{
			updateTextStyle: {
				objectId: boxId,
				textRange: { type: "ALL" },
				style: {
					fontSize: { magnitude: config.fontSize, unit: "PT" },
					fontFamily: main_font_family,
					foregroundColor: { opaqueColor: { rgbColor: config.textColor } },
					bold: false,
				},
				fields: "fontSize,fontFamily,foregroundColor,bold",
			},
		},
		{
			updateParagraphStyle: {
				objectId: boxId,
				textRange: { type: "ALL" },
				style: { alignment: "CENTER" },
				fields: "alignment",
			},
		},
	);
}

/**
 * Add section label
 */
function addSectionLabel(slideId, sectionNumber, requests, cache) {
	const labelId = `label_${slideId}_${getNextGuid()}`;

	requests.push(
		{
			createShape: {
				objectId: labelId,
				shapeType: "TEXT_BOX",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						width: { magnitude: 80, unit: "PT" },
						height: { magnitude: 25, unit: "PT" },
					},
					transform: {
						...cache.transforms.identity,
						translateX: 50,
						translateY: 50,
					},
				},
			},
		},
		{ insertText: { objectId: labelId, text: `Section: ${sectionNumber}` } },
		{
			updateShapeProperties: {
				objectId: labelId,
				shapeProperties: {
					contentAlignment: "MIDDLE",
					shapeBackgroundFill: {
						solidFill: { color: { rgbColor: cache.colors.main } },
					},
				},
				fields: "contentAlignment,shapeBackgroundFill.solidFill.color",
			},
		},
		{
			updateTextStyle: {
				objectId: labelId,
				textRange: { type: "ALL" },
				style: {
					fontSize: { magnitude: label_font_size, unit: "PT" },
					fontFamily: main_font_family,
					foregroundColor: { opaqueColor: { rgbColor: cache.colors.white } },
					bold: true,
				},
				fields: "fontSize,fontFamily,foregroundColor,bold",
			},
		},
		{
			updateParagraphStyle: {
				objectId: labelId,
				textRange: { type: "ALL" },
				style: { alignment: "CENTER" },
				fields: "alignment",
			},
		},
	);
}

/**
 * Add outline to second slide if it has title "Outline"
 */
function addOutlineToSecondSlide(slides, sectionsCache, requests, cache) {
	const secondSlide = slides[1];
	if (!secondSlide) return;

	const title = getOutlineSlideTitle(secondSlide);

	if (title === "Outline") {
		const outlineTitles = sectionsCache.map((s) => s.title);
		if (outlineTitles.length) {
			const outlineId = `outline_${secondSlide.getObjectId()}_${getNextGuid()}`;
			requests.push(
				{
					createShape: {
						objectId: outlineId,
						shapeType: "TEXT_BOX",
						elementProperties: {
							pageObjectId: secondSlide.getObjectId(),
							size: {
								width: { magnitude: 400, unit: "PT" },
								height: { magnitude: 300, unit: "PT" },
							},
							transform: {
								...cache.transforms.identity,
								translateX: 280,
								translateY: 51,
							},
						},
					},
				},
				{ insertText: { objectId: outlineId, text: outlineTitles.join("\n") } },
				{
					updateShapeProperties: {
						objectId: outlineId,
						shapeProperties: { contentAlignment: "MIDDLE" },
						fields: "contentAlignment",
					},
				},
				{
					updateTextStyle: {
						objectId: outlineId,
						textRange: { type: "ALL" },
						style: {
							fontSize: { magnitude: 28, unit: "PT" },
							fontFamily: main_font_family,
							foregroundColor: { opaqueColor: { rgbColor: cache.colors.main } },
							bold: false,
						},
						fields: "fontSize,fontFamily,foregroundColor,bold",
					},
				},
				{
					createParagraphBullets: {
						objectId: outlineId,
						textRange: { type: "ALL" },
						bulletPreset: "BULLET_DISC_CIRCLE_SQUARE",
					},
				},
			);
		}
	}
}
