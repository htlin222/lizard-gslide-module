// 🚀 SECTION ELEMENTS MODULE - Section-specific element generation
/**
 * Specialized generators for section slide elements
 * - Before/after section boxes
 * - Section labels
 * - Outline generation
 */

/**
 * Add section-specific elements (unified section box with styled lines)
 */
function addSectionElementsUltra(slides, sectionsCache, requests, cache) {
	if (!sectionsCache.length) return;

	const BOX_CONFIG = {
		x: 200,
		width: 500,
		height: 300,
		// Styles for different states
		before: { fontSize: 30, color: cache.colors.inactive, bold: false },
		current: { fontSize: 36, color: cache.colors.main, bold: true },
		after: { fontSize: 30, color: { red: 0, green: 0, blue: 0 }, bold: false },
	};

	// Calculate y position to center vertically (slide height ~405pt)
	const slideHeight = 405;
	BOX_CONFIG.y = (slideHeight - BOX_CONFIG.height) / 2;

	sectionsCache.forEach((sec, idx) => {
		const slideId = sec.slideId;

		// Create unified section box with all titles
		addUnifiedSectionBox(
			slideId,
			sectionsCache,
			idx,
			BOX_CONFIG,
			requests,
			cache,
		);

		// Section label
		addSectionLabel(slideId, idx + 1, requests, cache);
	});

	// Add outline to second slide if applicable
	addOutlineToSecondSlide(slides, sectionsCache, requests, cache);
}

/**
 * Add unified section box with all titles, each line styled differently
 * - Before sections: light gray, 30pt
 * - Current section: main color, bold, 36pt
 * - After sections: black, 32pt
 */
function addUnifiedSectionBox(
	slideId,
	sectionsCache,
	currentIdx,
	config,
	requests,
	cache,
) {
	const boxId = `sections_${slideId}_${getNextGuid()}`;

	// Build numbered list: "1. Title\n2. Title\n..."
	const lines = sectionsCache.map((s, i) => `${i + 1}. ${s.title}`);
	const fullText = lines.join("\n");

	// Create shape and insert text
	requests.push(
		{
			createShape: {
				objectId: boxId,
				shapeType: "TEXT_BOX",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						width: { magnitude: config.width, unit: "PT" },
						height: { magnitude: config.height, unit: "PT" },
					},
					transform: {
						...cache.transforms.identity,
						translateX: config.x,
						translateY: config.y,
					},
				},
			},
		},
		{ insertText: { objectId: boxId, text: fullText } },
		{
			updateShapeProperties: {
				objectId: boxId,
				shapeProperties: {
					contentAlignment: "MIDDLE",
					shapeBackgroundFill: {
						solidFill: { color: { rgbColor: cache.colors.white } },
					},
					outline: {
						outlineFill: {
							solidFill: { color: { rgbColor: cache.colors.white } },
						},
					},
				},
				fields:
					"contentAlignment,shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color",
			},
		},
		{
			updateParagraphStyle: {
				objectId: boxId,
				textRange: { type: "ALL" },
				style: { alignment: "START" },
				fields: "alignment",
			},
		},
	);

	// Calculate text ranges and apply styles for each line
	let charIndex = 0;
	lines.forEach((line, i) => {
		const startIndex = charIndex;
		const endIndex = charIndex + line.length;

		let style;
		if (i < currentIdx) {
			// Before: light gray, 30pt
			style = config.before;
		} else if (i === currentIdx) {
			// Current: main color, bold, 36pt
			style = config.current;
		} else {
			// After: black, 32pt
			style = config.after;
		}

		requests.push({
			updateTextStyle: {
				objectId: boxId,
				textRange: {
					type: "FIXED_RANGE",
					startIndex: startIndex,
					endIndex: endIndex,
				},
				style: {
					fontSize: { magnitude: style.fontSize, unit: "PT" },
					fontFamily: main_font_family,
					foregroundColor: { opaqueColor: { rgbColor: style.color } },
					bold: style.bold,
				},
				fields: "fontSize,fontFamily,foregroundColor,bold",
			},
		});

		// Move to next line (+1 for \n character)
		charIndex = endIndex + 1;
	});
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
