// 🚀 MEGA BATCH PROCESSOR - Single API call for all operations
/**
 * ULTIMATE OPTIMIZATION: Combines all 4 functions into 1 single API call
 * Performance: 4 API calls → 1 API call = ~75% faster than optimized individual calls
 * Expected: 3-5s → 1-2s for 20-slide presentation
 */

function runAllFunctionsMegaBatch() {
	const presentation = SlidesApp.getActivePresentation();
	const presentationId = presentation.getId();
	const slides = presentation.getSlides();
	const requests = [];

	// 🚀 STEP 1: Cache all expensive operations once
	const slideCache = createSlideCache(presentation, slides);
	const sectionsCache = getSectionHeadersOptimized(slides);

	// 🚀 STEP 2: Batch delete all old elements first
	batchDeleteOldElements(slides, requests);

	// 🚀 STEP 3: Add all operations to single requests array
	addProgressBarRequests(slides, requests, slideCache);
	addTabListRequests(slides, requests, slideCache, sectionsCache);
	addTitleFootnoteRequests(slides, requests, slideCache);
	addSectionBoxRequests(slides, requests, slideCache, sectionsCache);

	// 🚀 STEP 4: Single mega batch update for ALL operations
	if (requests.length) {
		Logger.log(`Mega batch: Processing ${requests.length} operations in 1 API call`);
		Slides.Presentations.batchUpdate({ requests }, presentationId);
		Logger.log(`Mega batch: Completed successfully`);
	}

	// Update date (separate small operation)
	updateDateInFirstSlide();
}

/**
 * Add progress bar requests to the batch
 */
function addProgressBarRequests(slides, requests, slideCache) {
	const totalSlides = slideCache.totalSlides;
	const maxWidth = slideCache.maxProgressWidth;
	const height = progressBarHeight;
	const yPosition = slideCache.progressBarY;
	const grayBackgroundColor = '#E0E0E0';

	for (let i = 1; i < totalSlides; i++) {
		const slideId = slideCache.slideData[i].id;
		const progressRatio = i / (totalSlides - 1);
		const barWidth = maxWidth * progressRatio;
		const progressId = `progress_${slideId}_${newGuid()}`;
		const backgroundId = `progress_bg_${slideId}_${newGuid()}`;

		// Background bar
		requests.push(
			{
				createShape: {
					objectId: backgroundId,
					shapeType: 'RECTANGLE',
					elementProperties: {
						pageObjectId: slideId,
						size: {
							height: { magnitude: height, unit: 'PT' },
							width: { magnitude: maxWidth, unit: 'PT' }
						},
						transform: {
							scaleX: 1, scaleY: 1, translateX: 0, translateY: yPosition, unit: 'PT'
						}
					}
				}
			},
			{
				updateShapeProperties: {
					objectId: backgroundId,
					shapeProperties: {
						shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(grayBackgroundColor) } } },
						outline: {
							weight: { magnitude: 0.1, unit: 'PT' },
							outlineFill: { solidFill: { color: { rgbColor: hexToRgb(grayBackgroundColor) } } }
						}
					},
					fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
				}
			},
			{ updatePageElementAltText: { objectId: backgroundId, title: 'PROGRESS_BG' } }
		);

		// Progress bar
		requests.push(
			{
				createShape: {
					objectId: progressId,
					shapeType: 'RECTANGLE',
					elementProperties: {
						pageObjectId: slideId,
						size: {
							height: { magnitude: height, unit: 'PT' },
							width: { magnitude: barWidth, unit: 'PT' }
						},
						transform: {
							scaleX: 1, scaleY: 1, translateX: 0, translateY: yPosition, unit: 'PT'
						}
					}
				}
			},
			{
				updateShapeProperties: {
					objectId: progressId,
					shapeProperties: {
						shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(main_color) } } },
						outline: {
							weight: { magnitude: 0.1, unit: 'PT' },
							outlineFill: { solidFill: { color: { rgbColor: hexToRgb(main_color) } } }
						}
					},
					fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
				}
			},
			{ updatePageElementAltText: { objectId: progressId, title: 'PROGRESS' } }
		);
	}
}

/**
 * Add tab list requests to the batch
 */
function addTabListRequests(slides, requests, slideCache, sectionsCache) {
	if (sectionsCache.length === 0) return;

	const CONFIG = {
		totalWidth: 720, height: 14, y: 0, fontSize: 8, padding: 0, spacing: 0,
		mainColor: main_color, mainFont: main_font_family, bgColor: "#FFFFFF",
		inactiveTextColor: "#888888", minWidth: 50
	};

	let currentSectionIdx = -1;
	const totalPages = slideCache.totalSlides;

	slideCache.slideData.forEach((slideData, idx) => {
		if (idx === 0) return;
		const slideId = slideData.id;

		if (currentSectionIdx + 1 < sectionsCache.length && idx >= sectionsCache[currentSectionIdx + 1].index) {
			currentSectionIdx++;
		}

		// Add page number
		const pageNumId = `page_num_${slideId}_${newGuid()}`;
		requests.push(
			{
				createShape: {
					objectId: pageNumId, shapeType: "TEXT_BOX",
					elementProperties: {
						pageObjectId: slideId,
						size: { height: { magnitude: 30, unit: "PT" }, width: { magnitude: 70, unit: "PT" } },
						transform: { translateX: 650, translateY: 370, scaleX: 1, scaleY: 1, unit: "PT" }
					}
				}
			},
			{ insertText: { objectId: pageNumId, text: `${idx + 1} / ${totalPages}` } },
			{
				updateTextStyle: {
					objectId: pageNumId, textRange: { type: "ALL" },
					style: {
						bold: true, fontFamily: CONFIG.mainFont, fontSize: { magnitude: 12, unit: "PT" },
						foregroundColor: { opaqueColor: { rgbColor: hexToRgb(CONFIG.inactiveTextColor) } }
					},
					fields: "bold,fontFamily,fontSize,foregroundColor"
				}
			},
			{
				updateParagraphStyle: {
					objectId: pageNumId, textRange: { type: "ALL" },
					style: { alignment: "CENTER" }, fields: "alignment"
				}
			}
		);

		// Skip section headers for tabs
		if (slideData.layoutName === "SECTION_HEADER") return;

		const currentSection = currentSectionIdx >= 0 ? currentSectionIdx : -1;
		addTabListToRequests(slideId, requests, sectionsCache, currentSection, CONFIG);
	});
}

function addTabListToRequests(slideId, requests, sections, currentSection, config) {
	const estCharW = config.fontSize * 0.75;
	const widths = sections.map(sec => Math.max(sec.title.length * estCharW + config.padding, config.minWidth));
	const totalTabsWidth = widths.reduce((a, b) => a + b, 0) + config.spacing * (widths.length - 1);
	const xStart = Math.max((config.totalWidth - totalTabsWidth) / 2, 0);
	let xPos = xStart;

	// Background bar
	const bgId = `tab_bg_${slideId}_${newGuid()}`;
	requests.push(
		{
			createShape: {
				objectId: bgId, shapeType: "RECTANGLE",
				elementProperties: {
					pageObjectId: slideId,
					size: { height: { magnitude: config.height, unit: "PT" }, width: { magnitude: config.totalWidth, unit: "PT" } },
					transform: { translateX: 0, translateY: config.y, scaleX: 1, scaleY: 1, unit: "PT" }
				}
			}
		},
		{
			updateShapeProperties: {
				objectId: bgId,
				shapeProperties: {
					shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(config.bgColor) } } },
					outline: {
						weight: { magnitude: 0.1, unit: "PT" },
						outlineFill: { solidFill: { color: { rgbColor: hexToRgb(config.bgColor) } } }
					}
				},
				fields: "shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color"
			}
		}
	);

	// Individual tabs
	sections.forEach((sec, idx) => {
		const isActive = idx === currentSection;
		const tabId = `tab_${slideId}_${newGuid()}`;
		const textColor = isActive ? "#FFFFFF" : config.inactiveTextColor;
		const fillColor = isActive ? config.mainColor : config.bgColor;

		requests.push(
			{
				createShape: {
					objectId: tabId, shapeType: "TEXT_BOX",
					elementProperties: {
						pageObjectId: slideId,
						size: { height: { magnitude: config.height, unit: "PT" }, width: { magnitude: widths[idx], unit: "PT" } },
						transform: { translateX: xPos, translateY: config.y, scaleX: 1, scaleY: 1, unit: "PT" }
					}
				}
			},
			{ insertText: { objectId: tabId, text: sec.title } },
			{
				updateShapeProperties: {
					objectId: tabId,
					shapeProperties: {
						shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(fillColor) } } },
						contentAlignment: "MIDDLE"
					},
					fields: "shapeBackgroundFill.solidFill.color,contentAlignment"
				}
			},
			{
				updateTextStyle: {
					objectId: tabId, textRange: { type: "ALL" },
					style: {
						bold: true, fontFamily: config.mainFont, fontSize: { magnitude: config.fontSize, unit: "PT" },
						foregroundColor: { opaqueColor: { rgbColor: hexToRgb(textColor) } },
						underline: false, link: { pageObjectId: sec.slideId }
					},
					fields: "bold,fontFamily,fontSize,foregroundColor,underline,link"
				}
			},
			{
				updateParagraphStyle: {
					objectId: tabId, textRange: { type: "ALL" },
					style: { alignment: "CENTER" }, fields: "alignment"
				}
			}
		);
		xPos += widths[idx] + config.spacing;
	});

	// Bottom line
	const lineId = `tab_line_${slideId}_${newGuid()}`;
	requests.push(
		{
			createLine: {
				objectId: lineId, lineCategory: "STRAIGHT",
				elementProperties: {
					pageObjectId: slideId,
					size: { height: { magnitude: 0, unit: "PT" }, width: { magnitude: config.totalWidth, unit: "PT" } },
					transform: { translateX: 0, translateY: config.y + config.height, scaleX: 1, scaleY: 1, unit: "PT" }
				}
			}
		},
		{
			updateLineProperties: {
				objectId: lineId,
				lineProperties: { lineFill: { solidFill: { color: { rgbColor: hexToRgb(config.mainColor) } } } },
				fields: "lineFill.solidFill.color"
			}
		}
	);
}

/**
 * Add title footnote requests to the batch
 */
function addTitleFootnoteRequests(slides, requests, slideCache) {
	if (slideCache.totalSlides < 2) return;

	const firstSlideId = slideCache.slideData[0].id;
	const mainTitle = getMainTitleFromFirstSlideOptimized(slides[0]);
	if (!mainTitle) return;

	const slideWidth = slideCache.width;
	const slideHeight = slideCache.height;

	for (let i = 1; i < slideCache.totalSlides; i++) {
		const slideId = slideCache.slideData[i].id;
		const footnoteId = `obj_${slideId}_${Date.now().toString(36)}_${newGuid()}`;
		const boxWidth = 360;
		const boxY = (slideHeight - boxWidth) / 2;

		requests.push(
			{
				createShape: {
					objectId: footnoteId, shapeType: 'TEXT_BOX',
					elementProperties: {
						pageObjectId: slideId,
						size: { width: { magnitude: boxWidth, unit: 'PT' }, height: { magnitude: 30, unit: 'PT' } },
						transform: {
							scaleX: 0, shearX: -1, shearY: 1, scaleY: 0,
							translateX: slideWidth, translateY: boxY, unit: 'PT'
						}
					}
				}
			},
			{ insertText: { objectId: footnoteId, insertionIndex: 0, text: mainTitle } },
			{
				updateTextStyle: {
					objectId: footnoteId, textRange: { type: 'ALL' },
					style: { link: { pageObjectId: firstSlideId } }, fields: 'link'
				}
			},
			{
				updateTextStyle: {
					objectId: footnoteId, textRange: { type: 'ALL' },
					style: {
						foregroundColor: { opaqueColor: { rgbColor: hexToRgb('#888888') } },
						fontSize: { magnitude: 10, unit: 'PT' }, fontFamily: main_font_family, underline: false
					},
					fields: 'foregroundColor,fontSize,fontFamily,underline'
				}
			},
			{
				updateParagraphStyle: {
					objectId: footnoteId, textRange: { type: 'ALL' },
					style: { alignment: 'CENTER' }, fields: 'alignment'
				}
			},
			{
				updateShapeProperties: {
					objectId: footnoteId, shapeProperties: { contentAlignment: 'MIDDLE' }, fields: 'contentAlignment'
				}
			},
			{ updatePageElementAltText: { objectId: footnoteId, title: 'MAIN_TITLE' } }
		);
	}
}

/**
 * Add section box requests to the batch
 */
function addSectionBoxRequests(slides, requests, slideCache, sectionsCache) {
	if (!sectionsCache.length) return;

	const BOX_CONFIG = {
		x: (720 - 600) / 2, yBefore: 30, yAfter: 240, width: 600, boxHeight: 150,
		fontSize: 20, fontFamily: main_font_family, textColor: '#aaaaaa', isBold: false
	};

	sectionsCache.forEach((sec, idx) => {
		const slideId = sec.slideId;
		const beforeTitles = sectionsCache.slice(0, idx).map(s => s.title);
		const afterTitles = sectionsCache.slice(idx + 1).map(s => s.title);

		// Before titles
		if (beforeTitles.length) {
			const beforeId = `before_${slideId}_${newGuid()}`;
			requests.push(
				{
					createShape: {
						objectId: beforeId, shapeType: 'TEXT_BOX',
						elementProperties: {
							pageObjectId: slideId,
							size: { width: { magnitude: BOX_CONFIG.width, unit: 'PT' }, height: { magnitude: BOX_CONFIG.boxHeight, unit: 'PT' } },
							transform: { translateX: BOX_CONFIG.x, translateY: BOX_CONFIG.yBefore, scaleX: 1, scaleY: 1, unit: 'PT' }
						}
					}
				},
				{
					updateShapeProperties: {
						objectId: beforeId, shapeProperties: { contentAlignment: 'BOTTOM' }, fields: 'contentAlignment'
					}
				},
				{ insertText: { objectId: beforeId, text: beforeTitles.join('\n') } },
				{
					updateTextStyle: {
						objectId: beforeId, textRange: { type: 'ALL' },
						style: {
							fontSize: { magnitude: BOX_CONFIG.fontSize, unit: 'PT' }, fontFamily: BOX_CONFIG.fontFamily,
							foregroundColor: { opaqueColor: { rgbColor: hexToRgb(BOX_CONFIG.textColor) } }, bold: BOX_CONFIG.isBold
						},
						fields: 'fontSize,fontFamily,foregroundColor,bold'
					}
				},
				{
					updateParagraphStyle: {
						objectId: beforeId, textRange: { type: 'ALL' }, style: { alignment: 'CENTER' }, fields: 'alignment'
					}
				}
			);
		}

		// After titles
		if (afterTitles.length) {
			const afterId = `after_${slideId}_${newGuid()}`;
			requests.push(
				{
					createShape: {
						objectId: afterId, shapeType: 'TEXT_BOX',
						elementProperties: {
							pageObjectId: slideId,
							size: { width: { magnitude: BOX_CONFIG.width, unit: 'PT' }, height: { magnitude: BOX_CONFIG.boxHeight, unit: 'PT' } },
							transform: { translateX: BOX_CONFIG.x, translateY: BOX_CONFIG.yAfter, scaleX: 1, scaleY: 1, unit: 'PT' }
						}
					}
				},
				{
					updateShapeProperties: {
						objectId: afterId, shapeProperties: { contentAlignment: 'TOP' }, fields: 'contentAlignment'
					}
				},
				{ insertText: { objectId: afterId, text: afterTitles.join('\n') } },
				{
					updateTextStyle: {
						objectId: afterId, textRange: { type: 'ALL' },
						style: {
							fontSize: { magnitude: BOX_CONFIG.fontSize, unit: 'PT' }, fontFamily: BOX_CONFIG.fontFamily,
							foregroundColor: { opaqueColor: { rgbColor: hexToRgb(BOX_CONFIG.textColor) } }, bold: BOX_CONFIG.isBold
						},
						fields: 'fontSize,fontFamily,foregroundColor,bold'
					}
				},
				{
					updateParagraphStyle: {
						objectId: afterId, textRange: { type: 'ALL' }, style: { alignment: 'CENTER' }, fields: 'alignment'
					}
				}
			);
		}

		// Section label
		const labelId = `label_${slideId}_${newGuid()}`;
		requests.push(
			{
				createShape: {
					objectId: labelId, shapeType: 'TEXT_BOX',
					elementProperties: {
						pageObjectId: slideId,
						size: { width: { magnitude: 80, unit: 'PT' }, height: { magnitude: 25, unit: 'PT' } },
						transform: { translateX: 50, translateY: 50, scaleX: 1, scaleY: 1, unit: 'PT' }
					}
				}
			},
			{
				updateShapeProperties: {
					objectId: labelId,
					shapeProperties: {
						contentAlignment: 'MIDDLE',
						shapeBackgroundFill: { solidFill: { color: { rgbColor: hexToRgb(main_color) } } }
					},
					fields: 'contentAlignment,shapeBackgroundFill.solidFill.color'
				}
			},
			{ insertText: { objectId: labelId, text: `Section: ${idx + 1}` } },
			{
				updateTextStyle: {
					objectId: labelId, textRange: { type: 'ALL' },
					style: {
						fontSize: { magnitude: label_font_size, unit: 'PT' }, fontFamily: BOX_CONFIG.fontFamily,
						foregroundColor: { opaqueColor: { rgbColor: hexToRgb('#FFFFFF') } }, bold: true
					},
					fields: 'fontSize,fontFamily,foregroundColor,bold'
				}
			},
			{
				updateParagraphStyle: {
					objectId: labelId, textRange: { type: 'ALL' }, style: { alignment: 'CENTER' }, fields: 'alignment'
				}
			}
		);
	});

	// Add outline to second page
	const secondSlide = slides[1];
	if (secondSlide) {
		let title = '';
		const placeholder = secondSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
		if (placeholder && placeholder.asShape) {
			title = placeholder.asShape().getText().asString().trim();
		} else {
			title = getFirstTextboxTextOptimized(secondSlide);
		}
		if (title === 'Outline') {
			const outlineTitles = sectionsCache.map(s => s.title);
			if (outlineTitles.length) {
				const outlineId = `outline_${secondSlide.getObjectId()}_${newGuid()}`;
				requests.push(
					{
						createShape: {
							objectId: outlineId, shapeType: 'TEXT_BOX',
							elementProperties: {
								pageObjectId: secondSlide.getObjectId(),
								size: { width: { magnitude: 400, unit: 'PT' }, height: { magnitude: 300, unit: 'PT' } },
								transform: { translateX: 280, translateY: 51, scaleX: 1, scaleY: 1, unit: 'PT' }
							}
						}
					},
					{
						updateShapeProperties: {
							objectId: outlineId, shapeProperties: { contentAlignment: 'MIDDLE' }, fields: 'contentAlignment'
						}
					},
					{ insertText: { objectId: outlineId, text: outlineTitles.join('\n') } },
					{
						updateTextStyle: {
							objectId: outlineId, textRange: { type: 'ALL' },
							style: {
								fontSize: { magnitude: 28, unit: 'PT' }, fontFamily: main_font_family,
								foregroundColor: { opaqueColor: { rgbColor: hexToRgb(main_color) } }, bold: false
							},
							fields: 'fontSize,fontFamily,foregroundColor,bold'
						}
					},
					{
						createParagraphBullets: {
							objectId: outlineId, textRange: { type: 'ALL' }, bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE'
						}
					}
				);
			}
		}
	}
}

function getMainTitleFromFirstSlideOptimized(slide) {
	const elements = slide.getPageElements();
	for (let el of elements) {
		if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const txt = el.asShape().getText().asString().trim();
			if (txt) return txt;
		}
	}
	return '';
}

function getFirstTextboxTextOptimized(slide) {
	for (const shape of slide.getShapes()) {
		if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
			const txt = shape.getText().asString().trim();
			if (txt) return txt;
		}
	}
	return '';
}

function newGuid() {
	return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}