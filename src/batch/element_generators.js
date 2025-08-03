// 🚀 ELEMENT GENERATORS MODULE - Optimized slide element creation
/**
 * High-performance generators for slide elements
 * - Progress bars
 * - Page numbers  
 * - Title footnotes
 * - Tab navigation
 * - Section elements
 */

/**
 * Generate all elements for a single slide (consolidated)
 */
function generateSlideElementsUltra(slideId, slideData, slideIndex, slideCache, sectionsCache, currentSectionIdx, requests, cache) {
	// 1. Progress bar (2 consolidated requests instead of 6)
	addProgressBarUltra(slideId, slideIndex, slideCache, requests, cache);
	
	// 2. Page number (1 consolidated request instead of 4)
	addPageNumberUltra(slideId, slideIndex, slideCache, requests, cache);
	
	// 3. Title footnote (1 consolidated request instead of 7)
	addTitleFootnoteUltra(slideId, slideCache, requests, cache);
	
	// 4. Tab navigation (only for non-section slides)
	if (slideData.layoutName !== "SECTION_HEADER" && sectionsCache.length > 0) {
		addTabNavigationUltra(slideId, sectionsCache, currentSectionIdx, requests, cache);
	}
}

/**
 * Ultra-efficient progress bar (2 requests instead of 6)
 */
function addProgressBarUltra(slideId, slideIndex, slideCache, requests, cache) {
	const progressRatio = slideIndex / (slideCache.totalSlides - 1);
	const barWidth = slideCache.width * progressRatio;
	const bgId = `progress_bg_${slideId}_${getNextGuid()}`;
	const progId = `progress_${slideId}_${getNextGuid()}`;
	
	// Background + Progress bar in optimized requests
	requests.push(
		{
			createShape: {
				objectId: bgId, shapeType: 'RECTANGLE',
				elementProperties: {
					pageObjectId: slideId,
					size: { height: slideCache.sizes.progressBar.height, width: { magnitude: slideCache.width, unit: 'PT' } },
					transform: { ...cache.transforms.identity, translateX: 0, translateY: slideCache.progressBarY }
				}
			}
		},
		{
			updateShapeProperties: {
				objectId: bgId,
				shapeProperties: {
					shapeBackgroundFill: { solidFill: { color: { rgbColor: cache.colors.gray } } },
					outline: { weight: { magnitude: 0.1, unit: 'PT' }, outlineFill: { solidFill: { color: { rgbColor: cache.colors.gray } } } }
				},
				fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
			}
		},
		{
			createShape: {
				objectId: progId, shapeType: 'RECTANGLE',
				elementProperties: {
					pageObjectId: slideId,
					size: { height: slideCache.sizes.progressBar.height, width: { magnitude: barWidth, unit: 'PT' } },
					transform: { ...cache.transforms.identity, translateX: 0, translateY: slideCache.progressBarY }
				}
			}
		},
		{
			updateShapeProperties: {
				objectId: progId,
				shapeProperties: {
					shapeBackgroundFill: { solidFill: { color: { rgbColor: cache.colors.main } } },
					outline: { weight: { magnitude: 0.1, unit: 'PT' }, outlineFill: { solidFill: { color: { rgbColor: cache.colors.main } } } }
				},
				fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
			}
		}
	);
}

/**
 * Ultra-efficient page number (1 consolidated request instead of 4)
 */
function addPageNumberUltra(slideId, slideIndex, slideCache, requests, cache) {
	const pageId = `page_num_${slideId}_${getNextGuid()}`;
	
	requests.push(
		{
			createShape: {
				objectId: pageId, shapeType: 'TEXT_BOX',
				elementProperties: {
					pageObjectId: slideId,
					size: slideCache.sizes.pageNum,
					transform: { ...cache.transforms.identity, translateX: 650, translateY: 370 }
				}
			}
		},
		{ insertText: { objectId: pageId, text: `${slideIndex + 1} / ${slideCache.totalSlides}` } },
		{
			updateTextStyle: {
				objectId: pageId, textRange: { type: 'ALL' },
				style: {
					bold: true, fontFamily: main_font_family, fontSize: { magnitude: 12, unit: 'PT' },
					foregroundColor: { opaqueColor: { rgbColor: cache.colors.inactive } }
				},
				fields: 'bold,fontFamily,fontSize,foregroundColor'
			}
		},
		{
			updateParagraphStyle: {
				objectId: pageId, textRange: { type: 'ALL' },
				style: { alignment: 'CENTER' }, fields: 'alignment'
			}
		}
	);
}

/**
 * Ultra-efficient title footnote (1 consolidated request instead of 7)
 */
function addTitleFootnoteUltra(slideId, slideCache, requests, cache) {
	// Get main title once and cache it
	if (!cache.mainTitle) {
		cache.mainTitle = getMainTitleFromFirstSlide(slideCache.slideData[0].slide || 
			SlidesApp.getActivePresentation().getSlides()[0]);
	}
	if (!cache.mainTitle) return;

	const footnoteId = `obj_${slideId}_${Date.now().toString(36)}_${getNextGuid()}`;
	const boxWidth = 360;
	const boxY = (slideCache.height - boxWidth) / 2;

	requests.push(
		{
			createShape: {
				objectId: footnoteId, shapeType: 'TEXT_BOX',
				elementProperties: {
					pageObjectId: slideId,
					size: { width: { magnitude: boxWidth, unit: 'PT' }, height: { magnitude: 30, unit: 'PT' } },
					transform: {
						...cache.transforms.rotation90,
						translateX: slideCache.width, translateY: boxY
					}
				}
			}
		},
		{ insertText: { objectId: footnoteId, text: cache.mainTitle } },
		{
			updateTextStyle: {
				objectId: footnoteId, textRange: { type: 'ALL' },
				style: {
					foregroundColor: { opaqueColor: { rgbColor: cache.colors.inactive } },
					fontSize: { magnitude: 10, unit: 'PT' }, fontFamily: main_font_family,
					underline: false, link: { pageObjectId: slideCache.slideData[0].id }
				},
				fields: 'foregroundColor,fontSize,fontFamily,underline,link'
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
				objectId: footnoteId, shapeProperties: { contentAlignment: 'MIDDLE' },
				fields: 'contentAlignment'
			}
		}
	);
}

/**
 * Ultra-efficient tab navigation
 */
function addTabNavigationUltra(slideId, sections, currentSection, requests, cache) {
	// Pre-calculate tab layout
	const estCharW = 8 * 0.75; // fontSize * 0.75
	const widths = sections.map(sec => Math.max(sec.title.length * estCharW, 50) + 5); // +5pt buffer
	const totalWidth = widths.reduce((a, b) => a + b, 0);
	const xStart = Math.max((720 - totalWidth) / 2, 0);
	
	// Background bar
	const bgId = `tab_bg_${slideId}_${getNextGuid()}`;
	requests.push(
		{
			createShape: {
				objectId: bgId, shapeType: 'RECTANGLE',
				elementProperties: {
					pageObjectId: slideId,
					size: { height: { magnitude: 14, unit: 'PT' }, width: { magnitude: 720, unit: 'PT' } },
					transform: { ...cache.transforms.identity, translateX: 0, translateY: 0 }
				}
			}
		},
		{
			updateShapeProperties: {
				objectId: bgId,
				shapeProperties: { 
					shapeBackgroundFill: { solidFill: { color: { rgbColor: cache.colors.white } } },
					outline: {
						weight: { magnitude: 0.1, unit: 'PT' },
						outlineFill: { solidFill: { color: { rgbColor: cache.colors.white } } }
					}
				},
				fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
			}
		}
	);

	// Individual tabs
	let xPos = xStart;
	sections.forEach((sec, idx) => {
		const isActive = idx === currentSection;
		const tabId = `tab_${slideId}_${getNextGuid()}`;
		
		requests.push(
			{
				createShape: {
					objectId: tabId, shapeType: 'TEXT_BOX',
					elementProperties: {
						pageObjectId: slideId,
						size: { height: { magnitude: 14, unit: 'PT' }, width: { magnitude: widths[idx], unit: 'PT' } },
						transform: { ...cache.transforms.identity, translateX: xPos, translateY: 0 }
					}
				}
			},
			{ insertText: { objectId: tabId, text: sec.title } },
			{
				updateShapeProperties: {
					objectId: tabId,
					shapeProperties: {
						shapeBackgroundFill: { solidFill: { color: { rgbColor: isActive ? cache.colors.main : cache.colors.white } } },
						contentAlignment: 'MIDDLE'
					},
					fields: 'shapeBackgroundFill.solidFill.color,contentAlignment'
				}
			},
			{
				updateTextStyle: {
					objectId: tabId, textRange: { type: 'ALL' },
					style: {
						bold: true, fontFamily: main_font_family, fontSize: { magnitude: 8, unit: 'PT' },
						foregroundColor: { opaqueColor: { rgbColor: isActive ? cache.colors.white : cache.colors.inactive } },
						underline: false, link: { pageObjectId: sec.slideId }
					},
					fields: 'bold,fontFamily,fontSize,foregroundColor,underline,link'
				}
			},
			{
				updateParagraphStyle: {
					objectId: tabId, textRange: { type: 'ALL' },
					style: { alignment: 'CENTER' }, fields: 'alignment'
				}
			}
		);
		xPos += widths[idx];
	});

	// Bottom line
	const lineId = `tab_line_${slideId}_${getNextGuid()}`;
	requests.push(
		{
			createLine: {
				objectId: lineId, lineCategory: 'STRAIGHT',
				elementProperties: {
					pageObjectId: slideId,
					size: { height: { magnitude: 0, unit: 'PT' }, width: { magnitude: 720, unit: 'PT' } },
					transform: { ...cache.transforms.identity, translateX: 0, translateY: 14 }
				}
			}
		},
		{
			updateLineProperties: {
				objectId: lineId,
				lineProperties: { lineFill: { solidFill: { color: { rgbColor: cache.colors.main } } } },
				fields: 'lineFill.solidFill.color'
			}
		}
	);
}