// Tab list module for Google Slides
/**
 * 主流程：先刪除所有舊的 TAB_LIST_* 元件，再批次建立新的頁籤列，並加入分頁
 */
function processTabs(slides) {
	const CONFIG = {
		totalWidth: 720,
		height: 14,
		y: 0,
		fontSize: 8,
		padding: 0,
		spacing: 0,
		mainColor: main_color,
		mainFont: main_font_family,
		bgColor: "#FFFFFF",
		inactiveTextColor: "#888888",
		minWidth: 50, // <-- e.g. 50pt minimum
	};

	const sections = getSectionHeaders(slides);
	if (sections.length === 0) return;

	// 先 client-side 刪除
	slides.forEach((slide, idx) => {
		if (idx === 0) return; // 跳過封面
		deleteOldTabsClient(slide);
	});

	// 再 batchUpdate 建立新的
	const requests = [];
	let currentSectionIdx = -1;
	const totalPages = slides.length;

	slides.forEach((slide, idx) => {
		if (idx === 0) return;
		const slideId = slide.getObjectId();

		// 計算這張 slide 應該是第幾個 section
		if (
			currentSectionIdx + 1 < sections.length &&
			idx >= sections[currentSectionIdx + 1].index
		) {
			currentSectionIdx++;
		}
		appendPageNumberToSlide({
			slideId,
			requests,
			currentPage: idx + 1,
			totalPages,
			config: CONFIG,
		});
		if (slide.getLayout().getLayoutName() === "SECTION_HEADER") return;
		const currentSection = currentSectionIdx >= 0 ? currentSectionIdx : -1;
		// 加入標籤列
		appendTabListToSlide({
			slideId,
			requests,
			sections,
			currentSection,
			config: CONFIG,
		});
		// 加入分頁文字
	});

	if (requests.length) {
		Slides.Presentations.batchUpdate(
			{ requests },
			SlidesApp.getActivePresentation().getId(),
		);
	}
}

/** 讀出所有 SECTION_HEADER 投影片的標題 */
function getSectionHeaders(slides) {
	return slides
		.map((slide, index) => {
			if (slide.getLayout().getLayoutName() === "SECTION_HEADER") {
				const title = getFirstTextboxText(slide);
				if (title) return { title, index, slideId: slide.getObjectId() };
			}
			return null;
		})
		.filter(Boolean);
}

/** 取第一個 textbox 的文字 */
function getFirstTextboxText(slide) {
	return (
		slide
			.getShapes()
			.filter((s) => s.getShapeType() === SlidesApp.ShapeType.TEXT_BOX)
			.map((s) => s.getText().asString().trim())
			.find((t) => t) || ""
	);
}

/** Client-side 直接刪除舊的 TAB_LIST_* shapes/lines，改用 objectId 前綴判斷 */
function deleteOldTabsClient(slide) {
	slide.getShapes().forEach((shape) => {
		const id = shape.getObjectId();
		if (
			id.startsWith("tab_") ||
			id.startsWith("tab_bg_") ||
			id.startsWith("tab_line_") ||
			id.startsWith("page_num_")
		) {
			shape.remove();
		}
	});
}

/** 把一整列標籤的 batchUpdate requests 推到 requests 陣列 */
function appendTabListToSlide({
	slideId,
	requests,
	sections,
	currentSection,
	config,
}) {
	const estCharW = config.fontSize * 0.75;
	const widths = sections.map((sec) =>
		Math.max(sec.title.length * estCharW + config.padding, config.minWidth),
	);
	const totalTabsWidth =
		widths.reduce((a, b) => a + b, 0) + config.spacing * (widths.length - 1);
	const xStart = Math.max((config.totalWidth - totalTabsWidth) / 2, 0);
	let xPos = xStart;

	addBackgroundTabBar(slideId, requests, config);

	sections.forEach((sec, idx) => {
		const isActive = idx === currentSection;
		appendTab({
			slideId,
			requests,
			title: sec.title,
			targetSlideId: sec.slideId,
			xPos,
			width: widths[idx],
			config,
			textColor: isActive ? "#FFFFFF" : config.inactiveTextColor,
			fillColor: isActive ? config.mainColor : config.bgColor,
		});
		xPos += widths[idx] + config.spacing;
	});

	addBottomLine(slideId, requests, config);
}

/** 批次建立頁碼文字 */
function appendPageNumberToSlide({
	slideId,
	requests,
	currentPage,
	totalPages,
	config,
}) {
	const pageNumId = `page_num_${slideId}_${newGuid()}`;
	requests.push(
		// 建立分頁文字框
		{
			createShape: {
				objectId: pageNumId,
				shapeType: "TEXT_BOX",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						height: { magnitude: 30, unit: "PT" },
						width: { magnitude: 60, unit: "PT" },
					},
					transform: {
						translateX: 650,
						translateY: 370,
						scaleX: 1,
						scaleY: 1,
						unit: "PT",
					},
				},
			},
		},
		// 插入分頁文字
		{
			insertText: {
				objectId: pageNumId,
				text: `${currentPage} / ${totalPages}`,
			},
		},
		// 文字樣式 & 對齊
		{
			updateTextStyle: {
				objectId: pageNumId,
				textRange: { type: "ALL" },
				style: {
					bold: true,
					fontFamily: config.mainFont,
					fontSize: { magnitude: 12, unit: "PT" },
					foregroundColor: {
						opaqueColor: { rgbColor: hexToRgb(config.inactiveTextColor) },
					},
				},
				fields: "bold,fontFamily,fontSize,foregroundColor",
			},
		},
		{
			updateParagraphStyle: {
				objectId: pageNumId,
				textRange: { type: "ALL" },
				style: { alignment: "CENTER" },
				fields: "alignment",
			},
		},
	);
}

function appendTab({
	slideId,
	requests,
	title,
	targetSlideId,
	xPos,
	width,
	config,
	textColor,
	fillColor,
}) {
	const tabId = `tab_${slideId}_${newGuid()}`;
	requests.push(
		{
			createShape: {
				objectId: tabId,
				shapeType: "TEXT_BOX",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						height: { magnitude: config.height, unit: "PT" },
						width: { magnitude: width, unit: "PT" },
					},
					transform: {
						translateX: xPos,
						translateY: config.y,
						scaleX: 1,
						scaleY: 1,
						unit: "PT",
					},
				},
			},
		},
		{ insertText: { objectId: tabId, text: title } },
		{
			updateShapeProperties: {
				objectId: tabId,
				shapeProperties: {
					shapeBackgroundFill: solidFill(fillColor),
					contentAlignment: "MIDDLE",
				},
				fields: "shapeBackgroundFill.solidFill.color,contentAlignment",
			},
		},
		{
			updateTextStyle: {
				objectId: tabId,
				textRange: { type: "ALL" },
				style: {
					bold: true,
					fontFamily: config.mainFont,
					fontSize: { magnitude: config.fontSize, unit: "PT" },
					foregroundColor: { opaqueColor: { rgbColor: hexToRgb(textColor) } },
					underline: false,
					link: { pageObjectId: targetSlideId },
				},
				fields: "bold,fontFamily,fontSize,foregroundColor,underline,link",
			},
		},
		{
			updateParagraphStyle: {
				objectId: tabId,
				textRange: { type: "ALL" },
				style: { alignment: "CENTER" },
				fields: "alignment",
			},
		},
	);
}

function addBackgroundTabBar(slideId, requests, config) {
	const bgId = `tab_bg_${slideId}_${newGuid()}`;
	requests.push(
		{
			createShape: {
				objectId: bgId,
				shapeType: "RECTANGLE",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						height: { magnitude: config.height, unit: "PT" },
						width: { magnitude: config.totalWidth, unit: "PT" },
					},
					transform: {
						translateX: 0,
						translateY: config.y,
						scaleX: 1,
						scaleY: 1,
						unit: "PT",
					},
				},
			},
		},
		{
			updateShapeProperties: {
				objectId: bgId,
				shapeProperties: {
					shapeBackgroundFill: solidFill(config.bgColor),
					outline: {
						weight: { magnitude: 0.1, unit: "PT" },
						outlineFill: {
							solidFill: { color: { rgbColor: hexToRgb(config.bgColor) } },
						},
					},
				},
				fields:
					"shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color",
			},
		},
	);
}

function addBottomLine(slideId, requests, config) {
	const lineId = `tab_line_${slideId}_${newGuid()}`;
	requests.push(
		{
			createLine: {
				objectId: lineId,
				lineCategory: "STRAIGHT",
				elementProperties: {
					pageObjectId: slideId,
					size: {
						height: { magnitude: 0, unit: "PT" },
						width: { magnitude: config.totalWidth, unit: "PT" },
					},
					transform: {
						translateX: 0,
						translateY: config.y + config.height,
						scaleX: 1,
						scaleY: 1,
						unit: "PT",
					},
				},
			},
		},
		{
			updateLineProperties: {
				objectId: lineId,
				lineProperties: { lineFill: solidFill(config.mainColor) },
				fields: "lineFill.solidFill.color",
			},
		},
	);
}

function solidFill(hex) {
	return { solidFill: { color: { rgbColor: hexToRgb(hex) } } };
}

function hexToRgb(hex) {
	const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
	return m
		? {
				red: parseInt(m[1], 16) / 255,
				green: parseInt(m[2], 16) / 255,
				blue: parseInt(m[3], 16) / 255,
			}
		: { red: 0, green: 0, blue: 0 };
}

function newGuid() {
	return Utilities.getUuid().replace(/-/g, "").slice(0, 8);
}
