// Page number module for Google Slides
/**
 * 批次建立頁碼文字
 */
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
						width: { magnitude: 70, unit: "PT" },
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