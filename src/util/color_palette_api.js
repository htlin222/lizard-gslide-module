/**
 * 配色方案 API 函數
 * 提供與 Google Slides 整合的配色功能
 */

/**
 * 生成配色方案並返回結果
 * @param {string} baseHex - 主色 HEX 值
 * @param {string} scheme - 配色方案類型
 * @returns {Object} 包含配色方案和成功狀態的物件
 */
function generateColorScheme(baseHex, scheme) {
	try {
		const colors = generateColorPalette(baseHex, scheme);
		return {
			success: true,
			colors: colors,
			message: `成功生成 ${scheme} 配色方案`,
		};
	} catch (error) {
		return {
			success: false,
			colors: [],
			message: error.message,
		};
	}
}

/**
 * 將顏色套用到選中的物件
 * @param {string} hexColor - HEX 顏色值
 * @returns {Object} 操作結果
 */
function applyColorToSelected(hexColor) {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const selection = SlidesApp.getActivePresentation().getSelection();

		if (!selection) {
			return {
				success: false,
				message: "請先選擇要套用顏色的物件",
			};
		}

		const selectedElements = selection.getPageElements();
		if (!selectedElements || selectedElements.length === 0) {
			return {
				success: false,
				message: "沒有選中任何物件",
			};
		}

		let appliedCount = 0;
		const requests = [];

		selectedElements.forEach((element) => {
			const elementId = element.getObjectId();

			// 根據元素類型套用顏色
			if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
				// 形狀背景色
				requests.push({
					updateShapeProperties: {
						objectId: elementId,
						shapeProperties: {
							shapeBackgroundFill: {
								solidFill: {
									color: {
										rgbColor: hexToRgb(hexColor),
									},
								},
							},
						},
						fields: "shapeBackgroundFill.solidFill.color",
					},
				});
				appliedCount++;
			} else if (
				element.getPageElementType() === SlidesApp.PageElementType.TABLE
			) {
				// 表格處理可以在這裡擴展
				appliedCount++;
			}
		});

		// 執行批次更新
		if (requests.length > 0) {
			runRequestProcessors([requests]);
		}

		return {
			success: true,
			message: `成功套用顏色到 ${appliedCount} 個物件`,
		};
	} catch (error) {
		return {
			success: false,
			message: `套用顏色時發生錯誤: ${error.message}`,
		};
	}
}

/**
 * 建立色卡頁面
 * @param {Array} colors - 顏色陣列
 * @param {string} schemeName - 配色方案名稱
 * @returns {Object} 操作結果
 */
function createColorPalettePage(colors, schemeName) {
	try {
		const presentation = SlidesApp.getActivePresentation();

		// 建立新的投影片
		const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

		// 設定標題
		const titleShape = newSlide.insertTextBox(
			`${schemeName} 配色方案`,
			50,
			50,
			600,
			60,
		);
		const titleTextStyle = titleShape.getText().getTextStyle();
		titleTextStyle
			.setFontSize(24)
			.setFontWeight("bold")
			.setForegroundColor("#333333");

		// 計算色塊位置
		const colorBlockWidth = 100;
		const colorBlockHeight = 80;
		const startX = 100;
		const startY = 150;
		const spacing = 120;

		const requests = [];

		// 為每個顏色建立色塊和標籤
		colors.forEach((color, index) => {
			const x = startX + index * spacing;
			const y = startY;

			// 建立色塊
			const colorBlockId = `colorBlock_${index}_${Date.now()}`;
			requests.push({
				createShape: {
					objectId: colorBlockId,
					shapeType: "RECTANGLE",
					elementProperties: {
						pageObjectId: newSlide.getObjectId(),
						size: {
							width: { magnitude: colorBlockWidth, unit: "PT" },
							height: { magnitude: colorBlockHeight, unit: "PT" },
						},
						transform: {
							scaleX: 1,
							scaleY: 1,
							translateX: x,
							translateY: y,
							unit: "PT",
						},
					},
				},
			});

			// 設定色塊顏色
			requests.push({
				updateShapeProperties: {
					objectId: colorBlockId,
					shapeProperties: {
						shapeBackgroundFill: {
							solidFill: {
								color: {
									rgbColor: hexToRgb(color),
								},
							},
						},
						outline: {
							outlineFill: {
								solidFill: {
									color: {
										rgbColor: { red: 0.8, green: 0.8, blue: 0.8 },
									},
								},
							},
							weight: {
								magnitude: 1,
								unit: "PT",
							},
						},
					},
					fields: "shapeBackgroundFill.solidFill.color,outline",
				},
			});

			// 建立 HEX 標籤
			const labelId = `colorLabel_${index}_${Date.now()}`;
			requests.push({
				createShape: {
					objectId: labelId,
					shapeType: "TEXT_BOX",
					elementProperties: {
						pageObjectId: newSlide.getObjectId(),
						size: {
							width: { magnitude: colorBlockWidth, unit: "PT" },
							height: { magnitude: 30, unit: "PT" },
						},
						transform: {
							scaleX: 1,
							scaleY: 1,
							translateX: x,
							translateY: y + colorBlockHeight + 10,
							unit: "PT",
						},
					},
				},
			});

			// 設定標籤文字
			requests.push({
				insertText: {
					objectId: labelId,
					text: color.toUpperCase(),
				},
			});

			// 設定標籤樣式
			requests.push({
				updateTextStyle: {
					objectId: labelId,
					style: {
						fontSize: { magnitude: 12, unit: "PT" },
						foregroundColor: {
							opaqueColor: {
								rgbColor: { red: 0.2, green: 0.2, blue: 0.2 },
							},
						},
						fontFamily: "Arial",
					},
					fields: "fontSize,foregroundColor,fontFamily",
				},
			});

			// 文字置中對齊
			requests.push({
				updateParagraphStyle: {
					objectId: labelId,
					style: {
						alignment: "CENTER",
					},
					fields: "alignment",
				},
			});
		});

		// 執行批次請求
		if (requests.length > 0) {
			runRequestProcessors([requests]);
		}

		return {
			success: true,
			message: `成功建立 ${schemeName} 色卡頁面`,
			slideId: newSlide.getObjectId(),
		};
	} catch (error) {
		return {
			success: false,
			message: `建立色卡頁面時發生錯誤: ${error.message}`,
		};
	}
}

/**
 * 將 HEX 顏色轉換為 RGB 物件（Google Slides API 格式）
 * @param {string} hex - HEX 顏色值
 * @returns {Object} RGB 顏色物件
 */
function hexToRgb(hex) {
	const normalizedHex = normalizeHex(hex).replace("#", "");
	const r = parseInt(normalizedHex.substr(0, 2), 16) / 255;
	const g = parseInt(normalizedHex.substr(2, 2), 16) / 255;
	const b = parseInt(normalizedHex.substr(4, 2), 16) / 255;

	return {
		red: r,
		green: g,
		blue: b,
	};
}

/**
 * 開啟配色方案側邊欄
 */
function openColorPaletteSidebar() {
	try {
		const htmlTemplate = HtmlService.createTemplateFromFile(
			"src/components/colorPaletteSidebar",
		);
		const htmlOutput = htmlTemplate
			.evaluate()
			.setTitle("配色方案生成器")
			.setWidth(350);

		SlidesApp.getUi().showSidebar(htmlOutput);
	} catch (error) {
		SlidesApp.getUi().alert("開啟側邊欄時發生錯誤: " + error.message);
	}
}

/**
 * 獲取可用的配色方案選項
 * @returns {Array} 配色方案選項陣列
 */
function getColorSchemeOptions() {
	return [
		{ value: "monochromatic", label: "單色系 (Monochromatic)" },
		{ value: "analogous", label: "類似色 (Analogous)" },
		{ value: "complementary", label: "互補色 (Complementary)" },
		{ value: "split-complementary", label: "分裂互補色 (Split Complementary)" },
		{ value: "triadic", label: "三分色 (Triadic)" },
		{ value: "tetradic", label: "四方色 (Tetradic)" },
		{ value: "golden-ratio", label: "黃金比例 (Golden Ratio)" },
	];
}

/**
 * 複製文字到剪貼簿（模擬功能，實際需要在前端實現）
 * @param {string} text - 要複製的文字
 * @returns {Object} 操作結果
 */
function copyToClipboard(text) {
	// 這個函數主要是為了 API 一致性
	// 實際的複製功能會在前端 JavaScript 中實現
	return {
		success: true,
		message: `已準備複製: ${text}`,
	};
}
