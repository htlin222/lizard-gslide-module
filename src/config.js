// Configuration settings for the Google Slides module
var main_color = "#3D6869";
var base_color = "#FFFFFF";
var text_color = "#333333";
var accent_color = "#f29424";
var sub1_color = "#E7EAE7";
var sub2_color = "#E7F9F5";
var main_font_family = "Source Sans Pro";
var water_mark_text = "ⓒ Hsieh-Ting Lin";
var label_font_size = 14;
var progressBarHeight = 2.5;
const sourcePresentationId = "1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220";

// Properties service keys for storing configuration
const CONFIG_KEYS = {
	MAIN_COLOR: "main_color",
	FONT_FAMILY: "main_font_family",
	WATERMARK_TEXT: "water_mark_text",
	FONT_SIZE: "label_font_size",
	PROGRESS_BAR_HEIGHT: "progress_bar_height",
};

/**
 * Runs automatically when the document is opened.
 * This is a simple trigger that has limited permissions.
 * Creates a custom menu and optionally applies theme if it's a new presentation
 */
function onOpen() {
	try {
		// Load saved configuration first
		loadSavedConfiguration();

		// Try to create the menu using the simple trigger
		createCustomMenu();

		// Check if this is a new presentation (no slides or just one empty slide)
		const presentation = SlidesApp.getActivePresentation();
		const slides = presentation.getSlides();

		if (slides.length <= 1) {
			// This appears to be a new presentation, automatically apply the theme
			applyThemeToCurrentPresentation();
			Logger.log("New presentation detected - theme automatically applied");
		}
	} catch (e) {
		// If it fails, log the error but don't halt execution
		console.log("Error in onOpen: " + e.message);
	}
}

/**
 * Function to manually show the menu.
 * This can be called from the UI when needed.
 */
function showMenuManually() {
	createCustomMenu(); // 呼叫真正建立選單的邏輯
}

/**
 * Creates the custom menu in the Google Slides UI.
 * This function is designed to work in both simple trigger and manual execution contexts.
 */
function createCustomMenu() {
	try {
		// Try to get the UI - this might fail in some contexts
		const ui = SlidesApp.getUi();

		// ── Menu 1: global / whole-deck operations (settings + batch updates) ──
		ui.createMenu("⚙ 設定與批次")
			// One-click: run every batch processor at once
			.addItem("🛠 同時執行所有功能", "confirmRunAll")
			// One-click: catch PPTX-imported (LZ-tagged) elements, apply canonical
			// font + style online, then rebuild all chrome
			.addItem("🦎 套用 PPTX 匯入樣式 (LZ)", "runLzApplyAll")
			.addSeparator()
			// Setup & configuration
			.addItem("🎨 套用蜥蜴主題", "applyThemeToCurrentPresentation")
			.addItem("⚙ 打開設定面板", "showConfigSidebar")
			.addItem("🔑 設定 AI 金鑰 (Groq)", "showAiKeySetup")
			.addItem("🔁 重新整理選單", "showMenuManually")
			.addSeparator()
			// Self-update: pull the latest script from GitHub into this clone
			.addItem("🔄 更新腳本", "fetchLatestScript")
			.addItem("🔍 檢查更新", "menuCheckForUpdate")
			.addItem("⏰ 啟用開啟時檢查更新", "enableUpdateOnOpenTrigger")
			.addSeparator()
			// Batch updates applied across all slides
			.addItem("🔄 更新進度條", "runUpdateProgressBars")
			.addItem("📑 更新標籤頁", "runProcessTabs")
			.addItem("📚 更新 SECTION Header", "runProcessSectionBoxes")
			.addItem("🦶 更新 Footer", "runUpdateTitleFootnotes")
			.addItem("📅 更新日期 yyyy-mm-dd", "updateDateInFirstSlide")
			.addItem("💧 浮水印開/關", "runToggleWaterMark")
			.addToUi();

		// ── Menu 2: add new elements to the current slide ──
		ui.createMenu("✨ 加入元素")
			// All "minter" generators live in one submenu so the menu stays tidy.
			.addSubMenu(
				ui
					.createMenu("🏭 鑄造器")
					.addItem("🔲 表格鑄造器⭐", "showTableMinterDialog")
					.addItem("🔳 網格鑄造器⭐", "showGridMinterDialog")
					.addItem("📌 Callout 鑄造器⭐", "showCalloutMinterDialog")
					.addSeparator()
					.addItem("📊 KPI 大數字鑄造器⭐", "showKpiMinterDialog")
					.addItem("📈 長條圖鑄造器⭐", "showBarChartMinterDialog")
					.addItem("⚖ 對照鑄造器⭐", "showCompareMinterDialog")
					.addItem("⏱ 時間軸鑄造器⭐", "showTimelineMinterDialog")
					.addItem("🪜 步驟鑄造器⭐", "showStepsMinterDialog")
					.addSeparator()
					.addItem("📋 議程/目錄鑄造器⭐", "showAgendaMinterDialog")
					.addItem("🎯 重點摘要鑄造器⭐", "showTakeawaysMinterDialog")
					.addItem("🖼 圖片陣列鑄造器⭐", "showGalleryMinterDialog")
					.addItem("😀 Icon 鑄造器⭐", "showIconMinterDialog"),
			)
			.addSeparator()
			// Shapes & decorations
			.addItem("↙ 加上一個大箭頭", "drawArrowOnCurrentSlide")
			.addItem("🔢 加上數字遞增圓圈", "addNextNumberCircle")
			.addItem("🔰 將文字轉換成 badge", "convertToBadges")
			.addItem("❄ 為元素加上 45 度影子", "createOffsetBlueShape")
			.addSeparator()
			// Tables & images
			.addItem("🍽 快速美化表格", "fastStyleSelectedTable")
			.addItem("🌆 原地貼上", "duplicateImageInPlace")
			.addItem("🏙 覆蓋半透明方塊", "coverImageWithWhite")
			.addItem("🏞 半透明遮罩⭐", "maskImage")
			.addSeparator()
			// Helpers
			.addItem("📏 開/關網格", "toggleGrids")
			.addItem("🔍 檢視物件屬性", "showSelectedObjectPropertiesDialog")
			.addToUi();

		// ── Menu 3: manipulate / lay out existing shapes ──
		ui.createMenu("🔷 形狀排版")
			// Split shapes
			.addItem("📐 分割成網格⭐", "showSplitShapeDialog")
			.addItem("📄 分割成多欄", "showMultipleColumnsDialog")
			.addSeparator()
			// Child shapes & flowchart
			.addItem("🏗 建立子形狀", "showCreateChildShapesDialog")
			.addItem("🔤 自動語法解析", "autoCreateChildShapesFromLines")
			.addItem("🔗 流程圖工具⭐", "showFlowchartSidebar")
			.addSeparator()
			// Spacing & alignment
			.addItem("📏 調整間距", "showSetGapDialog")
			.addItem("🧙 智能間距重設", "showSmartGapResetDialog")
			.addItem("📊 平均間距置中", "runAveragePadding")
			.addSeparator()
			// Connectors between shapes
			.addItem("⇣ 兩者間加上垂直線", "insertVerticalDashedLineBetween")
			.addItem("⇢ 兩者間加上水平線⭐", "insertHorizontalDashedLineBetween")
			.addSeparator()
			// Restyle a shape
			.addItem("**B** 套用粗體樣式", "applyBoldStyleToSelectedShape")
			.addToUi();

		// ── Menu 4: cross-slide workflow + import / export ──
		ui.createMenu("🔁 跨頁與匯出")
			// Title / slide navigation
			.addItem("👆 在上面加入前一頁的標題", "copyPreviousTitleText")
			.addItem("👇 新增一頁並加入當前標題", "createNextSlideWithCurrentTitle")
			.addSeparator()
			// Markdown import
			.addItem("📝 Markdown 轉換成投影片", "showMarkdownToSlidesDialog")
			.addItem("📋 Markdown 側邊欄⭐", "showMarkdownSidebar")
			.addItem("**B** Markdown 粗體格式", "runApplyMarkdownBoldFormatting")
			.addSeparator()
			// Markdown export
			.addItem("📤 匯出成 Markdown⭐", "showExportMarkdownDialog")
			.addItem("📤 匯出 Markdown (含圖片)", "showExportMarkdownWithImagesDialog")
			.addSeparator()
			// Speaker notes (AI)
			.addItem("🎤 AI 演講者備註", "showSpeakerNoteSidebar")
			.addItem("📥 批次下載演講者備註 (JSON)", "showExportSpeakerNotesDialog")
			.addSeparator()
			// Color
			.addItem("🎨 配色方案生成器", "openColorPaletteSidebar")
			.addToUi();

		return true; // Menu created successfully
	} catch (e) {
		// Log the error but don't halt execution
		console.log("Error creating menu: " + e.message);
		return false; // Menu creation failed
	}
}

/**
 * Runs one or more slide processing functions that collect batch update requests.
 * Each processor function should accept two parameters: (slides, requests)
 * and push their individual update requests into the shared `requests` array.
 * After collecting all requests, they are sent to the Slides API as a batch update.
 *
 * @param {...function(slides: GoogleAppsScript.Slides.Slide[], requests: Object[])} processors
 *        One or more functions that generate update requests for the Slides API.
 *
 * Example usage:
 *   runRequestProcessors(updateProgressBars);
 *   runRequestProcessors(updateProgressBars, processTabs);
 */
function runRequestProcessors(...processors) {
	const presentation = SlidesApp.getActivePresentation();
	const presentationId = presentation.getId();
	const slides = presentation.getSlides();
	const requests = [];

	processors.forEach((fn) => fn(slides, requests));

	if (requests.length) {
		Slides.Presentations.batchUpdate({ requests }, presentationId);
	}
}

// Menu actions
// 🚀 OPTIMIZED VERSIONS (default)
function runUpdateProgressBars() {
	runOptimizedRequestProcessors(updateProgressBarsOptimized);
}

function runProcessTabs() {
	runOptimizedRequestProcessors(processTabsOptimized);
}

function runUpdateTitleFootnotes() {
	runOptimizedRequestProcessors(updateTitleFootnotesOptimized);
}

function runProcessSectionBoxes() {
	runOptimizedRequestProcessors(processSectionBoxesOptimized);
}

// LZ-Protocol: one command for a freshly PPTX-imported deck — catch every
// LZ-tagged element, re-apply the canonical font + style, rebuild all chrome.
function runLzApplyAll() {
	var ui = SlidesApp.getUi();
	try {
		var n = lzApplyAll();
		ui.alert("🦎 LZ 套用完成", "已重新套用樣式的元素：" + n + " 個。", ui.ButtonSet.OK);
	} catch (e) {
		ui.alert("LZ 套用失敗", String(e), ui.ButtonSet.OK);
	}
}

// Legacy versions (for fallback if needed)
function runUpdateProgressBarsLegacy() {
	runRequestProcessors(updateProgressBars);
}

function runProcessTabsLegacy() {
	runRequestProcessors(processTabs);
}

function runUpdateTitleFootnotesLegacy() {
	runRequestProcessors(updateTitleFootnotes);
}

function runProcessSectionBoxesLegacy() {
	runRequestProcessors(processSectionBoxes);
}

// 🚀 ULTRA MEGA BATCH VERSION: Maximum possible performance
function runAllFunctions() {
	runAllFunctionsUltraMegaBatch();
}

// 🚀 MEGA BATCH VERSION: Single API call (fallback)
// Fallback function - uses same ultra mega batch
function runAllFunctionsFallback() {
	runAllFunctionsUltraMegaBatch();
}

function confirmRunAll() {
	// No confirmation needed - run directly with ultra performance
	runAllFunctions();
}

function runToggleWaterMark() {
	runRequestProcessors(toggleWaterMark);
}

/**
 * Shows the configuration sidebar.
 */
function showConfigSidebar() {
	// Use the modular sidebar approach
	const sidebar = createConfigSidebar();
	SlidesApp.getUi().showSidebar(sidebar);
}

/**
 * Gets the current configuration values for the sidebar.
 * @return {Object} The current configuration values.
 */
function getConfigValues() {
	// Try to load from Properties service first
	const userProperties = PropertiesService.getUserProperties();
	const savedMainColor = userProperties.getProperty(CONFIG_KEYS.MAIN_COLOR);
	const savedFontFamily = userProperties.getProperty(CONFIG_KEYS.FONT_FAMILY);
	const savedWatermarkText = userProperties.getProperty(
		CONFIG_KEYS.WATERMARK_TEXT,
	);
	const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);

	// Get available fonts
	const availableFonts = getAvailableFonts();

	// Get saved progress bar height
	const savedProgressBarHeight = userProperties.getProperty(
		CONFIG_KEYS.PROGRESS_BAR_HEIGHT,
	);

	// Return current values (from Properties if available, otherwise from variables)
	return {
		mainColor: savedMainColor || main_color,
		baseColor: base_color,
		textColor: text_color,
		sub1Color: sub1_color,
		accentColor: accent_color,
		fontFamily: savedFontFamily || main_font_family,
		watermarkText: savedWatermarkText || water_mark_text,
		fontSize: savedFontSize || label_font_size,
		progressBarHeight: savedProgressBarHeight || progressBarHeight,
		availableFonts: availableFonts,
	};
}

/**
 * Gets the available fonts in Google Slides.
 * @return {Array} Array of font family names.
 */
function getAvailableFonts() {
	try {
		// Get all available fonts from Google Slides
		const availableFonts = SlidesApp.getFonts();

		// Convert to an array of font family names
		const fontFamilies = availableFonts.map((font) => font.getFontFamily());

		// Sort alphabetically
		fontFamilies.sort();

		return fontFamilies;
	} catch (e) {
		// If there's an error, return a default list of common fonts
		console.log("Error getting fonts: " + e.message);
		return [
			"Arial",
			"Calibri",
			"Cambria",
			"Comic Sans MS",
			"Courier New",
			"Georgia",
			"Impact",
			"Lato",
			"Montserrat",
			"Open Sans",
			"Roboto",
			"Source Sans Pro",
			"Tahoma",
			"Times New Roman",
			"Trebuchet MS",
			"Verdana",
		];
	}
}

/**
 * Saves the configuration values from the sidebar.
 * @param {Object} config The configuration values to save.
 */
function saveConfigValues(config) {
	// Save to Properties service
	const userProperties = PropertiesService.getUserProperties();
	userProperties.setProperties({
		[CONFIG_KEYS.MAIN_COLOR]: config.mainColor,
		[CONFIG_KEYS.FONT_FAMILY]: config.fontFamily,
		[CONFIG_KEYS.WATERMARK_TEXT]: config.watermarkText,
		[CONFIG_KEYS.FONT_SIZE]: config.fontSize,
		[CONFIG_KEYS.PROGRESS_BAR_HEIGHT]: config.progressBarHeight,
	});

	// Update the global variables
	main_color = config.mainColor;
	main_font_family = config.fontFamily;
	water_mark_text = config.watermarkText;
	label_font_size = Number.parseInt(config.fontSize, 10);
	progressBarHeight = Number.parseInt(config.progressBarHeight, 10);

	return true;
}

/**
 * Saves the configuration values and applies them to the current presentation.
 * @param {Object} config The configuration values to save.
 */
function saveAndApplyConfig(config) {
	// Save the configuration first
	saveConfigValues(config);

	// Apply changes to the current presentation
	// This will update watermarks and other elements that use these settings
	runAllFunctions();

	return true;
}

/**
 * Loads configuration from Properties service when the script runs.
 * This ensures we're using the saved values from previous sessions.
 */
function loadSavedConfiguration() {
	const userProperties = PropertiesService.getUserProperties();
	const savedMainColor = userProperties.getProperty(CONFIG_KEYS.MAIN_COLOR);
	const savedFontFamily = userProperties.getProperty(CONFIG_KEYS.FONT_FAMILY);
	const savedWatermarkText = userProperties.getProperty(
		CONFIG_KEYS.WATERMARK_TEXT,
	);
	const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);
	const savedProgressBarHeight = userProperties.getProperty(
		CONFIG_KEYS.PROGRESS_BAR_HEIGHT,
	);

	// Update the global variables if saved values exist
	if (savedMainColor) main_color = savedMainColor;
	if (savedFontFamily) main_font_family = savedFontFamily;
	if (savedWatermarkText) water_mark_text = savedWatermarkText;
	if (savedFontSize) label_font_size = Number.parseInt(savedFontSize, 10);
	if (savedProgressBarHeight)
		progressBarHeight = Number.parseInt(savedProgressBarHeight, 10);
}

/**
 * Shows a dialog with the specified title and message.
 * Used by the sidebar to display success/error messages.
 *
 * @param {string} title The title of the dialog.
 * @param {string} message The message to display in the dialog.
 */
function showDialog(title, message) {
	const ui = SlidesApp.getUi();
	ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * Runs the averagePadding function to center an element between its neighbors
 */
function runAveragePadding() {
	try {
		const result = averagePadding();
		if (!result) {
			SlidesApp.getUi().alert(
				"Please select a single element or group to center",
			);
		}
	} catch (e) {
		SlidesApp.getUi().alert(
			"Error",
			"An error occurred while centering the element: " + e.message,
		);
	}
}

/**
 * Shows the Markdown to Slides dialog
 */
function showMarkdownToSlidesDialog() {
	try {
		// Create and show the HTML dialog
		const html = HtmlService.createHtmlOutputFromFile(
			"src/components/md2slides-dialog.html",
		)
			.setWidth(800)
			.setHeight(700)
			.setTitle("Markdown to Slides Converter");

		SlidesApp.getUi().showModalDialog(html, "Markdown to Slides");
	} catch (e) {
		console.error("Error showing Markdown to Slides dialog: " + e.message);
		SlidesApp.getUi().alert(
			"Error",
			"Could not open the Markdown to Slides dialog: " + e.message,
		);
	}
}

/**
 * Shows the Markdown to Slides converter as a fixed right sidebar
 */
function showMarkdownSidebar() {
	try {
		// Use the modular sidebar approach
		const sidebar = createMarkdownSidebar();
		SlidesApp.getUi().showSidebar(sidebar);
	} catch (e) {
		console.error("Error showing Markdown sidebar: " + e.message);
		SlidesApp.getUi().alert(
			"Error: Could not open the Markdown sidebar: " + e.message,
		);
	}
}

/**
 * Shows the Table Minter dialog.
 * Lets users turn a Markdown table into a Google Slides-ready styled table
 * on the clipboard (paste into Slides for a native table).
 */
function showTableMinterDialog() {
	try {
		const dialog = createTableMinterDialog();
		SlidesApp.getUi().showModalDialog(dialog, "🔲 表格鑄造器 Table Minter");
	} catch (e) {
		console.error("Error showing Table Minter dialog: " + e.message);
		SlidesApp.getUi().alert(
			"Error: Could not open the Table Minter dialog: " + e.message,
		);
	}
}

/**
 * Shows the Grid Minter dialog.
 * Turns content (pasted or AI-generated) into a grid of styled "unit" cards
 * laid out on the current slide, spilling overflow onto new slides.
 */
function showGridMinterDialog() {
	try {
		const dialog = createGridMinterDialog();
		SlidesApp.getUi().showModalDialog(dialog, "🔳 網格鑄造器 Grid Minter");
	} catch (e) {
		console.error("Error showing Grid Minter dialog: " + e.message);
		SlidesApp.getUi().alert(
			"Error: Could not open the Grid Minter dialog: " + e.message,
		);
	}
}

/**
 * Shows the Callout Minter dialog.
 * Inserts a styled callout from a template, or converts the selected shape into
 * one (header banner + left accent bar), with a live HTML preview.
 */
function showCalloutMinterDialog() {
	try {
		const dialog = createCalloutMinterDialog();
		SlidesApp.getUi().showModalDialog(dialog, "📌 Callout 鑄造器 Callout Minter");
	} catch (e) {
		console.error("Error showing Callout Minter dialog: " + e.message);
		SlidesApp.getUi().alert(
			"Error: Could not open the Callout Minter dialog: " + e.message,
		);
	}
}

/**
 * Opens a minter dialog by factory + title, with shared error handling.
 * @param {function(): HtmlOutput} factory
 * @param {string} title
 */
function showMinterDialog_(factory, title) {
	try {
		SlidesApp.getUi().showModalDialog(factory(), title);
	} catch (e) {
		console.error("Error showing " + title + ": " + e.message);
		SlidesApp.getUi().alert("Error: Could not open " + title + ": " + e.message);
	}
}

/** Shows the KPI / Big Number Minter dialog. */
function showKpiMinterDialog() {
	showMinterDialog_(createKpiMinterDialog, "📊 KPI 大數字鑄造器 KPI Minter");
}

/** Shows the Timeline / Roadmap Minter dialog. */
function showTimelineMinterDialog() {
	showMinterDialog_(createTimelineMinterDialog, "⏱ 時間軸鑄造器 Timeline Minter");
}

/** Shows the Comparison Minter dialog. */
function showCompareMinterDialog() {
	showMinterDialog_(createCompareMinterDialog, "⚖ 對照鑄造器 Compare Minter");
}

/** Shows the Steps Minter dialog. */
function showStepsMinterDialog() {
	showMinterDialog_(createStepsMinterDialog, "🪜 步驟鑄造器 Steps Minter");
}

/** Shows the Image Gallery Minter dialog. */
function showGalleryMinterDialog() {
	showMinterDialog_(createGalleryMinterDialog, "🖼 圖片陣列鑄造器 Gallery Minter");
}

/** Shows the Agenda / TOC Minter dialog. */
function showAgendaMinterDialog() {
	showMinterDialog_(createAgendaMinterDialog, "📋 議程/目錄鑄造器 Agenda Minter");
}

/** Shows the Takeaways Minter dialog. */
function showTakeawaysMinterDialog() {
	showMinterDialog_(createTakeawaysMinterDialog, "🎯 重點摘要鑄造器 Takeaways Minter");
}

/** Shows the Icon Minter dialog. */
function showIconMinterDialog() {
	showMinterDialog_(createIconMinterDialog, "😀 Icon 鑄造器 Icon Minter");
}

/** Shows the Bar Chart Minter dialog. */
function showBarChartMinterDialog() {
	showMinterDialog_(createBarChartMinterDialog, "📈 長條圖鑄造器 Bar Chart Minter");
}
