// Configuration settings for the Google Slides module
var main_color = "#3D6869";
var main_font_family = "Source Sans Pro";
var water_mark_text = "ⓒ Hsieh-Ting Lin";
var label_font_size = 14;
const sourcePresentationId = "1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220";
const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();

// Properties service keys for storing configuration
const CONFIG_KEYS = {
  MAIN_COLOR: 'main_color',
  FONT_FAMILY: 'main_font_family',
  WATERMARK_TEXT: 'water_mark_text',
  FONT_SIZE: 'label_font_size'
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

		// Create the batch processing menu as a top-level menu
		ui.createMenu("🗃 批次處理")
			.addItem("🔁 點這手動更新", "showMenuManually")
			.addItem("🛠 同時更新所有", "confirmRunAll")
			.addItem("🎨 套用主題", "applyThemeToCurrentPresentation")
			.addItem("⚙️ 設定面板", "showConfigSidebar")
			.addItem("🔄 更新進度條", "runUpdateProgressBars")
			.addItem("📑 更新標籤頁", "runProcessTabs")
			.addItem("📚 更新章節導覽", "runProcessSectionBoxes")
			.addItem("🦶 更新 Footer", "runUpdateTitleFootnotes")
			.addItem("💧 切換浮水印", "runToggleWaterMark")
			.addToUi();

		// Create the beautify menu as a top-level menu
		ui.createMenu("🎨 單頁美化")
			.addItem("📅 更新日期", "updateDateInFirstSlide")
			.addItem("📏 加上網格", "toggleGrids")
			.addItem("❄ 加上影子", "createOffsetBlueShape")
			.addItem("↙ 加上一個大箭頭 ", "drawArrowOnCurrentSlide")
			.addItem("⇣ 兩者間加上垂直線", "insertVerticalDashedLineBetween")
			.addItem("⇢ 兩者間加上水平線", "insertHorizontalDashedLineBetween")
			.addItem("🔰 加上badge", "convertToBadges")
			.addItem("🍡 貼上在同一處", "duplicateImageInPlace")
			.addItem("🔢 加上數字圓圈", "addNextNumberCircle")
			.addItem('📐 分割成網格', 'showSplitShapeDialog')
			.addItem('💬 轉換成標注框', 'convertShapeToCallout')
			.addToUi();

		// Create the add new content menu as a top-level menu
		ui.createMenu("🖖 新增")
			.addItem("👆 取得前一頁的標題", "copyPreviousTitleText")
			.addItem("👇 標題加到新的下頁", "createNextSlideWithCurrentTitle")
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
function runUpdateProgressBars() {
	runRequestProcessors(updateProgressBars);
}

function runProcessTabs() {
	runRequestProcessors(processTabs);
}

function runUpdateTitleFootnotes() {
	runRequestProcessors(updateTitleFootnotes);
}

function runProcessSectionBoxes() {
	runRequestProcessors(processSectionBoxes);
}

function runAllFunctions() {
	runRequestProcessors(
		updateProgressBars,
		processTabs,
		updateTitleFootnotes,
		runProcessSectionBoxes,
	);
	updateDateInFirstSlide();
}

function confirmRunAll() {
	const ui = SlidesApp.getUi();
	const response = ui.alert(
		"確定要執行所有功能？將會執行以下: \nupdateProgressBars, \nprocessTabs, \nupdateTitleFootnotes, \nrunProcessSectionBoxes",
		ui.ButtonSet.YES_NO,
	);
	if (response === ui.Button.YES) {
		runAllFunctions();
	}
}

function runToggleWaterMark() {
	runRequestProcessors(toggleWaterMark);
}

/**
 * Shows the configuration sidebar.
 */
function showConfigSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Lizard Slides Configuration')
      .setWidth(300);
  SlidesApp.getUi().showSidebar(html);
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
  const savedWatermarkText = userProperties.getProperty(CONFIG_KEYS.WATERMARK_TEXT);
  const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);
  
  // Return current values (from Properties if available, otherwise from variables)
  return {
    mainColor: savedMainColor || main_color,
    fontFamily: savedFontFamily || main_font_family,
    watermarkText: savedWatermarkText || water_mark_text,
    fontSize: savedFontSize || label_font_size
  };
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
    [CONFIG_KEYS.FONT_SIZE]: config.fontSize
  });
  
  // Update the global variables
  main_color = config.mainColor;
  main_font_family = config.fontFamily;
  water_mark_text = config.watermarkText;
  label_font_size = parseInt(config.fontSize, 10);
  
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
  const savedWatermarkText = userProperties.getProperty(CONFIG_KEYS.WATERMARK_TEXT);
  const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);
  
  // Update the global variables if saved values exist
  if (savedMainColor) main_color = savedMainColor;
  if (savedFontFamily) main_font_family = savedFontFamily;
  if (savedWatermarkText) water_mark_text = savedWatermarkText;
  if (savedFontSize) label_font_size = parseInt(savedFontSize, 10);
}
