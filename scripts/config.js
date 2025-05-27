// Configuration settings for the Google Slides module
var main_color = "#3D6869";
var main_font_family = "Source Sans Pro";
var water_mark_text = "ⓒ Hsieh-Ting Lin";
var label_font_size = 14;
var progressBarHeight = 5;
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

    // Create the batch processing menu as a top-level menu
    ui.createMenu("🗃 批次處理")
      .addItem("🛠 同時執行所有功能", "confirmRunAll")
      .addItem("🎨 套用蜥蜴主題", "applyThemeToCurrentPresentation")
      .addItem("⚙️ 打開設定面板", "showConfigSidebar")
      .addItem("🔄 更新進度條", "runUpdateProgressBars")
      .addItem("📑 更新標籤頁", "runProcessTabs")
      .addItem("📚 更新 SECTION Header", "runProcessSectionBoxes")
      .addItem("🦶 更新 Footer", "runUpdateTitleFootnotes")
      .addItem("📅 更新日期 yyyy-mm-dd", "updateDateInFirstSlide")
      .addItem("💧 浮水印開/關", "runToggleWaterMark")
      .addItem("🔁 更新選單項目", "showMenuManually")
      .addToUi();

    // Create the beautify menu as a top-level menu
    ui.createMenu("🎨 加入元素")
      .addItem("📏 開/關網格", "toggleGrids")
      .addItem("❄ 為元素加上45度影子", "createOffsetBlueShape")
      .addItem("↙ 加上一個大箭頭 ", "drawArrowOnCurrentSlide")
      .addItem("⇣ 兩者間加上垂直線", "insertVerticalDashedLineBetween")
      .addItem("⇢ 兩者間加上水平線", "insertHorizontalDashedLineBetween")
      .addItem("🔰 將文字轉換成badge", "convertToBadges")
      .addItem("🌆 原地貼上", "duplicateImageInPlace")
      .addItem("🏙️ 覆蓋半透明方塊", "coverImageWithWhite")
      .addItem("🏞️ 半透明遮罩", "maskImage")
      .addItem("🔢 加上數字遞增圓圈", "addNextNumberCircle")
      .addItem("📐 分割成網格", "showSplitShapeDialog")
      .addItem("🎹 轉換成標注框", "convertShapeToCallout")
      .addItem("🍽️ 快速美化表格", "fastStyleSelectedTable")
      .addItem("📊 平均間距置中", "runAveragePadding")
      .addToUi();

    // Create the add new content menu as a top-level menu
    ui.createMenu("🖖 跨頁功能")
      .addItem("👆 在上面加入前一頁的標題", "copyPreviousTitleText")
      .addItem("👇 新增一頁並加入當前標題", "createNextSlideWithCurrentTitle")
      .addItem("📝 Markdown 轉換成投影片", "showMarkdownDialog")
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
    runProcessSectionBoxes
  );
  updateDateInFirstSlide();
}

function confirmRunAll() {
  const ui = SlidesApp.getUi();
  const response = ui.alert(
    "確定要執行所有功能？將會執行以下: \nupdateProgressBars, \nprocessTabs, \nupdateTitleFootnotes, \nrunProcessSectionBoxes",
    ui.ButtonSet.YES_NO
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
  // Use the modular sidebar approach
  const sidebar = createModularSidebar();
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
    CONFIG_KEYS.WATERMARK_TEXT
  );
  const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);

  // Get available fonts
  const availableFonts = getAvailableFonts();

  // Get saved progress bar height
  const savedProgressBarHeight = userProperties.getProperty(
    CONFIG_KEYS.PROGRESS_BAR_HEIGHT
  );

  // Return current values (from Properties if available, otherwise from variables)
  return {
    mainColor: savedMainColor || main_color,
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
    const fontFamilies = availableFonts.map(function (font) {
      return font.getFontFamily();
    });

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
  label_font_size = parseInt(config.fontSize, 10);
  progressBarHeight = parseInt(config.progressBarHeight, 10);

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
    CONFIG_KEYS.WATERMARK_TEXT
  );
  const savedFontSize = userProperties.getProperty(CONFIG_KEYS.FONT_SIZE);
  const savedProgressBarHeight = userProperties.getProperty(
    CONFIG_KEYS.PROGRESS_BAR_HEIGHT
  );

  // Update the global variables if saved values exist
  if (savedMainColor) main_color = savedMainColor;
  if (savedFontFamily) main_font_family = savedFontFamily;
  if (savedWatermarkText) water_mark_text = savedWatermarkText;
  if (savedFontSize) label_font_size = parseInt(savedFontSize, 10);
  if (savedProgressBarHeight)
    progressBarHeight = parseInt(savedProgressBarHeight, 10);
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
      SlidesApp.getUi().alert('Please select a single element or group to center');
    }
  } catch (e) {
    SlidesApp.getUi().alert('Error', 'An error occurred while centering the element: ' + e.message);
  }
}
