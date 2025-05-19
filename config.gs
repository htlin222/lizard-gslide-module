// Configuration settings for the Google Slides module
var main_color = '#3D6869';
var main_font_family = 'Source Sans Pro'; 
var water_mark_text = 'ⓒ Hsieh-Ting Lin';
var label_font_size = 14
const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();

function onOpen() {
  createCustomMenu(); // 改用共用的 menu 建立邏輯
}

function showMenuManually() {
  createCustomMenu(); // 呼叫真正建立選單的邏輯
}

function createCustomMenu() {
  const ui = SlidesApp.getUi();

  const batchMenu = ui.createMenu("🗃️ 批次處理")
    .addItem("🛠 同時更新所有", "confirmRunAll")
    .addItem("🔄 更新進度條", "runUpdateProgressBars")
    .addItem("📑 更新標籤頁", "runProcessTabs")
    .addItem("📚 更新章節導覽", "runProcessSectionBoxes")
    .addItem("🦶 更新 Footer", "runUpdateTitleFootnotes")
    .addItem("💧 切換浮水印", "runToggleWaterMark");

  const beautifyMenu = ui.createMenu("🎨 單頁美化")
    .addItem("📅 更新日期", "updateDateInFirstSlide")
    .addItem("📏 加上網格", "toggleGrids")
    .addItem("🔰 加上badge", "convertToBadges")
    .addItem("🍡 貼上在同一處", "duplicateImageInPlace");

  const createMenu = ui.createMenu("🖖 新增")
    .addItem("👆 取得前一頁的標題", "copyPreviousTitleText")
    .addItem("👇 標題加到新的下頁", "createNextSlideWithCurrentTitle");

  ui.createMenu("🛠 工具選單")
    .addSubMenu(batchMenu)
    .addSubMenu(beautifyMenu)
    .addSubMenu(createMenu)
    .addItem("🔁 點這手動更新", "showMenuManually")
    .addToUi();
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

function runProcessSectionBoxes (){
  runRequestProcessors(processSectionBoxes);
}

function runAllFunctions() {
	runRequestProcessors(updateProgressBars, processTabs, updateTitleFootnotes, runProcessSectionBoxes);
  updateDateInFirstSlide();
}

function confirmRunAll() {
  const ui = SlidesApp.getUi();
  const response = ui.alert("確定要執行所有功能？將會執行以下: \nupdateProgressBars, \nprocessTabs, \nupdateTitleFootnotes, \nrunProcessSectionBoxes", ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    runAllFunctions();
  }
}

function runToggleWaterMark() {
  runRequestProcessors(toggleWaterMark);
}