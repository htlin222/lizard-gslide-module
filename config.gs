// Configuration settings for the Google Slides module
var main_color = '#3D6869';
var main_font_family = 'Source Sans Pro'; 
var water_mark_text = 'ⓒ Hsieh-Ting Lin';
var label_font_size = 14
const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();

/**
 * Runs automatically when the document is opened.
 * This is a simple trigger that has limited permissions.
 * Creates a custom menu and optionally applies theme if it's a new presentation
 */
function onOpen() {
  try {
    // Try to create the menu using the simple trigger
    createCustomMenu();
    
    // Check if this is a new presentation (no slides or just one empty slide)
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    
    if (slides.length <= 1) {
      // This appears to be a new presentation, automatically apply the theme
      applyThemeToCurrentPresentation();
      Logger.log('New presentation detected - theme automatically applied');
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
    
    // Create the batch processing submenu
    const batchMenu = ui.createMenu("🗃️ 批次處理")
      .addItem("🛠 同時更新所有", "confirmRunAll")
      .addItem("🔄 更新進度條", "runUpdateProgressBars")
      .addItem("📑 更新標籤頁", "runProcessTabs")
      .addItem("📚 更新章節導覽", "runProcessSectionBoxes")
      .addItem("🦶 更新 Footer", "runUpdateTitleFootnotes")
      .addItem("💧 切換浮水印", "runToggleWaterMark");

    // Create the beautify submenu
    const beautifyMenu = ui.createMenu("🎨 單頁美化")
      .addItem("📅 更新日期", "updateDateInFirstSlide")
      .addItem("📏 加上網格", "toggleGrids")
      .addItem("🔰 加上badge", "convertToBadges")
      .addItem("🍡 貼上在同一處", "duplicateImageInPlace");

    // Create the add new content submenu
    const createMenu = ui.createMenu("🖖 新增")
      .addItem("👆 取得前一頁的標題", "copyPreviousTitleText")
      .addItem("👇 標題加到新的下頁", "createNextSlideWithCurrentTitle")
      .addItem("🎨 套用主題", "applyThemeToCurrentPresentation");

    // Add all submenus to the main menu and add it to the UI
    ui.createMenu("🛠 工具選單")
      .addSubMenu(batchMenu)
      .addSubMenu(beautifyMenu)
      .addSubMenu(createMenu)
      .addItem("🔁 點這手動更新", "showMenuManually")
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

/**
 * Apply theme from a source presentation to the current presentation
 * This preserves the content of the current presentation while applying the theme/styles from the source
 */
function applyThemeToCurrentPresentation() {
  // Add debugging information
  Logger.log('Starting theme application process...');
  
  // Source presentation with the desired theme/styles
  const sourcePresentationId = '1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220'; 
  Logger.log('Source presentation ID: ' + sourcePresentationId);
  
  // Get the current presentation ID from the script properties
  const currentPresentationId = PropertiesService.getScriptProperties().getProperty('presentationId') || 
                              SlidesApp.getActivePresentation().getId();
  Logger.log('Current presentation ID: ' + currentPresentationId);
  
  // Open both presentations
  const sourcePresentation = SlidesApp.openById(sourcePresentationId);
  const currentPresentation = SlidesApp.openById(currentPresentationId);
  
  // Apply the theme from source to current presentation
  applyTheme(sourcePresentation, currentPresentation);
  
  Logger.log('Theme applied to current presentation: ' + currentPresentationId);
}

/**
 * Apply theme from source presentation to target presentation
 * @param {SlidesApp.Presentation} sourcePresentation - The presentation to copy theme from
 * @param {SlidesApp.Presentation} targetPresentation - The presentation to apply theme to
 */
function applyTheme(sourcePresentation, targetPresentation) {
  try {
    // According to the documentation, when we append a slide from another presentation,
    // the master slides and layouts are automatically copied if they don't exist in the target
    Logger.log('Starting applyTheme function...');
    
    // 1. Get a slide from the source presentation to copy
    const sourceSlides = sourcePresentation.getSlides();
    
    if (sourceSlides.length === 0) {
      Logger.log('Error: Source presentation has no slides');
      return false;
    }
    
    // 2. Append the first slide from the source to the target presentation
    // This will automatically copy the theme (master slides and layouts)
    const copiedSlide = targetPresentation.appendSlide(sourceSlides[0]);
    Logger.log('Theme slide copied successfully');
    
    // 3. Update the title text box on the copied slide to match the current presentation's name
    try {
      // Get the current presentation name
      const presentationName = targetPresentation.getName();
      Logger.log('Current presentation name: ' + presentationName);
      
      // Find the title shape on the copied slide
      const shapes = copiedSlide.getShapes();
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        
        // Check if this is a text box that might be the title
        if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
          const textRange = shape.getText();
          const placeholder = shape.getPlaceholderType();
          
          // If it's a title placeholder or the first text box, update it
          if (placeholder === SlidesApp.PlaceholderType.TITLE || i === 0) {
            textRange.setText(presentationName);
            Logger.log('Updated title text to: ' + presentationName);
            break;
          }
        }
      }
    } catch (titleError) {
      Logger.log('Error updating title: ' + titleError.toString());
      // Continue even if updating the title fails
    }
    
    // For new presentations, this is sufficient - the theme is now available
    // Any new slides created will use the new theme
    
    // If the presentation already has other slides, we'll keep the copied slide
    // at the end for reference, but we won't try to modify existing slides
    // since the setLayout method doesn't work on existing slides
    
    Logger.log('Theme imported successfully - new slides will use this theme');
    return true;
  } catch (error) {
    Logger.log('Error applying theme: ' + error.toString());
    return false;
  }
}

/**
 * Legacy function - creates a new presentation with the theme from the source
 * Kept for backward compatibility
 */
function createStyledPresentation() {
  const sourcePresentationId = '1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220'; // 來源簡報 ID
  const newTitle = 'New Presentation with Copied Style';

  // 1. 複製整份簡報
  const newFile = DriveApp.getFileById(sourcePresentationId).makeCopy(newTitle);
  const newPresentationId = newFile.getId();

  // 2. 打開新簡報
  const presentation = SlidesApp.openById(newPresentationId);

  // 3. 清除內容（保留版面樣式）
  const slides = presentation.getSlides();
  slides.forEach(slide => slide.remove());

  // 4. 新增一張空白幻燈片，使用原母片的預設佈局
  presentation.appendSlide(presentation.getMasters()[0].getLayouts()[0]);

  Logger.log('New presentation created with ID: ' + newPresentationId);
}