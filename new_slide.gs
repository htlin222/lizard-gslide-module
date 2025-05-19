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
    slides.forEach(slide => presentation.removeSlide(slide));
  
    // 4. 新增一張空白幻燈片，使用原母片的預設佈局
    presentation.appendSlide(presentation.getMasters()[0].getLayouts()[0]);
  
    Logger.log('New presentation created with ID: ' + newPresentationId);
  }