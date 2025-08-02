// Logging functionality for Google Slides module

function getSlideTitles() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const titles = [];

  slides.forEach((slide, index) => {
    const pageElements = slide.getPageElements();
    let titleText = "";

    for (let element of pageElements) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const text = shape.getText().asString().trim();

        if (text.length > 0) {
          titleText = text;
          break; // 假設第一個有文字的 shape 是 title
        }
      }
    }

    titles.push(`Slide ${index + 1}: ${titleText || '[No title found]'}`);
  });

  Logger.log(titles);
  return titles;
}   
 



function logSlideLayouts() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var layout = slide.getLayout();
    var layoutName = layout ? layout.getLayoutName() : 'No layout';
    console.log('Slide ' + (i + 1) + ' layout: ' + layoutName);
  }
} 

function logSectionTitleSlides() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var layout = slide.getLayout();
    var layoutName = layout ? layout.getLayoutName() : '';

    if (layoutName === 'SECTION_TITLE_AND_DESCRIPTION') {
      var shapes = slide.getShapes();
      var titleText = 'No non-empty text box found';

      for (var j = 0; j < shapes.length; j++) {
        if (shapes[j].getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
          var text = shapes[j].getText().asString().trim();
          if (text) {
            titleText = text;
            break;
          }
        }
      }

      console.log('Slide ' + (i + 1) + ' layout: ' + layoutName);
      console.log('Title text: ' + titleText);
    }
  }
}

function logCurrentSlideElements() {
  // Get the active presentation
  const presentation = SlidesApp.getActivePresentation();
  
  // Get the current slide (the one being viewed/edited)
  const slides = presentation.getSlides();
  const currentSlide = slides[0]; // This gets the first slide - see note below
  
  console.log(`Analyzing slide: ${currentSlide.getObjectId()}`);
  console.log('='.repeat(50));
  
  // Get all page elements on the current slide
  const pageElements = currentSlide.getPageElements();
  
  console.log(`Total elements found: ${pageElements.length}`);
  console.log('');
  
  // Loop through each element and log its details
  pageElements.forEach((element, index) => {
    console.log(`Element ${index + 1}:`);
    console.log(`  Type: ${element.getPageElementType()}`);
    console.log(`  Object ID: ${element.getObjectId()}`);
    
    // Get position and size
    const transform = element.getTransform();
    console.log(`  Position: (${transform.getTranslateX()}, ${transform.getTranslateY()})`);
    console.log(`  Size: ${element.getWidth()} x ${element.getHeight()}`);
    
    // Get type-specific information
    const elementType = element.getPageElementType();
    
    switch (elementType) {
      case SlidesApp.PageElementType.SHAPE:
        const shape = element.asShape();
        console.log(`  Shape Type: ${shape.getShapeType()}`);
        console.log(`  Text: "${shape.getText().asString()}"`);
        break;
        
      case SlidesApp.PageElementType.TEXT_BOX:
        const textBox = element.asShape();
        console.log(`  Text Content: "${textBox.getText().asString()}"`);
        break;
        
      case SlidesApp.PageElementType.IMAGE:
        const image = element.asImage();
        console.log(`  Image Title: ${image.getTitle() || 'No title'}`);
        console.log(`  Image Alt Text: ${image.getDescription() || 'No alt text'}`);
        break;
        
      case SlidesApp.PageElementType.TABLE:
        const table = element.asTable();
        console.log(`  Table Size: ${table.getNumRows()} rows x ${table.getNumColumns()} columns`);
        break;
        
      case SlidesApp.PageElementType.LINE:
        console.log(`  Line element`);
        break;
        
      case SlidesApp.PageElementType.VIDEO:
        const video = element.asVideo();
        console.log(`  Video Title: ${video.getTitle() || 'No title'}`);
        break;
        
      default:
        console.log(`  Other element type: ${elementType}`);
    }
    
    console.log('  ---');
  });
}

// Alternative function to analyze a specific slide by index
function logSlideElementsByIndex(slideIndex = 0) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  if (slideIndex >= slides.length) {
    console.log(`Error: Slide index ${slideIndex} doesn't exist. Total slides: ${slides.length}`);
    return;
  }
  
  const slide = slides[slideIndex];
  console.log(`Analyzing slide ${slideIndex + 1} of ${slides.length}`);
  console.log(`Slide ID: ${slide.getObjectId()}`);
  console.log('='.repeat(50));
  
  const pageElements = slide.getPageElements();
  
  pageElements.forEach((element, index) => {
    console.log(`Element ${index + 1}: ${element.getPageElementType()} (ID: ${element.getObjectId()})`);
  });
}

// Function to analyze all slides
function logAllSlidesElements() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  slides.forEach((slide, slideIndex) => {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`SLIDE ${slideIndex + 1}`);
    console.log(`${'='.repeat(60)}`);
    
    const pageElements = slide.getPageElements();
    
    pageElements.forEach((element, elementIndex) => {
      console.log(`  ${elementIndex + 1}. ${element.getPageElementType()} - ${element.getObjectId()}`);
    });
  });
}

function logCurrentSlideElements() {
  try {
    // Get the active presentation
    const presentation = SlidesApp.getActivePresentation();
    
    // Get the selection to find current slide
    const selection = SlidesApp.getActivePresentation().getSelection();
    let currentSlide;
    
    // Try to get current slide from selection
    if (selection && selection.getCurrentPage()) {
      currentSlide = selection.getCurrentPage();
    } else {
      // Fallback to first slide if no selection
      currentSlide = presentation.getSlides()[0];
      console.log("Note: Using first slide as fallback (no current slide detected)");
    }
    
    console.log(`Current Slide ID: ${currentSlide.getObjectId()}`);
    console.log('='.repeat(40));
    
    // Get all elements on current slide
    const elements = currentSlide.getPageElements();
    
    console.log(`Total elements: ${elements.length}`);
    console.log('');
    
    // Log each element with its order
    elements.forEach((element, index) => {
      console.log(`${index + 1}. ${element.getPageElementType()} (ID: ${element.getObjectId()})`);
      
      // Add brief content info for text elements
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const text = shape.getText().asString().trim();
        if (text) {
          console.log(`   Text: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
        }
      }
    });
    
  } catch (error) {
    console.log(`Error: ${error.toString()}`);
  }
}