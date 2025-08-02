function applyListToSelectedTextbox() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectedPageElements = selection.getPageElementRange().getPageElements();

  for (const element of selectedPageElements) {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      const textRange = shape.getText();
      const textString = textRange.asString();

      Logger.log("Text content: " + textString);

      const paragraphs = textRange.getParagraphs();
      Logger.log("Paragraph count: " + paragraphs.length);

      // 嘗試套用清單樣式
      try {
        shape.getText().getListStyle()
          .applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
        Logger.log("✅ 成功套用清單樣式");
      } catch (e) {
        Logger.log("❌ 套用失敗: " + e.message);
      }
    }
  }
}

// The applyMarkdownBoldFormatting function has been moved to src/util/markdown_formatter.js

function createAlphaWhiteSquare() {
  // Get the current presentation
  const presentation = SlidesApp.getActivePresentation();
  
  // Get the current slide (or first slide if none selected)
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  // Define square dimensions (100x100 points)
  const size = 100;
  
  // Get slide dimensions to center the square
  const pageWidth = currentSlide.getPageElementById(currentSlide.getPageElements()[0].getObjectId()).getInherentWidth() || 720;
  const pageHeight = currentSlide.getPageElementById(currentSlide.getPageElements()[0].getObjectId()).getInherentHeight() || 540;
  
  // Calculate position to center the square
  const left = (pageWidth - size) / 2;
  const top = (pageHeight - size) / 2;
  
  // Create a rectangle shape (square)
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, size, size);
  
  // Set the fill to white with 50% transparency
  const fill = shape.getFill();
  fill.setSolidFill('#FFFFFF', 0.5); // White color with alpha 0.5
  
  // Optional: Remove border/stroke
  const line = shape.getBorder();
  line.setTransparent();
  
  console.log('White square with 50% transparency created successfully!');
  
  return shape;
}

// Alternative function with custom position and size
function createCustomAlphaWhiteSquare(leftPos = 100, topPos = 100, squareSize = 100) {
  const presentation = SlidesApp.getActivePresentation();
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  // Create the square
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, leftPos, topPos, squareSize, squareSize);
  
  // Set white fill with 50% transparency
  shape.getFill().setSolidFill('#FFFFFF', 0.5);
  
  // Remove border
  shape.getBorder().setTransparent();
  
  return shape;
}

function createGradientSquare() {
  // Get the current presentation
  const presentation = SlidesApp.getActivePresentation();
  
  // Get the current slide (or first slide if none selected)
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  // Define square dimensions (150x150 points)
  const size = 150;
  
  // Get slide dimensions to center the square
  const pageWidth = 720; // Standard slide width
  const pageHeight = 540; // Standard slide height
  
  // Calculate position to center the square
  const left = (pageWidth - size) / 2;
  const top = (pageHeight - size) / 2;
  
  // Create a rectangle shape (square)
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, size, size);
  
  // Create gradient fill
  const fill = shape.getFill();
  fill.setGradientFill(SlidesApp.GradientType.LINEAR, 90); // 90 degrees (top to bottom)
  
  // Add gradient stops
  const gradientFill = fill.getGradientFill();
  gradientFill.addStop('#FF6B6B', 0, 0.8);   // Red at start with 80% opacity
  gradientFill.addStop('#4ECDC4', 0.5, 0.6); // Teal in middle with 60% opacity  
  gradientFill.addStop('#45B7D1', 1, 0.4);   // Blue at end with 40% opacity
  
  // Remove border
  const line = shape.getBorder();
  line.setTransparent();
  
  console.log('Gradient square created successfully!');
  
  return shape;
}

function createAlphaWhiteSquare() {
  // Get the current presentation
  const presentation = SlidesApp.getActivePresentation();
  
  // Get the current slide (or first slide if none selected)
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  // Define square dimensions (100x100 points)
  const size = 100;
  
  // Get slide dimensions to center the square
  const pageWidth = 720; // Standard slide width
  const pageHeight = 540; // Standard slide height
  
  // Calculate position to center the square
  const left = (pageWidth - size) / 2;
  const top = (pageHeight - size) / 2;
  
  // Create a rectangle shape (square)
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, size, size);
  
  // Set the fill to white with 50% transparency
  const fill = shape.getFill();
  fill.setSolidFill('#FFFFFF', 0.5); // White color with alpha 0.5
  
  // Optional: Remove border/stroke
  const line = shape.getBorder();
  line.setTransparent();
  
  console.log('White square with 50% transparency created successfully!');
  
  return shape;
}

// Create radial gradient square
function createRadialGradientSquare() {
  const presentation = SlidesApp.getActivePresentation();
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  const size = 150;
  const left = (720 - size) / 2;
  const top = (540 - size) / 2;
  
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, size, size);
  
  // Create radial gradient
  const fill = shape.getFill();
  fill.setGradientFill(SlidesApp.GradientType.RADIAL);
  
  const gradientFill = fill.getGradientFill();
  gradientFill.addStop('#FFFFFF', 0, 0.9);    // White center with 90% opacity
  gradientFill.addStop('#FF6B6B', 0.6, 0.7);  // Red middle with 70% opacity
  gradientFill.addStop('#2C3E50', 1, 0.3);    // Dark blue edge with 30% opacity
  
  shape.getBorder().setTransparent();
  
  console.log('Radial gradient square created!');
  return shape;
}

// Create custom gradient with specified colors and direction
function createCustomGradient(colors = ['#FF6B6B', '#4ECDC4', '#45B7D1'], angle = 45, alphas = [0.8, 0.6, 0.4]) {
  const presentation = SlidesApp.getActivePresentation();
  const currentSlide = presentation.getSelection().getCurrentPage() || presentation.getSlides()[0];
  
  const size = 150;
  const left = (720 - size) / 2;
  const top = (540 - size) / 2;
  
  const shape = currentSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, size, size);
  
  const fill = shape.getFill();
  fill.setGradientFill(SlidesApp.GradientType.LINEAR, angle);
  
  const gradientFill = fill.getGradientFill();
  
  // Add color stops based on provided arrays
  for (let i = 0; i < colors.length; i++) {
    const position = i / (colors.length - 1); // Distribute evenly from 0 to 1
    const alpha = alphas[i] || 1; // Default to full opacity if not specified
    gradientFill.addStop(colors[i], position, alpha);
  }
  
  shape.getBorder().setTransparent();
  
  console.log(`Custom gradient square created with ${colors.length} colors at ${angle}° angle`);
  return shape;
}