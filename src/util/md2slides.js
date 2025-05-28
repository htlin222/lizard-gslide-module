/**
 * Markdown to Slides Converter Utility
 * 
 * This utility converts markdown text to Google Slides with the following rules:
 * - H1 headings become SECTION_HEADER slides
 * - H2 headings become TITLE_AND_BODY slides
 * - Text below H2 headings becomes bullet points in the body
 *
 * The approach is modular:
 * 1. Parse markdown into a structured format
 * 2. Create all slides based on the parsed structure
 * 3. Add content to each slide
 * 4. Apply formatting (like bullet points) to the content
 * 5. Apply markdown bold formatting (**text**) to the content
 */

/**
 * Shows a dialog for the user to paste markdown content
 */
function showMarkdownDialog() {
  const html = HtmlService.createTemplateFromFile('src/components/md2slides-dialog')
    .evaluate()
    .setWidth(600)
    .setHeight(500)
    .setTitle('Markdown to Slides Converter');
  
  SlidesApp.getUi().showModalDialog(html, 'Markdown to Slides Converter');
}

/**
 * Converts markdown text to slides
 * @param {string} markdownText - The markdown text to convert
 * @return {boolean} - Success status
 */
function convertMarkdownToSlides(markdownText) {
  try {
    // Step 1: Parse the markdown into a structured format
    const slideStructure = parseMarkdownToStructure(markdownText);
    
    if (slideStructure.length === 0) {
      return false;
    }
    
    // Step 2: Determine where to insert the slides
    const presentation = SlidesApp.getActivePresentation();
    let insertIndex = getInsertIndex(presentation);
    
    // Step 3: Create all slides first
    const createdSlides = [];
    for (let i = 0; i < slideStructure.length; i++) {
      const slideInfo = slideStructure[i];
      
      let slide;
      if (slideInfo.layout === 'SECTION_HEADER') {
        slide = presentation.insertSlide(insertIndex, SlidesApp.PredefinedLayout.SECTION_HEADER);
      } else if (slideInfo.layout === 'TITLE_AND_BODY') {
        slide = presentation.insertSlide(insertIndex, SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      }
      
      createdSlides.push({
        slide: slide,
        info: slideInfo
      });
      
      insertIndex++;
    }
    
    // Step 4: Add content to each slide
    for (let i = 0; i < createdSlides.length; i++) {
      const slideObj = createdSlides[i];
      const slide = slideObj.slide;
      const info = slideObj.info;
      
      // Add content to the slide
      
      // Add title to all slides using the correct approach
      const shapes = slide.getShapes();
      let titleAdded = false;
      
      // First pass: Look for TITLE placeholder
      for (let j = 0; j < shapes.length; j++) {
        const shape = shapes[j];
        try {
          if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.TITLE) {
            shape.getText().setText(info.title);
            // Title added using TITLE placeholder
            titleAdded = true;
            break;
          }
        } catch (e) {
          Logger.log('Error checking placeholder type: ' + e.message);
        }
      }
      
      // If title wasn't added, try another approach
      if (!titleAdded) {
        try {
          const titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
          if (titleShape) {
            titleShape.getText().setText(info.title);
            // Title added using getPlaceholder method
            titleAdded = true;
          }
        } catch (e) {
          Logger.log('Error getting title placeholder: ' + e.message);
        }
      }
      
      // If title still wasn't added, use the first text box
      if (!titleAdded) {
        for (let j = 0; j < shapes.length; j++) {
          const shape = shapes[j];
          try {
            if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
              shape.getText().setText(info.title);
              // Title added using first text box
              titleAdded = true;
              break;
            }
          } catch (e) {
            Logger.log('Error using text box for title: ' + e.message);
          }
        }
      }
      
      // Add body content if it exists for TITLE_AND_BODY slides
      if (info.layout === 'TITLE_AND_BODY' && info.bodyItems && info.bodyItems.length > 0) {
        let bodyContentAdded = false;
        
        // Look for BODY placeholder
        for (let j = 0; j < shapes.length; j++) {
          const shape = shapes[j];
          try {
            if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
              const textRange = shape.getText();
              textRange.clear();
              
              // Add each body item as a paragraph
              for (let k = 0; k < info.bodyItems.length; k++) {
                if (k === 0) {
                  textRange.setText(info.bodyItems[k]);
                } else {
                  textRange.appendParagraph(info.bodyItems[k]);
                }
              }
              
              // Body content added using BODY placeholder
              bodyContentAdded = true;
              break;
            }
          } catch (e) {
            Logger.log('Error checking for BODY placeholder: ' + e.message);
          }
        }
        
        // If body content wasn't added, try another approach
        if (!bodyContentAdded) {
          try {
            const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
            if (bodyShape) {
              const textRange = bodyShape.getText();
              textRange.clear();
              
              // Add each body item as a paragraph
              for (let k = 0; k < info.bodyItems.length; k++) {
                if (k === 0) {
                  textRange.setText(info.bodyItems[k]);
                } else {
                  textRange.appendParagraph(info.bodyItems[k]);
                }
              }
              
              // Body content added using getPlaceholder method
              bodyContentAdded = true;
            }
          } catch (e) {
            Logger.log('Error getting body placeholder: ' + e.message);
          }
        }
        
        // If body content still wasn't added, find a suitable text box or create one
        if (!bodyContentAdded) {
          // Look for a text box that's not the title
          let textBoxFound = false;
          for (let j = 0; j < shapes.length; j++) {
            const shape = shapes[j];
            try {
              if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX && 
                  shape.getText().asString() !== info.title) {
                const textRange = shape.getText();
                textRange.clear();
                
                // Add each body item as a paragraph
                for (let k = 0; k < info.bodyItems.length; k++) {
                  if (k === 0) {
                    textRange.setText(info.bodyItems[k]);
                  } else {
                    textRange.appendParagraph(info.bodyItems[k]);
                  }
                }
                
                // Body content added using existing text box
                textBoxFound = true;
                bodyContentAdded = true;
                break;
              }
            } catch (e) {
              Logger.log('Error using text box for body: ' + e.message);
            }
          }
          
          // If no suitable text box was found, create one
          if (!textBoxFound) {
            try {
              const slideWidth = slide.getWidth();
              const slideHeight = slide.getHeight();
              
              const textBox = slide.insertTextBox(
                slideWidth * 0.1,  // Left position
                slideHeight * 0.3, // Top position
                slideWidth * 0.8,  // Width
                slideHeight * 0.6  // Height
              );
              
              const textRange = textBox.getText();
              
              // Add each body item as a paragraph
              for (let k = 0; k < info.bodyItems.length; k++) {
                if (k === 0) {
                  textRange.setText(info.bodyItems[k]);
                } else {
                  textRange.appendParagraph(info.bodyItems[k]);
                }
              }
              
              // Body content added using new text box
              bodyContentAdded = true;
            } catch (e) {
              Logger.log('Error creating new text box: ' + e.message);
            }
          }
        }
      }
      
      // Content added successfully
    }
    
    // Step 5: Apply list formatting to all TITLE_AND_BODY slides (only to body content, not titles)
    for (let i = 0; i < createdSlides.length; i++) {
      const slideObj = createdSlides[i];
      if (slideObj.info.layout === 'TITLE_AND_BODY' && slideObj.info.bodyItems.length > 0) {
        try {
          const shapes = slideObj.slide.getShapes();
          let bodyFormattingApplied = false;
          
          // Look for BODY placeholder
          for (let j = 0; j < shapes.length; j++) {
            const shape = shapes[j];
            try {
              if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
                shape.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
                // Bullet formatting applied to BODY placeholder
                bodyFormattingApplied = true;
                break;
              }
            } catch (e) {
              Logger.log('Error checking placeholder type for bullet formatting: ' + e.message);
            }
          }
          
          // If body formatting wasn't applied, try another approach
          if (!bodyFormattingApplied) {
            try {
              const bodyShape = slideObj.slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
              if (bodyShape) {
                bodyShape.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
                // Bullet formatting applied using getPlaceholder
                bodyFormattingApplied = true;
              }
            } catch (e) {
              Logger.log('Error getting body placeholder for bullet formatting: ' + e.message);
            }
          }
          
          // If body formatting still wasn't applied, try to find text boxes that aren't the title
          if (!bodyFormattingApplied) {
            for (let j = 0; j < shapes.length; j++) {
              const shape = shapes[j];
              try {
                if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
                  const text = shape.getText().asString().trim();
                  // Skip if this is the title text box
                  if (text !== '' && text !== slideObj.info.title) {
                    shape.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
                    // Bullet formatting applied to text box
                    bodyFormattingApplied = true;
                    break;
                  }
                }
              } catch (e) {
                Logger.log('Error applying bullet formatting to text box: ' + e.message);
              }
            }
          }
          
          // If still no formatting applied, try a fallback approach - manually add bullet points
          if (!bodyFormattingApplied) {
            for (let j = 0; j < shapes.length; j++) {
              const shape = shapes[j];
              try {
                if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
                  const textRange = shape.getText();
                  const text = textRange.asString().trim();
                  
                  // Skip if this is the title text box
                  if (text !== '' && text !== slideObj.info.title) {
                    // Clear the text box
                    textRange.clear();
                    
                    // Add each line with a bullet point
                    const lines = text.split('\n');
                    for (let k = 0; k < lines.length; k++) {
                      const line = lines[k].trim();
                      if (line !== '') {
                        if (k === 0) {
                          textRange.setText('\u2022 ' + line);
                        } else {
                          textRange.appendParagraph('\u2022 ' + line);
                        }
                      }
                    }
                    
                    // Manual bullet points applied
                    bodyFormattingApplied = true;
                    break;
                  }
                }
              } catch (e) {
                Logger.log('Error applying manual bullet points: ' + e.message);
              }
            }
          }
        } catch (e) {
          Logger.log('Error applying bullet formatting to slide ' + (i+1) + ': ' + e.message);
        }
      }
    }
    
    // Step 6: Apply markdown bold formatting to all slides
    applyMarkdownBoldToSlides(createdSlides.map(obj => obj.slide));
    
    return true;
  } catch (error) {
    console.error('Error converting markdown to slides: ' + error.message);
    return false;
  }
}

/**
 * Parse markdown text into a structured slide format
 * @param {string} markdownText - The markdown text to parse
 * @return {Array} Array of slide objects with layout, title, and bodyItems
 */
function parseMarkdownToStructure(markdownText) {
  try {
    const lines = markdownText.split('\n');
    const slideStructure = [];
    let currentSlide = null;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // Skip empty lines
      if (line === '') continue;
      
      // Check for horizontal rule (---) as slide separator
      if (line === '---') {
        // Skip the separator line - don't add it to any slide content
        continue;
      }
      // Check for H1 heading (# Heading)
      else if (line.startsWith('# ')) {
        // Create a new SECTION_HEADER slide
        currentSlide = {
          layout: 'SECTION_HEADER',
          title: line.substring(2).trim(),
          bodyItems: []
        };
        slideStructure.push(currentSlide);
      }
      // Check for H2 heading (## Heading)
      else if (line.startsWith('## ')) {
        // Create a new TITLE_AND_BODY slide
        currentSlide = {
          layout: 'TITLE_AND_BODY',
          title: line.substring(3).trim(),
          bodyItems: []
        };
        slideStructure.push(currentSlide);
      }
      // Add content to current slide if it's a TITLE_AND_BODY
      else if (currentSlide && currentSlide.layout === 'TITLE_AND_BODY') {
        // Process list items and regular text
        let content = line;
        
        // Remove list markers if present
        if (line.startsWith('- ')) {
          content = line.substring(2).trim();
        } else if (line.startsWith('* ')) {
          content = line.substring(2).trim();
        } else if (/^\d+\.\s/.test(line)) {
          content = line.substring(line.indexOf('.') + 1).trim();
        }
        
        currentSlide.bodyItems.push(content);
      }
    }
    
    return slideStructure;
  } catch (error) {
    console.error('Error parsing markdown: ' + error.message);
    return [];
  }
}

/**
 * Determines the index where new slides should be inserted
 * @param {Presentation} presentation - The active presentation
 * @return {number} - The index to insert slides at
 */
function getInsertIndex(presentation) {
  try {
    const selection = presentation.getSelection();
    
    if (selection) {
      const currentPage = selection.getCurrentPage();
      if (currentPage) {
        // Find the index of the current slide
        const slides = presentation.getSlides();
        for (let i = 0; i < slides.length; i++) {
          if (slides[i].getObjectId() === currentPage.getObjectId()) {
            // Insert after the current slide
            return i + 1;
          }
        }
      }
    }
    
    // Default to the end of the presentation if we can't determine the current slide
    return presentation.getSlides().length;
  } catch (error) {
    // Default to the end of the presentation
    return presentation.getSlides().length;
  }
}

/**
 * Add content to a slide as bullet points
 * @param {Slide} slide - The slide to add content to
 * @param {string[]} contentItems - Array of content items
 */
function addContentToSlide(slide, contentItems) {
  try {
    Logger.log('Adding content to slide: ' + slide.getObjectId());
    
    // First try to get the body placeholder
    const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    
    if (bodyShape) {
      // Use the body placeholder
      const textRange = bodyShape.getText();
      textRange.clear();
      
      // Add each content item
      for (let i = 0; i < contentItems.length; i++) {        
        if (i > 0) {
          textRange.appendParagraph(contentItems[i]);
        } else {
          textRange.setText(contentItems[i]);
        }
      }
    } else {
      // Fallback: Create a text box
      const slideWidth = slide.getWidth();
      const slideHeight = slide.getHeight();
      
      const textBox = slide.insertTextBox(
        slideWidth * 0.1,  // Left position
        slideHeight * 0.3, // Top position
        slideWidth * 0.8,  // Width
        slideHeight * 0.6  // Height
      );
      
      const textBoxText = textBox.getText();
      
      // Add each content item with a bullet point
      for (let i = 0; i < contentItems.length; i++) {
        const bulletItem = '• ' + contentItems[i];
        if (i > 0) {
          textBoxText.appendParagraph(bulletItem);
        } else {
          textBoxText.setText(bulletItem);
        }
      }
    }
    
    Logger.log('Successfully added content to slide');
  } catch (error) {
    Logger.log('Error adding content to slide: ' + error.message);
    // Continue execution
  }
}

// Debug function removed to improve performance

/**
 * Apply list style to all slides that have been created
 * @param {Array} slides - Array of slides to apply list style to
 */
function applyListStyleToSlides(slides) {
  for (let i = 0; i < slides.length; i++) {
    try {
      const slide = slides[i];
      const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
      
      if (bodyShape) {
        const textRange = bodyShape.getText();
        if (textRange.asString().trim() !== '') {
          textRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
        }
      }
    } catch (error) {
      // Continue with next slide
    }
  }
}

/**
 * Apply markdown bold formatting to text enclosed in double asterisks (**text**)
 * in all slides
 * @param {Array} slides - Array of slides to apply formatting to
 */
function applyMarkdownBoldToSlides(slides) {
  for (let i = 0; i < slides.length; i++) {
    try {
      const slide = slides[i];
      const shapes = slide.getShapes();
      
      // Process each shape in the slide
      for (let j = 0; j < shapes.length; j++) {
        const shape = shapes[j];
        
        try {
          // Only process text boxes and placeholders
          if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX ||
              shape.getPlaceholderType) {
            
            const textRange = shape.getText();
            const originalText = textRange.asString();
            
            // Find all **text** format matches
            const matches = [...originalText.matchAll(/\*\*(.+?)\*\*/g)];
            
            if (matches.length === 0) {
              continue; // No markdown bold formatting found in this element
            }
            
            let newText = '';
            let lastIndex = 0;
            const formattingRanges = [];
            
            matches.forEach(match => {
              const matchStart = match.index;
              const matchEnd = match.index + match[0].length;
              const content = match[1];
              
              // Add text before the match
              newText += originalText.substring(lastIndex, matchStart);
              
              // Record the position of formatted text in the new text
              const formatStart = newText.length;
              newText += content;
              const formatEnd = newText.length;
              
              // Store range (end is exclusive)
              formattingRanges.push({ start: formatStart, end: formatEnd });
              
              lastIndex = matchEnd;
            });
            
            // Add remaining original text
            newText += originalText.substring(lastIndex);
            
            // Replace text
            textRange.setText(newText);
            
            // Apply formatting
            formattingRanges.forEach(({ start, end }) => {
              const range = textRange.getRange(start, end);
              range.getTextStyle().setBold(true);
              
              // Check if main_color is defined in the global scope
              try {
                if (typeof main_color !== 'undefined') {
                  range.getTextStyle().setForegroundColor(main_color);
                }
              } catch (e) {
                // If main_color is not defined, we just skip setting the color
                Logger.log("Note: main_color not defined, skipping color formatting");
              }
            });
          }
        } catch (e) {
          Logger.log('Error processing shape for bold formatting: ' + e.message);
          // Continue with next shape
        }
      }
      
      Logger.log('Applied markdown bold formatting to slide ' + (i+1));
    } catch (error) {
      Logger.log('Error applying markdown bold formatting to slide ' + (i+1) + ': ' + error.message);
      // Continue with next slide
    }
  }
}

/**
 * Registers the md2slides utility in the menu
 */
function registerMd2SlidesMenu() {
  const ui = SlidesApp.getUi();
  ui.createMenu('Lizard Utilities')
    .addItem('Markdown to Slides', 'showMarkdownDialog')
    .addToUi();
}
