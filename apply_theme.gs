// This file previously contained the onOpen function and theme-related functions
// These have been moved to config.gs to consolidate menu functionality
// This file is kept for backward compatibility


/**
 * Apply theme from a source presentation to the current presentation
 * This preserves the content of the current presentation while applying the theme/styles from the source
 */
function applyThemeToCurrentPresentation() {
    // Add debugging information
    Logger.log('Starting theme application process...');
    
    // Source presentation with the desired theme/styles
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