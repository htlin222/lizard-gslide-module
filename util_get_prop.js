/**
 * Utility functions to get and log properties of selected objects in Google Slides
 */

/**
 * Logs all accessible properties of the currently selected object in Google Slides
 * This function detects the type of selection and logs appropriate properties
 */
function logSelectedObjectProperties() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  
  Logger.log("Selection Type: " + selectionType);
  console.log("Selection Type: " + selectionType);
  
  // Different types of selections
  switch(selectionType) {
    case SlidesApp.SelectionType.TEXT:
      logTextSelectionProperties(selection);
      break;
      
    case SlidesApp.SelectionType.TABLE_CELL:
      logTableCellProperties(selection);
      break;
      
    case SlidesApp.SelectionType.PAGE_ELEMENT:
      logPageElementProperties(selection);
      break;
      
    case SlidesApp.SelectionType.CURRENT_PAGE:
      logCurrentPageProperties(selection);
      break;
      
    default:
      Logger.log("No supported selection type found or nothing is selected");
      console.log("No supported selection type found or nothing is selected");
  }
}

/**
 * Logs properties of a text selection
 * @param {Selection} selection - The current selection object
 */
function logTextSelectionProperties(selection) {
  const textRange = selection.getTextRange();
  const textStyle = textRange.getTextStyle();
  
  Logger.log("=== TEXT SELECTION PROPERTIES ===");
  Logger.log("Text content: " + textRange.getText());
  
  // Text style properties
  Logger.log("Bold: " + textStyle.isBold());
  Logger.log("Italic: " + textStyle.isItalic());
  Logger.log("Underline: " + textStyle.isUnderline());
  Logger.log("Strikethrough: " + textStyle.isStrikethrough());
  Logger.log("Font family: " + textStyle.getFontFamily());
  Logger.log("Font size: " + textStyle.getFontSize());
  
  // Try to get foreground color (may be null if mixed)
  try {
    const foregroundColor = textStyle.getForegroundColor();
    if (foregroundColor) {
      Logger.log("Foreground color: " + JSON.stringify(foregroundColor));
    }
  } catch (e) {
    Logger.log("Could not get foreground color: " + e.message);
  }
  
  // Get paragraph style if available
  try {
    const paragraphStyle = textRange.getParagraphStyle();
    Logger.log("Alignment: " + paragraphStyle.getAlignment());
    Logger.log("Line spacing: " + paragraphStyle.getLineSpacing());
    Logger.log("Space above: " + paragraphStyle.getSpaceAbove());
    Logger.log("Space below: " + paragraphStyle.getSpaceBelow());
  } catch (e) {
    Logger.log("Could not get paragraph style: " + e.message);
  }
  
  // Also log to console for Apps Script dashboard
  console.log("=== TEXT SELECTION PROPERTIES ===");
  console.log(JSON.stringify({
    content: textRange.getText(),
    isBold: textStyle.isBold(),
    isItalic: textStyle.isItalic(),
    isUnderline: textStyle.isUnderline(),
    fontFamily: textStyle.getFontFamily(),
    fontSize: textStyle.getFontSize()
  }, null, 2));
}

/**
 * Logs properties of a table cell selection
 * @param {Selection} selection - The current selection object
 */
function logTableCellProperties(selection) {
  const tableCellRange = selection.getTableCellRange();
  const cells = tableCellRange.getTableCells();
  
  Logger.log("=== TABLE CELL PROPERTIES ===");
  Logger.log("Number of selected cells: " + cells.length);
  
  if (cells.length > 0) {
    const cell = cells[0];
    const table = cell.getParentTable();
    
    Logger.log("Table dimensions: " + table.getNumRows() + " rows × " + table.getNumColumns() + " columns");
    Logger.log("Selected cell position: Row " + cell.getRowIndex() + ", Column " + cell.getColumnIndex());
    
    // Cell content
    Logger.log("Cell content: " + cell.getText().asString());
    
    // Cell appearance
    try {
      const fill = cell.getFill();
      if (fill.getSolidFill()) {
        const color = fill.getSolidFill().getColor();
        Logger.log("Cell background color: " + JSON.stringify(color));
      }
    } catch (e) {
      Logger.log("Could not get cell fill: " + e.message);
    }
    
    // Also log to console
    console.log("=== TABLE CELL PROPERTIES ===");
    console.log(JSON.stringify({
      tableSize: {
        rows: table.getNumRows(),
        columns: table.getNumColumns()
      },
      selectedCell: {
        row: cell.getRowIndex(),
        column: cell.getColumnIndex(),
        content: cell.getText().asString()
      }
    }, null, 2));
  }
}

/**
 * Logs properties of a page element selection (shapes, images, etc.)
 * @param {Selection} selection - The current selection object
 */
function logPageElementProperties(selection) {
  const pageElementRange = selection.getPageElementRange();
  const pageElements = pageElementRange.getPageElements();
  
  Logger.log("=== PAGE ELEMENT PROPERTIES ===");
  Logger.log("Number of selected elements: " + pageElements.length);
  
  if (pageElements.length > 0) {
    const element = pageElements[0];
    const elementType = element.getPageElementType();
    
    Logger.log("Element type: " + elementType);
    Logger.log("Element ID: " + element.getObjectId());
    Logger.log("Position: Left " + element.getLeft() + ", Top " + element.getTop());
    
    // Try to get size - might be null for some elements
    let width = "null";
    let height = "null";
    try {
      width = element.getWidth();
      height = element.getHeight();
    } catch (e) {}
    Logger.log("Size: Width " + width + ", Height " + height);
    Logger.log("Rotation: " + element.getRotation() + " degrees");
    
    // Log all available methods on the element
    const methods = [];
    for (let prop in element) {
      if (typeof element[prop] === 'function' && prop[0] !== '_') {
        methods.push(prop);
      }
    }
    Logger.log("Available methods: " + methods.join(", "));
    
    // Type-specific properties
    switch(elementType) {
      case SlidesApp.PageElementType.SHAPE:
        logShapeProperties(element.asShape());
        break;
      case SlidesApp.PageElementType.IMAGE:
        logImageProperties(element.asImage());
        break;
      case SlidesApp.PageElementType.TABLE:
        logTableProperties(element.asTable());
        break;
      case SlidesApp.PageElementType.GROUP:
        Logger.log("Group contains " + element.asGroup().getChildren().length + " elements");
        break;
    }
    
    // Also log to console
    console.log("=== PAGE ELEMENT PROPERTIES ===");
    console.log(JSON.stringify({
      type: elementType,
      id: element.getObjectId(),
      position: {
        left: element.getLeft(),
        top: element.getTop()
      },
      size: {
        width: width,
        height: height
      },
      rotation: element.getRotation(),
      availableMethods: methods
    }, null, 2));
  }
}

/**
 * Logs properties specific to shapes
 * @param {Shape} shape - The shape object
 */
function logShapeProperties(shape) {
  Logger.log("--- Shape Properties ---");
  Logger.log("Shape type: " + shape.getShapeType());
  
  // Fill
  try {
    const fill = shape.getFill();
    if (fill.getSolidFill()) {
      const color = fill.getSolidFill().getColor();
      Logger.log("Fill color: " + JSON.stringify(color));
    }
  } catch (e) {
    Logger.log("Could not get shape fill: " + e.message);
  }
  
  // Border
  try {
    const border = shape.getBorder();
    if (border) {
      Logger.log("Border weight: " + border.getWeight());
      Logger.log("Border dash style: " + border.getDashStyle());
    }
  } catch (e) {
    Logger.log("Could not get shape border: " + e.message);
  }
  
  // Text content if any
  if (shape.getText()) {
    Logger.log("Shape text: " + shape.getText().asString());
  }
}

/**
 * Logs properties specific to images
 * @param {Image} image - The image object
 */
function logImageProperties(image) {
  Logger.log("--- Image Properties ---");
  Logger.log("Content URL: " + image.getContentUrl());
  Logger.log("Source URL: " + image.getSourceUrl());
  
  // Try to get image properties that might not be available
  try {
    Logger.log("Brightness: " + image.getBrightness());
    Logger.log("Contrast: " + image.getContrast());
    Logger.log("Transparency: " + image.getTransparency());
  } catch (e) {
    Logger.log("Could not get some image properties: " + e.message);
  }
}

/**
 * Logs properties specific to tables
 * @param {Table} table - The table object
 */
function logTableProperties(table) {
  Logger.log("--- Table Properties ---");
  Logger.log("Table dimensions: " + table.getNumRows() + " rows × " + table.getNumColumns() + " columns");
  
  // Log table structure and content
  let tableData = [];
  for (let r = 0; r < Math.min(table.getNumRows(), 5); r++) { // Limit to first 5 rows
    let rowData = [];
    for (let c = 0; c < Math.min(table.getNumColumns(), 5); c++) { // Limit to first 5 columns
      try {
        const cell = table.getCell(r, c);
        const cellText = cell.getText().asString();
        
        // Try to get cell background color
        let bgColor = "unknown";
        try {
          const fill = cell.getFill();
          if (fill.getSolidFill()) {
            const color = fill.getSolidFill().getColor();
            if (color) {
              bgColor = JSON.stringify(color);
            }
          }
        } catch (e) {}
        
        rowData.push({
          text: cellText,
          bgColor: bgColor
        });
      } catch (e) {
        rowData.push({ error: e.message });
      }
    }
    tableData.push(rowData);
  }
  
  Logger.log("Table data sample: " + JSON.stringify(tableData, null, 2));
  
  // Try to get border properties
  try {
    const cell = table.getCell(0, 0);
    const border = cell.getBorder();
    if (border) {
      Logger.log("Border properties available: " + (border !== null));
    }
  } catch (e) {
    Logger.log("Could not get border properties: " + e.message);
  }
  
  // Log table object properties
  const tableProps = {};
  for (let prop in table) {
    if (typeof table[prop] === 'function') {
      tableProps[prop] = 'function';
    } else if (prop[0] !== '_') { // Skip private properties
      tableProps[prop] = table[prop];
    }
  }
  Logger.log("Available table properties: " + JSON.stringify(tableProps, null, 2));
}

/**
 * Logs properties of the current page
 * @param {Selection} selection - The current selection object
 */
function logCurrentPageProperties(selection) {
  const currentPage = selection.getCurrentPage();
  
  Logger.log("=== CURRENT PAGE PROPERTIES ===");
  Logger.log("Page ID: " + currentPage.getObjectId());
  Logger.log("Page index: " + currentPage.getObjectId());
  Logger.log("Page type: " + (currentPage.isGrouped() ? "Group" : "Regular"));
  Logger.log("Page elements count: " + currentPage.getPageElements().length);
  
  // Background
  try {
    const background = currentPage.getBackground();
    if (background.getSolidFill()) {
      const color = background.getSolidFill().getColor();
      Logger.log("Background color: " + JSON.stringify(color));
    }
  } catch (e) {
    Logger.log("Could not get page background: " + e.message);
  }
  
  // Also log to console
  console.log("=== CURRENT PAGE PROPERTIES ===");
  console.log(JSON.stringify({
    id: currentPage.getObjectId(),
    elementsCount: currentPage.getPageElements().length
  }, null, 2));
}


/**
 * Specifically logs the structure of a selected table with more detail
 */
function logSelectedTableStructure() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const selectionType = selection.getSelectionType();
  
  if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
    const pageElements = selection.getPageElementRange().getPageElements();
    if (pageElements.length > 0) {
      const element = pageElements[0];
      if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        const table = element.asTable();
        
        // Create a detailed structure representation
        Logger.log("=== DETAILED TABLE STRUCTURE ===");
        Logger.log("Table ID: " + element.getObjectId());
        Logger.log("Dimensions: " + table.getNumRows() + " rows × " + table.getNumColumns() + " columns");
        
        // Log all cells with content and formatting
        for (let r = 0; r < table.getNumRows(); r++) {
          for (let c = 0; c < table.getNumColumns(); c++) {
            try {
              const cell = table.getCell(r, c);
              const text = cell.getText().asString();
              
              if (text.trim() !== "") {
                Logger.log(`Cell [${r},${c}]: "${text}"`);
                
                // Try to get text style
                try {
                  const textStyle = cell.getText().getTextStyle();
                  Logger.log(`  - Bold: ${textStyle.isBold()}`);
                  Logger.log(`  - Font: ${textStyle.getFontFamily() || 'default'}`);
                  Logger.log(`  - Size: ${textStyle.getFontSize() || 'default'}`);
                } catch (e) {}
                
                // Try to get cell fill
                try {
                  const fill = cell.getFill();
                  if (fill.getSolidFill()) {
                    const color = fill.getSolidFill().getColor();
                    if (color) {
                      Logger.log(`  - Background: ${JSON.stringify(color)}`);
                    }
                  }
                } catch (e) {}
              }
            } catch (e) {
              Logger.log(`Error accessing cell [${r},${c}]: ${e.message}`);
            }
          }
        }
        
        return;
      }
    }
  }
  
  SlidesApp.getUi().alert("Please select a table first");
}


