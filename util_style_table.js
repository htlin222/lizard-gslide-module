/**
 * Quickly styles a selected table with alternating row colors and header styling
 */
function fastStyleSelectedTable() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const tableCellRange = selection.getTableCellRange();

  if (!tableCellRange) {
    SlidesApp.getUi().alert("Please select a table cell.");
    return;
  }

  const table = tableCellRange.getTableCells()[0].getParentTable();
  const numRows = table.getNumRows();
  const numCols = table.getNumColumns();

  const white = "#FFFFFF";
  const blue = main_color;
  const gray = "#AAAAAA";
  const black = "#000000";
  const weight = 3;
  
  // Default border settings
  const borderColor = white; // White border
  const borderWidth = 3;     // 3pt width

  for (let r = 0; r < numRows; r++) {
    const isHeader = r === 0;
    const isEven = r % 2 === 0;

    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(r, c);
      if (!cell) continue;

      // Set background color and text style based on row
      const text = cell.getText();
      const textStyle = text.getTextStyle();

      if (isHeader) {
        cell.getFill().setSolidFill(blue);
        textStyle.setForegroundColor(white).setBold(true);
      } else if (isEven) {
        cell.getFill().setSolidFill(white);
        textStyle.setForegroundColor(black).setBold(false);
      } else {
        cell.getFill().setSolidFill(gray);
        textStyle.setForegroundColor(black).setBold(false);
      }
    }
  }
  
  // Apply borders using the Advanced Slides API
  applyTableBorders(table, borderColor, borderWidth);
}

/**
 * Applies borders to a table using the Slides Advanced Service API
 * @param {Table} table - The table to apply borders to
 * @param {string} borderColor - The border color in hex format (e.g., '#000000')
 * @param {number} borderWeight - The border weight in points
 */
function applyTableBorders(table, borderColor, borderWeight) {
  try {
    // Convert hex color to RGB components
    const r = parseInt(borderColor.substring(1, 3), 16) / 255;
    const g = parseInt(borderColor.substring(3, 5), 16) / 255;
    const b = parseInt(borderColor.substring(5, 7), 16) / 255;
    
    const presentationId = SlidesApp.getActivePresentation().getId();
    const tableId = table.getObjectId();
    
    // Create requests for different border positions
    const borderPositions = ['ALL']; // Can also use 'INNER', 'OUTER', 'INNER_HORIZONTAL', 'INNER_VERTICAL'
    
    const requests = borderPositions.map(position => ({
      "updateTableBorderProperties": {
        "objectId": tableId,
        "borderPosition": position,
        "tableBorderProperties": {
          "tableBorderFill": {
            "solidFill": {
              "color": {
                "rgbColor": {
                  "red": r,
                  "green": g,
                  "blue": b
                }
              }
            }
          },
          "weight": {
            "magnitude": borderWeight,
            "unit": "PT" // Points
          },
          "dashStyle": "SOLID"
        },
        "fields": "tableBorderFill,weight,dashStyle"
      }
    }));
    
    // Execute the batch update
    Slides.Presentations.batchUpdate({"requests": requests}, presentationId);
    console.log("Table borders applied successfully");
  } catch (e) {
    console.error("Error applying table borders: " + e.message);
  }
}

/**
 * Applies a complete table styling with borders and custom colors
 */
function styleTableWithBorders() {
  const selection = SlidesApp.getActivePresentation().getSelection();
  const pageElementRange = selection.getPageElementRange();
  
  if (!pageElementRange) {
    SlidesApp.getUi().alert("Please select a table.");
    return;
  }
  
  const pageElements = pageElementRange.getPageElements();
  
  if (pageElements.length === 0) {
    SlidesApp.getUi().alert("No elements selected.");
    return;
  }
  
  const element = pageElements[0];
  if (element.getPageElementType() !== SlidesApp.PageElementType.TABLE) {
    SlidesApp.getUi().alert("Please select a table.");
    return;
  }
  
  const table = element.asTable();
  const numRows = table.getNumRows();
  const numCols = table.getNumColumns();
  
  // Define colors
  const headerBg = "#1E88E5";  // Blue
  const headerText = "#FFFFFF"; // White
  const evenRowBg = "#FFFFFF";  // White
  const oddRowBg = "#F5F5F5";   // Light gray
  const textColor = "#212121";  // Dark gray
  const borderColor = "#616161"; // Medium gray
  const borderWeight = 1;       // 1pt
  
  // Style cells
  for (let r = 0; r < numRows; r++) {
    const isHeader = r === 0;
    const isEven = r % 2 === 0;
    
    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(r, c);
      const text = cell.getText();
      const textStyle = text.getTextStyle();
      
      if (isHeader) {
        cell.getFill().setSolidFill(headerBg);
        textStyle.setForegroundColor(headerText).setBold(true);
      } else if (isEven) {
        cell.getFill().setSolidFill(evenRowBg);
        textStyle.setForegroundColor(textColor).setBold(false);
      } else {
        cell.getFill().setSolidFill(oddRowBg);
        textStyle.setForegroundColor(textColor).setBold(false);
      }
    }
  }
  
  // Apply borders
  applyTableBorders(table, borderColor, borderWeight);
}