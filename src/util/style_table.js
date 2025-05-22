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
  const gray = "#CCCCCC";
  const black = "#000000";  
  // Default border settings
  const borderColor = main_color; // White border
  const borderWidth = 0.3;     // 3pt width

  for (let r = 0; r < numRows; r++) {
    const isHeader = r === 0;
    const isEven = r % 2 === 0;

    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(r, c);
      if (!cell) continue;

      // Set background color based on row
      if (isHeader) {
        cell.getFill().setSolidFill(blue);
      } else if (isEven) {
        cell.getFill().setSolidFill(gray);
      } else {
        cell.getFill().setSolidFill(white);
      }
      
      // Only set text style if there is text content
      try {
        const text = cell.getText();
        if (text && text.asString().trim() !== "") {
          const textStyle = text.getTextStyle();
          if (isHeader) {
            textStyle.setForegroundColor(white).setBold(true);
          } else {
            textStyle.setForegroundColor(black).setBold(false);
          }
        }
        
        // Set content alignment to center both horizontally and vertically
        // MIDDLE only centers vertically, we need to set both horizontal and vertical alignment
        cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        
        // We also need to set paragraph alignment for horizontal centering
        try {
          const textRange = cell.getText();
          const paragraphStyle = textRange.getParagraphStyle();
          paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        } catch (alignError) {
          console.log("Could not set paragraph alignment for cell at row " + r + ", column " + c);
        }
      } catch (e) {
        console.log("Skipping text styling for empty cell at row " + r + ", column " + c);
      }
    }
  }
  
  // Apply borders using the Advanced Slides API - different border for header row
  applyTableBordersWithHeaderRow(table, white, borderColor, borderWidth);
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
 * Applies borders to a table with special styling for the header row
 * @param {Table} table - The table to apply borders to
 * @param {string} headerBorderColor - The border color for the header row in hex format
 * @param {string} bodyBorderColor - The border color for the body rows in hex format
 * @param {number} borderWeight - The border weight in points
 */
function applyTableBordersWithHeaderRow(table, headerBorderColor, bodyBorderColor, borderWeight) {
  try {
    const presentationId = SlidesApp.getActivePresentation().getId();
    const tableId = table.getObjectId();
    
    // Log the table ID to verify we're targeting the right table
    console.log("Applying borders to table ID: " + tableId);
    
    // Convert header border color to RGB
    const hr = parseInt(headerBorderColor.substring(1, 3), 16) / 255;
    const hg = parseInt(headerBorderColor.substring(3, 5), 16) / 255;
    const hb = parseInt(headerBorderColor.substring(5, 7), 16) / 255;
    
    // Convert body border color to RGB
    const br = parseInt(bodyBorderColor.substring(1, 3), 16) / 255;
    const bg = parseInt(bodyBorderColor.substring(3, 5), 16) / 255;
    const bb = parseInt(bodyBorderColor.substring(5, 7), 16) / 255;
    
    // Log the colors to verify conversion
    console.log("Header border RGB: " + hr + ", " + hg + ", " + hb);
    console.log("Body border RGB: " + br + ", " + bg + ", " + bb);
    
    // Try a simpler approach - just set all borders to the body color
    const requests = [{
      "updateTableBorderProperties": {
        "objectId": tableId,
        "borderPosition": "ALL",
        "tableBorderProperties": {
          "tableBorderFill": {
            "solidFill": {
              "color": {
                "rgbColor": {
                  "red": br,
                  "green": bg,
                  "blue": bb
                }
              }
            }
          },
          "weight": {
            "magnitude": borderWeight,
            "unit": "PT"
          },
          "dashStyle": "SOLID"
        },
        "fields": "tableBorderFill,weight,dashStyle"
      }
    }];
    
    // Execute the batch update
    const response = Slides.Presentations.batchUpdate({"requests": requests}, presentationId);
    console.log("Table borders applied successfully");
    console.log("API Response: " + JSON.stringify(response));
    
    // Now try a second approach for the header row - using a different method
    // Instead of using tableRange which might not be working as expected,
    // we'll try to directly style the first row using a different approach
    
    // First, get the current slide
    const currentSlide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
    if (!currentSlide) {
      console.log("Could not get current slide");
      return;
    }
    
    // According to the API documentation, valid border positions are: ALL, BOTTOM, INNER, INNER_HORIZONTAL, INNER_VERTICAL, LEFT, OUTER, RIGHT, TOP
    try {
      // We only want to change the vertical borders of the header row to white
      // Since the API doesn't allow targeting specific row borders directly,
      // we'll use a simpler approach - just make the vertical borders white
      
      // Apply white color to all INNER_VERTICAL borders
      const verticalRequests = [{
        "updateTableBorderProperties": {
          "objectId": tableId,
          "borderPosition": "INNER_VERTICAL", // This affects all vertical borders between columns
          "tableBorderProperties": {
            "tableBorderFill": {
              "solidFill": {
                "color": {
                  "rgbColor": {
                    "red": 1.0,  // Pure white
                    "green": 1.0,
                    "blue": 1.0
                  }
                }
              }
            },
            "weight": {
              "magnitude": borderWeight * 1.5, // Slightly thicker for visibility
              "unit": "PT"
            },
            "dashStyle": "SOLID"
          },
          "fields": "tableBorderFill,weight,dashStyle"
        }
      }];
      
      // Also apply white to LEFT and RIGHT borders for completeness
      const leftRightRequests = [
        {
          "updateTableBorderProperties": {
            "objectId": tableId,
            "borderPosition": "LEFT",
            "tableBorderProperties": {
              "tableBorderFill": {
                "solidFill": {
                  "color": {
                    "rgbColor": {
                      "red": 1.0,
                      "green": 1.0,
                      "blue": 1.0
                    }
                  }
                }
              },
              "weight": {
                "magnitude": borderWeight * 1.5,
                "unit": "PT"
              },
              "dashStyle": "SOLID"
            },
            "fields": "tableBorderFill,weight,dashStyle"
          }
        },
        {
          "updateTableBorderProperties": {
            "objectId": tableId,
            "borderPosition": "RIGHT",
            "tableBorderProperties": {
              "tableBorderFill": {
                "solidFill": {
                  "color": {
                    "rgbColor": {
                      "red": 1.0,
                      "green": 1.0,
                      "blue": 1.0
                    }
                  }
                }
              },
              "weight": {
                "magnitude": borderWeight * 1.5,
                "unit": "PT"
              },
              "dashStyle": "SOLID"
            },
            "fields": "tableBorderFill,weight,dashStyle"
          }
        }
      ];
      
      // Execute the requests
      const verticalResponse = Slides.Presentations.batchUpdate({"requests": verticalRequests}, presentationId);
      console.log("Vertical borders applied successfully");
      
      const leftRightResponse = Slides.Presentations.batchUpdate({"requests": leftRightRequests}, presentationId);
      console.log("Left and right borders applied successfully");
      
      // Unfortunately, the API doesn't allow us to target just the vertical borders of the header row
      // This approach changes all vertical borders to white, which should at least make them visible
      // against the header background
    } catch (headerError) {
      console.error("Error applying header row borders: " + headerError.message);
    }
    
  } catch (e) {
    console.error("Error applying table borders with header styling: " + e.message);
    // Try a fallback approach
    try {
      // Simple fallback - just apply a single border style to the whole table
      applyTableBorders(table, bodyBorderColor, borderWeight);
    } catch (fallbackError) {
      console.error("Fallback border styling also failed: " + fallbackError.message);
    }
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