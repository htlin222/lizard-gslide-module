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
  const blue = "#1E88E5";
  const gray = "#EEEEEE";
  const black = "#000000";
  const weight = 3;

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

      // Set all borders once
      const borders = [
        cell.getBorderTop(),
        cell.getBorderBottom(),
        cell.getBorderLeft(),
        cell.getBorderRight(),
      ];
      for (const border of borders) {
        border.setWeight(weight).getLineFill().setSolidFill(white);
      }
    }
  }
}