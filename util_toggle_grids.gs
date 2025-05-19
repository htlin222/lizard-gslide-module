// Utility for toggling grid visibility in Google Slides
function toggleGrids() {
  var slide = SlidesApp
    .getActivePresentation()
    .getSelection()
    .getCurrentPage();

  // 1) Find existing GRID_* elements
  var grids = slide.getPageElements().filter(function(el) {
    var t = el.getTitle();
    return t && t.indexOf('GRID_') === 0;
  });
  
  if (grids.length) {
    // Remove them all → toggle off
    grids.forEach(function(el) { el.remove(); });
    return;
  }

  // 2) Otherwise, draw and title each line
  var width = 720, height = 405;
  var coords = [
    // [x1, y1, x2, y2]
    [180, 0,   180, height],  // GRID_1
    [360, 0,   360, height],  // GRID_2
    [540, 0,   540, height],  // GRID_3
    [0,   135, width, 135],   // GRID_4
    [0,   270, width, 270]    // GRID_5
  ];
  var lineColor = '#dddddd';
  var lineWeight = 0.5;

  coords.forEach(function(c, i) {
    var line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      c[0], c[1],
      c[2], c[3]
    );
    line.getLineFill().setSolidFill(lineColor);
    line.setWeight(lineWeight);
    line.setTitle('GRID_' + (i+1));
    line.sendToBack();
  });
}