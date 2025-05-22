// Utility for creating and managing boxes in Google Slides
function createStyledBoxesOnCurrentSlide() {
  const ui = SlidesApp.getUi();
  const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();
  const response = ui.prompt('Create Boxes', 'Enter the number of boxes to create:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const n = parseInt(response.getResponseText());

  if (isNaN(n) || n <= 0) {
    ui.alert('Invalid input. Please enter a positive number.');
    return;
  }

  const selection = SlidesApp.getActivePresentation().getSelection();
  const currentSlide = selection.getCurrentPage();

  if (!currentSlide) {
    ui.alert('No slide is currently selected.');
    return;
  }
  const boxWidth = 200;
  const boxHeight = 60;
  const verticalSpacing = 15;
  const startY = 100;

  for (let i = 0; i < n; i++) {
    const x = (slideWidth - boxWidth) / 2;
    const y = startY + i * (boxHeight + verticalSpacing);

    const box = currentSlide.insertTextBox('', x, y, boxWidth, boxHeight);
    box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); // Vertically center

    const textRange = box.getText();
    textRange.setText('Box ' + (i + 1));
    textRange.getTextStyle().setFontSize(20);
    textRange.getTextStyle().setForegroundColor('#FFFFFF');
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER); // Horizontally center

    box.getFill().setSolidFill(main_color); // Blue background
  }

  ui.alert(`${n} fully centered boxes created on the current slide.`);
}