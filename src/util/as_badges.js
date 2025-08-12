// Utility for creating badges in Google Slides
function convertToBadges() {
	const presentation = SlidesApp.getActivePresentation();
	const slide = presentation.getSelection().getCurrentPage();
	const selection = presentation.getSelection();

	const shape = validateSelectedTextBox(selection);
	if (!shape) return;

	const text = shape.getText().asString().trim();
	shape.getParentGroup() ? shape.getParentGroup().remove() : shape.remove();

	const items = text
		.split(",")
		.map((item) => item.trim())
		.filter((item) => item.length > 0);
	items.reverse();
	// Constants
	const fontSize = 10;
	const padding = 2;
	const height = 14;
	const startXRight = 710; // Align to right
	const startY = 350;
	const gap = 3;
	const charWidth = 6.5;
	const minWidth = 40;

	let y = startY;

	items.forEach((item) => {
		const textWidth = Math.max(item.length * charWidth + padding * 2, minWidth);
		const x = startXRight - textWidth; // Align right

		const badge = slide.insertShape(
			SlidesApp.ShapeType.TEXT_BOX,
			x,
			y,
			textWidth,
			height,
		);
		badge.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

		// Set background color to blue
		badge.getFill().setSolidFill(main_color); // or use 'blue' named color

		const badgeText = badge.getText();
		badgeText.setText(item);

		// Set font size and font color to white
		badgeText
			.getTextStyle()
			.setFontSize(fontSize)
			.setForegroundColor(base_color); // white

		// Align text to the right
		badgeText
			.getParagraphStyle()
			.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

		y -= height + gap; // Move to next line
	});
}
// 🔧 Validation Helper Function
function validateSelectedTextBox(selection) {
	const pageElementRange = selection.getPageElementRange();

	if (!pageElementRange) {
		SlidesApp.getUi().alert("請選取一個非空的文字框。");
		return null;
	}

	const selectedElements = pageElementRange.getPageElements();

	if (selectedElements.length !== 1) {
		SlidesApp.getUi().alert("請選取一個文字框。");
		return null;
	}

	const element = selectedElements[0];

	if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
		SlidesApp.getUi().alert("選取的不是文字框。");
		return null;
	}

	const shape = element.asShape();
	const text = shape.getText().asString().trim();

	if (text === "") {
		SlidesApp.getUi().alert("文字框是空的。");
		return null;
	}

	return shape;
}
