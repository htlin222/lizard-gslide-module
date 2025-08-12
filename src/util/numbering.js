function addNextNumberCircle() {
	const presentation = SlidesApp.getActivePresentation();
	const slide = presentation.getSelection().getCurrentPage(); // active slide

	const pageElements = slide.getPageElements();
	const numberElements = [];

	// Step 1: Collect all elements with title "NUMBER"
	for (let i = 0; i < pageElements.length; i++) {
		const el = pageElements[i];
		if (el.getTitle && el.getTitle() === "NUMBER") {
			numberElements.push(el);
		}
	}

	// Step 2: Extract numbers from text content
	const numbers = numberElements
		.map((el) => {
			try {
				const text = el.asShape().getText().asString().trim();
				return parseInt(text, 10);
			} catch (e) {
				return null;
			}
		})
		.filter((n) => !isNaN(n));

	const maxNum = numbers.length > 0 ? Math.max(...numbers) : 0;
	const newNumber = maxNum + 1;

	// Step 3: Insert circle
	const shape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, 25, 25, 36, 36);
	shape.getText().setText(newNumber.toString());

	// Step 4: Styling
	shape.getFill().setSolidFill(main_color); // blue
	shape.getText().getTextStyle().setForegroundColor(base_color); // white
	shape.getText().getTextStyle().setFontSize(18); // set font size to 18
	shape
		.getText()
		.getParagraphStyle()
		.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
	shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Step 5: Set title
	shape.setTitle("NUMBER");
}
