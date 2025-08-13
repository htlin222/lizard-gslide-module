// Utility for managing chains of titles in Google Slides
function copyPreviousTitleText() {
	const presentation = SlidesApp.getActivePresentation();
	const slides = presentation.getSlides();
	const currentSlide = presentation.getSelection().getCurrentPage();
	const currentIndex = slides.findIndex(
		(slide) => slide.getObjectId() === currentSlide.getObjectId(),
	);

	if (currentIndex <= 0) {
		return;
	}

	// === Step 1: Remove existing 'PREVIOUS_TITLE' from current slide ===
	const currentElements = currentSlide.getPageElements();
	for (const element of currentElements) {
		if (
			element.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
			element.getTitle() === "PREVIOUS_TITLE"
		) {
			element.remove();
		}
	}

	// === Step 2: Try to get text from previous slide ===
	const previousSlide = slides[currentIndex - 1];
	const pageElements = previousSlide.getPageElements();
	let titleText = "";

	// Try to find shape with title "PREVIOUS_TITLE"
	for (const element of pageElements) {
		if (
			element.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
			element.getTitle() === "PREVIOUS_TITLE"
		) {
			const text = element.asShape().getText().asString().trim();
			if (text.length > 0) {
				titleText = text;
				break;
			}
		}
	}

	// Fallback: first non-empty text shape
	if (!titleText) {
		for (const element of pageElements) {
			if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
				const shape = element.asShape();
				const text = shape.getText().asString().trim();
				if (text.length > 0) {
					titleText = text;
					break;
				}
			}
		}
	}

	if (!titleText) {
		return;
	}
	// === Step 3: Insert new text box on current slide ===
	const newShape = currentSlide.insertTextBox(titleText, 24, 18, 500, 25);
	newShape.setTitle("PREVIOUS_TITLE");

	// === Step 4: Style it ===
	const textRange = newShape.getText();
	textRange.getTextStyle().setFontSize(12);
	textRange.getTextStyle().setForegroundColor("#888888");
}

function createNextSlideWithCurrentTitle() {
	const presentation = SlidesApp.getActivePresentation();
	const slides = presentation.getSlides();
	const currentPageId = presentation
		.getSelection()
		.getCurrentPage()
		.getObjectId();

	// Find current slide index
	const currentIndex = slides.findIndex(
		(slide) => slide.getObjectId() === currentPageId,
	);
	if (currentIndex < 0 || currentIndex >= slides.length) {
		Logger.log("Current slide not found.");
		return;
	}

	const currentSlide = slides[currentIndex];

	// ✅ Fix: Get layout using indexed slide reference
	const layout = slides[currentIndex].getLayout();
	const layoutName = layout ? layout.getLayoutName() : "No layout";
	Logger.log("Current slide layout: " + layoutName);

	// Create new slide after current with same layout
	const newSlide = presentation.insertSlide(currentIndex + 1, layout);
	Logger.log("Inserted new slide after current one.");

	// Try to get text from shape titled 'PREVIOUS_TITLE'
	const pageElements = currentSlide.getPageElements();
	let titleText = "";

	for (const element of pageElements) {
		if (
			element.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
			element.getTitle() === "PREVIOUS_TITLE"
		) {
			const text = element.asShape().getText().asString().trim();
			if (text.length > 0) {
				titleText = text;
				Logger.log("Found text via title 'PREVIOUS_TITLE': " + titleText);
				break;
			}
		}
	}
	// Fallback: use first non-empty text box
	if (!titleText) {
		for (const element of pageElements) {
			if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
				const shape = element.asShape();
				const text = shape.getText().asString().trim();
				if (text.length > 0) {
					titleText = text;
					Logger.log("Fallback: Found first non-empty text: " + titleText);
					break;
				}
			}
		}
	}
	if (!titleText) {
		Logger.log("No suitable text found on current slide.");
		return;
	}
	insertStyledTitleBox(newSlide, titleText, "#888888");
}

function insertStyledTitleBox(slide, titleText, color) {
	// Insert styled text box at fixed position and size
	const newShape = slide.insertTextBox(titleText, 24, 18, 200, 25);
	newShape.setTitle("PREVIOUS_TITLE");
	// Apply font styling
	const textRange = newShape.getText();
	textRange.getTextStyle().setFontSize(12);
	textRange.getTextStyle().setForegroundColor(color);
}
