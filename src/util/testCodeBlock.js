/**
 * 測試 Code Block 功能
 *
 * 這個函數用來測試在當前投影片上創建 code block 是否正常運作
 */

/**
 * 測試在當前投影片上創建一個簡單的 code block
 */
function testCreateCodeBlock() {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const slides = presentation.getSlides();

		if (slides.length === 0) {
			console.log("No slides found in presentation");
			return false;
		}

		// 使用當前選取的投影片，或第一張投影片
		const selection = presentation.getSelection();
		let slide;

		if (selection && selection.getCurrentPage()) {
			slide = selection.getCurrentPage();
		} else {
			slide = slides[0];
		}

		console.log("Creating test code block on slide...");

		// 獲取投影片尺寸
		const slideWidth = slide.getPageWidth();
		const slideHeight = slide.getPageHeight();

		console.log(`Slide dimensions: ${slideWidth} x ${slideHeight}`);

		// 創建一個簡單的矩形 shape
		const testShape = slide.insertShape(
			SlidesApp.ShapeType.RECTANGLE,
			50, // x position
			200, // y position
			300, // width
			100, // height
		);

		console.log("Shape created, setting content...");

		// 設定文字內容
		const textRange = testShape.getText();
		textRange.setText('print("hello world")');

		// 設定樣式
		const textStyle = textRange.getTextStyle();
		textStyle.setFontSize(12);
		textStyle.setFontFamily("Courier New");
		textStyle.setForegroundColor("#000000");

		// 設定背景和邊框
		testShape.getFill().setSolidFill("#f5f5f5");
		testShape.getBorder().setWeight(1);
		testShape.getBorder().getLineFill().setSolidFill("#cccccc");

		console.log("Test code block created successfully!");
		return true;
	} catch (error) {
		console.error(`Error creating test code block: ${error.message}`);
		console.error(`Error stack: ${error.stack}`);
		return false;
	}
}

/**
 * 測試解析包含 code block 的 markdown
 */
function testMarkdownParsing() {
	const testMarkdown = `## 測試投影片

- 第一個項目
- 第二個項目

\`\`\`r
print("hello world")
x <- 1:10
\`\`\`

> 這是講者備註`;

	console.log("Testing markdown parsing...");

	try {
		const slideStructure = parseMarkdownToStructure(testMarkdown);
		console.log("Parsed structure:", JSON.stringify(slideStructure, null, 2));

		if (slideStructure.length > 0 && slideStructure[0].codeBlocks) {
			console.log("Code blocks found:", slideStructure[0].codeBlocks);
			return true;
		} else {
			console.log("No code blocks found in parsed structure");
			return false;
		}
	} catch (error) {
		console.error(`Error parsing markdown: ${error.message}`);
		return false;
	}
}

/**
 * 完整測試 markdown to slides 轉換
 */
function testFullMarkdownConversion() {
	const testMarkdown = `## 測試程式碼區塊

- 這是第一個項目
- 這是第二個項目

\`\`\`r
print("hello world")
x <- 1:10
mean(x)
\`\`\`

> 這是一個測試投影片的講者備註`;

	console.log("Starting full markdown conversion test...");

	try {
		const result = convertMarkdownToSlides(testMarkdown);
		console.log(`Conversion result: ${result}`);
		return result;
	} catch (error) {
		console.error(`Error in full conversion: ${error.message}`);
		return false;
	}
}
