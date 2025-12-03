/**
 * Test functions for H3 header functionality
 */

/**
 * Test H3 header parsing and parent title functionality
 */
function testH3HeaderParsing() {
	const testMarkdown = `# Section Title

## Main Topic
- First point
- Second point

### Subtopic A
- Detailed point 1
- Detailed point 2

### Subtopic B
- Another detailed point
- One more point

## Another Topic
- Different content

### Subtopic C
- Under different topic
- More content
`;

	console.log("Testing H3 header parsing...");

	try {
		const structure = parseMarkdownToStructure(testMarkdown);
		console.log("Parsed structure:", structure);

		// Check that we have the right number of slides
		console.log(`Total slides created: ${structure.length}`);

		// Check each slide
		structure.forEach((slide, index) => {
			console.log(`Slide ${index + 1}:`);
			console.log(`  Layout: ${slide.layout}`);
			console.log(`  Title: ${slide.title}`);
			console.log(`  Parent Title: ${slide.parentTitle || "None"}`);
			console.log(`  Body Items: ${slide.bodyItems.length}`);
		});

		// Verify specific expectations
		const h3Slides = structure.filter((slide) => slide.parentTitle);
		console.log(`H3 slides with parent titles: ${h3Slides.length}`);

		// Check specific parent title assignments
		const subtopicA = structure.find((slide) => slide.title === "Subtopic A");
		if (subtopicA && subtopicA.parentTitle === "Main Topic") {
			console.log("✅ Subtopic A correctly assigned parent title 'Main Topic'");
		} else {
			console.log("❌ Subtopic A parent title assignment failed");
		}

		const subtopicC = structure.find((slide) => slide.title === "Subtopic C");
		if (subtopicC && subtopicC.parentTitle === "Another Topic") {
			console.log(
				"✅ Subtopic C correctly assigned parent title 'Another Topic'",
			);
		} else {
			console.log("❌ Subtopic C parent title assignment failed");
		}
	} catch (error) {
		console.error("Error in H3 header parsing test:", error);
	}
}

/**
 * Test the complete H3 workflow by creating slides
 */
function testH3CompleteWorkflow() {
	const testMarkdown =
		"## Data Analysis\n\n- Overview of data\n- Key findings\n\n### Data Collection\n- Survey methods\n- Sample size\n- Data quality\n\n### Results\n- Statistical analysis\n- Key insights\n\n```r\nsummary(data)\nplot(x, y)\n```";

	console.log("Testing complete H3 workflow...");

	try {
		// Parse the markdown
		const structure = parseMarkdownToStructure(testMarkdown);
		console.log("Structure:", structure);

		// Create slides (this will test the full pipeline)
		const createdSlides = createSlidesFromStructure(structure);
		console.log("Created slides:", createdSlides);

		// Add content (this will test parent title functionality)
		const success = addContentToSlides(createdSlides);
		console.log(`Content addition success: ${success}`);
	} catch (error) {
		console.error("Error in complete H3 workflow test:", error);
	}
}

/**
 * Test parent title styling
 */
function testParentTitleStyling() {
	console.log("Testing parent title styling...");

	try {
		const presentation = SlidesApp.getActivePresentation();
		const currentSlide = presentation.getSelection().getCurrentPage();

		// Test adding a parent title
		const success = addParentTitleToSlide(currentSlide, "Test Parent Title");
		console.log(`Parent title addition success: ${success}`);

		if (success) {
			console.log("✅ Parent title added successfully");
			console.log(
				"Check the slide for a gray text box at position (24, 18) with 'Test Parent Title'",
			);
		}
	} catch (error) {
		console.error("Error in parent title styling test:", error);
	}
}
