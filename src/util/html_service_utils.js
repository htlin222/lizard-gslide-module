/**
 * Includes an HTML file from the specified path and evaluates any script or style blocks.
 * This is the key function that enables HTML modularization in Google Apps Script.
 *
 * @param {string} filename - The name of the HTML file to include without the .html extension
 * @return {string} The evaluated HTML content
 */
function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Creates and returns an HTML template from the specified file.
 * This is a wrapper around the standard HtmlService.createTemplateFromFile
 * that makes it easier to work with modular HTML files.
 *
 * @param {string} filename - The name of the HTML file to create a template from
 * @return {HtmlTemplate} The HTML template
 */
function createModularHtmlTemplate(filename) {
	const template = HtmlService.createTemplateFromFile(filename);
	template.include = include; // Add the include function to the template
	return template;
}

/**
 * Creates and returns the configuration sidebar UI with all components loaded.
 *
 * @return {HtmlOutput} The HTML output for the configuration sidebar
 */
function createConfigSidebar() {
	const template = createModularHtmlTemplate(
		"src/components/config-sidebar/index",
	);
	return template.evaluate().setTitle("Lizard Slides").setWidth(300);
}

/**
 * Creates and returns the Markdown to Slides sidebar UI.
 *
 * @return {HtmlOutput} The HTML output for the markdown sidebar
 */
function createMarkdownSidebar() {
	const template = createModularHtmlTemplate(
		"src/components/markdown-sidebar/index",
	);
	return template.evaluate().setTitle("Markdown to Slides").setWidth(320);
}

/**
 * Creates and returns the Table Minter dialog UI.
 * Generates Google Slides-ready HTML tables from Markdown for clipboard paste.
 *
 * @return {HtmlOutput} The HTML output for the table minter dialog
 */
function createTableMinterDialog() {
	const template = createModularHtmlTemplate(
		"src/components/table-minter/index",
	);
	return template.evaluate().setWidth(640).setHeight(520);
}

/**
 * Creates and returns the Grid Minter dialog UI.
 * Turns content into a grid of styled "unit" cards inserted onto the slide.
 *
 * @return {HtmlOutput} The HTML output for the grid minter dialog
 */
function createGridMinterDialog() {
	const template = createModularHtmlTemplate(
		"src/components/grid-minter/index",
	);
	return template.evaluate().setWidth(680).setHeight(560);
}

/**
 * Creates and returns the Callout Minter dialog UI.
 * Inserts (or converts a selection into) a styled callout from a template.
 *
 * @return {HtmlOutput} The HTML output for the callout minter dialog
 */
function createCalloutMinterDialog() {
	const template = createModularHtmlTemplate(
		"src/components/callout-minter/index",
	);
	return template.evaluate().setWidth(460).setHeight(560);
}

/**
 * Factory for a minter dialog: builds + evaluates a modular template at the
 * given path and sizes it. Shared by all the small minter dialogs.
 * @param {string} path - component index path (without extension)
 * @param {number} width
 * @param {number} height
 * @return {HtmlOutput}
 */
function createMinterDialog_(path, width, height) {
	return createModularHtmlTemplate(path).evaluate().setWidth(width).setHeight(height);
}

/** @return {HtmlOutput} KPI / Big Number minter dialog */
function createKpiMinterDialog() {
	return createMinterDialog_("src/components/kpi-minter/index", 640, 560);
}

/** @return {HtmlOutput} Timeline / Roadmap minter dialog */
function createTimelineMinterDialog() {
	return createMinterDialog_("src/components/timeline-minter/index", 640, 560);
}

/** @return {HtmlOutput} Comparison minter dialog */
function createCompareMinterDialog() {
	return createMinterDialog_("src/components/compare-minter/index", 680, 560);
}

/** @return {HtmlOutput} Steps minter dialog */
function createStepsMinterDialog() {
	return createMinterDialog_("src/components/steps-minter/index", 640, 560);
}

/** @return {HtmlOutput} Image Gallery minter dialog */
function createGalleryMinterDialog() {
	return createMinterDialog_("src/components/gallery-minter/index", 680, 560);
}

/** @return {HtmlOutput} Agenda / TOC minter dialog */
function createAgendaMinterDialog() {
	return createMinterDialog_("src/components/agenda-minter/index", 600, 560);
}

/** @return {HtmlOutput} Takeaways minter dialog */
function createTakeawaysMinterDialog() {
	return createMinterDialog_("src/components/takeaways-minter/index", 640, 560);
}

/** @return {HtmlOutput} Icon minter dialog */
function createIconMinterDialog() {
	return createMinterDialog_("src/components/icon-minter/index", 460, 560);
}

/** @return {HtmlOutput} Bar Chart minter dialog */
function createBarChartMinterDialog() {
	return createMinterDialog_("src/components/barchart-minter/index", 640, 560);
}

/**
 * Creates and returns the Flowchart sidebar UI with all components loaded.
 *
 * @return {HtmlOutput} The HTML output for the flowchart sidebar
 */
function createFlowchartSidebar() {
	const template = createModularHtmlTemplate("src/components/flowchart/index");
	return template.evaluate().setTitle("Flowchart Tools").setWidth(300);
}
