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

// Uniform size for every minter dialog (wide + consistent across all minters).
const MINTER_DIALOG_WIDTH_ = 780;
const MINTER_DIALOG_HEIGHT_ = 620;

/**
 * Factory for a minter dialog. Evaluates the modular template at the uniform
 * minter size and inlines `preload` (templates / items) into the HTML as
 * window.__MINTER_PRELOAD__ — so the dialog does NOT need a second
 * google.script.run round-trip on load (faster open).
 *
 * @param {string} path - component index path (without extension)
 * @param {Object} [preload] - data inlined for the dialog to read on load
 * @return {HtmlOutput}
 */
function createMinterDialog_(path, preload) {
	const template = createModularHtmlTemplate(path);
	template.preloadJson = JSON.stringify(preload || {});
	return template
		.evaluate()
		.setWidth(MINTER_DIALOG_WIDTH_)
		.setHeight(MINTER_DIALOG_HEIGHT_);
}

/** @return {HtmlOutput} Table Minter dialog (no on-load fetch). */
function createTableMinterDialog() {
	return createMinterDialog_("src/components/table-minter/index", {});
}

/** @return {HtmlOutput} Grid Minter dialog. */
function createGridMinterDialog() {
	return createMinterDialog_("src/components/grid-minter/index", {
		styles: getStyleDefinitions(),
	});
}

/** @return {HtmlOutput} Callout Minter dialog. */
function createCalloutMinterDialog() {
	return createMinterDialog_("src/components/callout-minter/index", {
		templates: getCalloutTemplates(),
	});
}

/** @return {HtmlOutput} KPI / Big Number minter dialog. */
function createKpiMinterDialog() {
	return createMinterDialog_("src/components/kpi-minter/index", {
		templates: getKpiTemplates(),
	});
}

/** @return {HtmlOutput} Timeline / Roadmap minter dialog. */
function createTimelineMinterDialog() {
	return createMinterDialog_("src/components/timeline-minter/index", {
		templates: getTimelineTemplates(),
	});
}

/** @return {HtmlOutput} Comparison minter dialog. */
function createCompareMinterDialog() {
	return createMinterDialog_("src/components/compare-minter/index", {
		templates: getCompareTemplates(),
	});
}

/** @return {HtmlOutput} Steps minter dialog. */
function createStepsMinterDialog() {
	return createMinterDialog_("src/components/steps-minter/index", {
		templates: getStepsTemplates(),
	});
}

/** @return {HtmlOutput} Image Gallery minter dialog. */
function createGalleryMinterDialog() {
	return createMinterDialog_("src/components/gallery-minter/index", {
		templates: getGalleryTemplates(),
	});
}

/** @return {HtmlOutput} Agenda / TOC minter dialog. */
function createAgendaMinterDialog() {
	return createMinterDialog_("src/components/agenda-minter/index", {
		items: getAgendaItems(),
		templates: getAgendaTemplates(),
	});
}

/** @return {HtmlOutput} Takeaways minter dialog. */
function createTakeawaysMinterDialog() {
	return createMinterDialog_("src/components/takeaways-minter/index", {
		templates: getTakeawaysTemplates(),
	});
}

/** @return {HtmlOutput} Icon minter dialog (no on-load fetch). */
function createIconMinterDialog() {
	return createMinterDialog_("src/components/icon-minter/index", {});
}

/** @return {HtmlOutput} Bar Chart minter dialog. */
function createBarChartMinterDialog() {
	return createMinterDialog_("src/components/barchart-minter/index", {
		templates: getBarChartTemplates(),
	});
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
