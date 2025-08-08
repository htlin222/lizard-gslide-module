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
 * Creates and returns the sidebar UI with all components loaded.
 *
 * @return {HtmlOutput} The HTML output for the sidebar
 */
function createModularSidebar() {
	const template = createModularHtmlTemplate("sidebar");
	return template.evaluate().setTitle("Lizard Slides").setWidth(300);
}

/**
 * Creates and returns the Markdown to Slides sidebar UI.
 *
 * @return {HtmlOutput} The HTML output for the markdown sidebar
 */
function createMarkdownSidebar() {
	const template = createModularHtmlTemplate("markdown-sidebar");
	return template.evaluate().setTitle("Markdown to Slides").setWidth(320);
}
