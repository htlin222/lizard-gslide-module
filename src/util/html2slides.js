/**
 * HTML to Slides Converter Utility
 *
 * Imports gslide-html (module-compatible pure HTML, contract in
 * .claude/skills/gslide-html/reference.md) and creates slides using all 11
 * official predefined layouts.
 *
 * Split of responsibilities:
 * - Dialog front-end parses the HTML with the browser's DOMParser
 *   (src/components/html2slides/parser-client.html) — Apps Script has no
 *   server-side HTML parser.
 * - Server receives structured JSON and builds slides
 *   (src/util/html2slides/slideBuilder.js).
 */

/**
 * Shows the HTML to Slides dialog (paste / upload / URL tabs).
 */
function showHtmlToSlidesDialog() {
	try {
		const html = HtmlService.createTemplateFromFile(
			"src/components/html2slides-dialog",
		)
			.evaluate()
			.setWidth(650)
			.setHeight(560)
			.setTitle("HTML to Slides Converter");
		SlidesApp.getUi().showModalDialog(html, "HTML to Slides");
	} catch (e) {
		console.error("Error showing HTML to Slides dialog: " + e.message);
		SlidesApp.getUi().alert(
			"Could not open the HTML to Slides dialog: " + e.message,
		);
	}
}

/**
 * Fetches raw HTML for the dialog's URL tab. Parsing stays client-side.
 * @param {string} url - http(s) URL
 * @return {string} the response body
 */
function fetchHtmlFromUrl(url) {
	if (!url || !/^https?:\/\//i.test(String(url).trim())) {
		throw new Error("Please provide an http(s):// URL.");
	}
	const response = UrlFetchApp.fetch(String(url).trim(), {
		muteHttpExceptions: true,
		followRedirects: true,
	});
	const code = response.getResponseCode();
	if (code < 200 || code >= 300) {
		throw new Error("Fetch failed with HTTP " + code + ".");
	}
	const text = response.getContentText();
	// keep payloads sane for the round-trip back to the dialog
	const MAX_CHARS = 2000000;
	if (text.length > MAX_CHARS) {
		throw new Error("Fetched HTML is too large (" + text.length + " chars).");
	}
	return text;
}
