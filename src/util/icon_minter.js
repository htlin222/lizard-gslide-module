// TODO(shared-migration): replace the typeof *_color fallback block with
//   getThemeColors() (shared/theme_colors.js); see kpi_minter.js for the pattern.
//   If/when this minter moves to batch requests, use shared/shape_requests.js builders.
/**
 * Server-side core for the Icon Minter dialog.
 *
 * Google Apps Script cannot load an external SVG icon library, so this minter
 * uses emoji / unicode glyphs inserted as a styled text box. The curated glyph
 * set lives client-side in scripts.html (the picker); the server only needs to
 * place the chosen glyph at the chosen size + color.
 *
 * Two placements, auto-detected on insert:
 *  - Near selection: a shape/text box is selected → drop the icon just beside it.
 *  - Center: nothing selected → place the icon at the slide center.
 *
 * Built with the SlidesApp service (not the batch API) because it needs to read
 * the selected element's geometry.
 */

/**
 * Inserts an emoji/glyph icon as a styled text box on the current slide.
 *
 * @param {{glyph: string, size: number, color: string}} payload
 * @return {{success: boolean, error?: string, mode?: string}}
 */
function insertIconIntoSlide(payload) {
	try {
		const p = payload || {};
		const glyph = (p.glyph != null ? String(p.glyph) : "").trim() || "⭐";
		let size = parseFloat(p.size);
		if (!isFinite(size) || size <= 0) size = 48;
		const color =
			(p.color && /^#?[0-9a-fA-F]{6}$/.test(String(p.color))
				? (String(p.color).charAt(0) === "#" ? p.color : "#" + p.color)
				: null) ||
			(typeof main_color !== "undefined" && main_color) ||
			"#3D6869";
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		// Icon box geometry: roughly square, sized generously from the glyph size so
		// the emoji/glyph is never clipped (emoji render taller/wider than their em,
		// and the text box adds internal line spacing). Matches the dialog preview,
		// which shows the glyph un-clipped with generous padding.
		const W = Math.max(size * 2.2, 60);
		const H = Math.max(size * 2.2, 60);

		// Determine target slide + position.
		let slide = null;
		let X = null;
		let Y = null;
		let mode = "center";

		if (
			selection.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT
		) {
			const els = selection.getPageElementRange().getPageElements();
			const el = els && els[0];
			if (el) {
				slide = el.getParentPage();
				// Drop the icon just to the right of the selected element, vertically centered.
				X = el.getLeft() + el.getWidth() + 8;
				Y = el.getTop() + (el.getHeight() - H) / 2;
				mode = "near";
			}
		}

		if (!slide) {
			try {
				slide = selection.getCurrentPage().asSlide();
			} catch (e) {
				slide = presentation.getSlides()[0];
			}
		}
		if (!slide) return { success: false, error: "No slide available." };

		if (X == null || Y == null) {
			X = Math.max((presentation.getPageWidth() - W) / 2, 0);
			Y = Math.max((presentation.getPageHeight() - H) / 2, 0);
		}

		const box = slide.insertTextBox(glyph, X, Y, W, H);
		const text = box.getText();
		text
			.getTextStyle()
			.setFontSize(size)
			.setForegroundColor(color)
			.setFontFamily(font);
		text
			.getParagraphStyle()
			.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
		box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

		return { success: true, mode: mode };
	} catch (e) {
		console.error("Error inserting icon: " + e.message);
		return { success: false, error: e.message };
	}
}
