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

/**
 * Groq system prompt for picking a single representative emoji. Built from an
 * array .join("\n") (mirrors KPI_AI_SYSTEM_PROMPT).
 */
const ICON_AI_SYSTEM_PROMPT = [
	"Reply with exactly ONE emoji character that best represents the content.",
	"No words, no explanation.",
].join("\n");

/**
 * Generates a single representative emoji from free-form context via Groq.
 * Called from the Auto Minter flow through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateIconFromContext(context) {
	const text = (context || "").trim();
	if (!text) {
		return { success: false, error: "No context provided." };
	}

	// Don't throw on a missing key — let the dialog show a friendly prompt to
	// run the explicit "🔑 設定 AI 金鑰 (Groq)" menu item.
	if (!hasUserApiKey()) {
		return {
			success: false,
			needKey: true,
			error:
				"No AI key set. Run 🖖 跨頁功能 → 🔑 設定 AI 金鑰 (Groq) first, then try again.",
		};
	}

	return callGroq_(ICON_AI_SYSTEM_PROMPT, text, {
		maxTokens: 20,
		temperature: 0.4,
	});
}

/**
 * Auto Minter adapter: extracts the FIRST emoji grapheme from the generated
 * text (pragmatic match: pictographic base plus optional VS16/ZWJ sequence;
 * GAS V8 supports unicode property escapes). Falls back to the first
 * non-whitespace character when no emoji is found. `color` is intentionally
 * omitted from the payload — insertIconIntoSlide already defaults it to
 * main_color when absent.
 *
 * @param {string} generatedText - Text from the AI step (ideally one emoji).
 * @param {Object} hints - Unused for icons; kept for the adapter contract.
 * @return {{glyph: string, size: number}|null} insert payload, or null
 */
function autoBuildIconPayload_(generatedText, hints) {
	var h = hints || {};
	const text = String(generatedText == null ? "" : generatedText).trim();
	if (!text) return null;

	let glyph = "";
	try {
		const m = text.match(
			/\p{Extended_Pictographic}(?:️|‍\p{Extended_Pictographic})*/u,
		);
		if (m) glyph = m[0];
	} catch (e) {
		glyph = "";
	}
	if (!glyph) {
		const fallback = text.match(/\S/);
		glyph = fallback ? fallback[0] : "";
	}
	if (!glyph) return null;

	const size =
		typeof h.size === "number" && h.size >= 24 && h.size <= 200 ? h.size : 96;
	return { glyph: glyph, size: size };
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "icon",
	label: "Icon 圖示",
	emoji: "😀",
	order: 120,
	whenToUse:
		"purely decorative single symbol; last resort when no structured layout fits",
	hintsSpec: "",
	generate: "generateIconFromContext",
	buildPayload: "autoBuildIconPayload_",
	insert: "insertIconIntoSlide",
	previewPartial: "src/components/icon-minter/preview",
	previewKind: "glyph",
	precheck: "",
	options: [
		{
			name: "size",
			label: "大小 (pt)",
			type: "number",
			min: 24,
			max: 200,
			default: 96,
		},
	],
});
