// TODO(shared-migration): replace the typeof *_color fallback block with
//   getThemeColors() (shared/theme_colors.js); see kpi_minter.js for the pattern.
//   If/when this minter moves to batch requests, use shared/shape_requests.js builders.
/**
 * Server-side core for the Agenda / TOC Minter (目錄鑄造器) dialog.
 *
 * Builds an "Agenda" table-of-contents block on the current slide. The dialog
 * pre-fills its textarea with section titles auto-detected from the deck — slides
 * using a SECTION_HEADER layout (falling back to the first non-empty text of
 * each slide) — which the user can then edit before inserting.
 *
 * Mirrors the Callout / Grid Minter pattern:
 *  - Templates resolve their colors from the global palette (main_color, etc.)
 *    so they track the configured theme.
 *  - Insertion uses the SlidesApp service so the inserted text boxes can be
 *    grouped together afterwards.
 *
 * Exposes exactly three server functions to the dialog:
 *  - getAgendaItems()
 *  - getAgendaTemplates()
 *  - insertAgendaIntoSlide(payload)
 */

/**
 * Auto-detects section titles from the active deck for pre-filling the dialog.
 *
 * Detection strategy:
 *  1. Prefer slides whose layout is "SECTION_HEADER"; take each one's title
 *     placeholder text (or, failing that, the first non-empty text shape).
 *  2. If no section-header slides are found, fall back to the first non-empty
 *     title/text of every slide so the user still gets a useful starting list.
 *
 * Returns [] safely on any error so the dialog never breaks.
 *
 * @return {Array<string>} ordered, de-duplicated section titles
 */
function getAgendaItems() {
	// Cache briefly so repeat dialog opens don't re-scan every slide. The short
	// TTL keeps it fresh enough as the deck changes.
	try {
		const cache = CacheService.getDocumentCache();
		if (cache) {
			const hit = cache.get("agenda_items_v1");
			if (hit) return JSON.parse(hit);
		}
		const items = computeAgendaItems_();
		if (cache) cache.put("agenda_items_v1", JSON.stringify(items), 30);
		return items;
	} catch (e) {
		return computeAgendaItems_();
	}
}

/**
 * Computes agenda items by scanning the deck (uncached). See getAgendaItems().
 * @return {Array<string>} ordered, de-duplicated section titles
 */
function computeAgendaItems_() {
	try {
		const presentation = SlidesApp.getActivePresentation();
		const slides = presentation.getSlides();
		const items = [];

		// Pass 1: SECTION_HEADER layout slides.
		for (let i = 0; i < slides.length; i++) {
			const slide = slides[i];
			let layoutName = "";
			try {
				const layout = slide.getLayout();
				layoutName = layout ? layout.getLayoutName() : "";
			} catch (e) {
				layoutName = "";
			}
			if (layoutName === "SECTION_HEADER") {
				const title = agendaSlideTitle_(slide);
				if (title) items.push(title);
			}
		}

		if (items.length) return dedupeAgendaItems_(items);

		// Pass 2 (fallback): first non-empty title/text of each slide, skipping
		// the very first slide (usually the cover/title slide).
		for (let i = 1; i < slides.length; i++) {
			const title = agendaSlideTitle_(slides[i]);
			if (title) items.push(title);
		}

		return dedupeAgendaItems_(items);
	} catch (e) {
		console.error("Error detecting agenda items: " + e.message);
		return [];
	}
}

/**
 * Returns the first usable title text for a slide: the TITLE placeholder if it
 * has text, otherwise the first non-empty text shape on the slide.
 *
 * @param {Slide} slide
 * @return {string} trimmed title text, or "" if none
 */
function agendaSlideTitle_(slide) {
	try {
		const placeholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
		if (placeholder && placeholder.asShape) {
			const t = placeholder.asShape().getText().asString().trim();
			if (t) return t;
		}
	} catch (e) {
		// no title placeholder — fall through to the text-shape scan
	}
	try {
		const shapes = slide.getShapes();
		for (const shape of shapes) {
			const txt = shape.getText().asString().trim();
			if (txt) return txt;
		}
	} catch (e) {
		// ignore
	}
	return "";
}

/**
 * De-duplicates agenda items (case-insensitive) while preserving order and
 * dropping empty entries.
 *
 * @param {Array<string>} items
 * @return {Array<string>}
 */
function dedupeAgendaItems_(items) {
	const seen = {};
	const out = [];
	for (const raw of items || []) {
		const item = (raw || "").trim();
		if (!item) continue;
		const key = item.toLowerCase();
		if (seen[key]) continue;
		seen[key] = true;
		out.push(item);
	}
	return out;
}

/**
 * Single source of truth for agenda templates. Colors resolve from the global
 * palette so they track the configured theme.
 *
 * layout:
 *  - 'numbered'   = single-column numbered list
 *  - 'bulleted'   = single-column bulleted list
 *  - 'twoColumn'  = numbered list split across two columns
 *
 * @return {Array<{id,name,layout,heading,headingColor,itemColor,markerColor}>}
 */
function buildAgendaTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const text = (typeof text_color !== "undefined" && text_color) || "#000000";
	// Readable mid-gray for subdued item text. NOT sub1_color — that's a
	// near-white background tint (#E7EAE7) and would be almost invisible.
	const sub1 = "#666666";
	return [
		{
			id: "numbered",
			name: "Numbered List",
			layout: "numbered",
			heading: "Agenda",
			headingColor: main,
			itemColor: text,
			markerColor: accent,
		},
		{
			id: "bulleted",
			name: "Bulleted List",
			layout: "bulleted",
			heading: "目錄",
			headingColor: main,
			itemColor: text,
			markerColor: main,
		},
		{
			id: "twoColumn",
			name: "Two Column",
			layout: "twoColumn",
			heading: "Agenda",
			headingColor: main,
			itemColor: sub1,
			markerColor: accent,
		},
	];
}

/**
 * Returns the agenda templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getAgendaTemplates() {
	return buildAgendaTemplates_();
}

/**
 * Resolves a template by id, defaulting to the first template.
 * @param {string} id
 * @return {Object}
 */
function resolveAgendaTemplate_(id) {
	const templates = buildAgendaTemplates_();
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === id) return templates[i];
	}
	return templates[0];
}

/**
 * Parses the textarea payload (one item per line) into a clean array.
 * @param {Array<string>|string} items
 * @return {Array<string>}
 */
function parseAgendaItems_(items) {
	let list = [];
	if (Array.isArray(items)) {
		list = items;
	} else if (typeof items === "string") {
		list = items.split(/\r?\n/);
	}
	const out = [];
	for (const raw of list) {
		const item = (raw || "").trim();
		if (item) out.push(item);
	}
	return out;
}

/**
 * Inserts an agenda / TOC block onto the current slide: a heading text box plus
 * the items rendered as a numbered or bulleted list (or split into two columns
 * for the two-column template / when there are many items). Styled with theme
 * colors, left-aligned, positioned near the top-left. All inserted boxes are
 * grouped together.
 *
 * @param {{items: Array<string>|string, templateId: string}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertAgendaIntoSlide(payload) {
	try {
		const p = payload || {};
		const items = parseAgendaItems_(p.items);
		if (!items.length) {
			return { success: false, error: "No agenda items to insert." };
		}

		const tpl = resolveAgendaTemplate_(p.templateId);
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();

		// Resolve the target slide (fall back to the first slide).
		let slide = null;
		try {
			slide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		const pageW = presentation.getPageWidth();

		// Layout constants.
		const LEFT = 50;
		const TOP = 110;
		const HEADING_H = 40;
		const GAP = 8; // gap between heading and list
		const itemH = 24; // per-item height estimate
		const usableW = Math.max(pageW - LEFT - 50, 200);

		const group = [];

		// Heading.
		const heading = slide.insertShape(
			SlidesApp.ShapeType.TEXT_BOX,
			LEFT,
			TOP,
			usableW,
			HEADING_H,
		);
		heading.getText().setText(tpl.heading || "Agenda");
		heading
			.getText()
			.getTextStyle()
			.setForegroundColor(tpl.headingColor)
			.setBold(true)
			.setFontSize(28)
			.setFontFamily(font);
		heading
			.getText()
			.getParagraphStyle()
			.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
		group.push(heading);

		const listTop = TOP + HEADING_H + GAP;
		const useTwoCol = tpl.layout === "twoColumn" || items.length > 8;

		if (useTwoCol) {
			// Split items as evenly as possible into two columns.
			const half = Math.ceil(items.length / 2);
			const colGap = 20;
			const colW = (usableW - colGap) / 2;
			const col1 = items.slice(0, half);
			const col2 = items.slice(half);

			const box1 = buildAgendaListBox_(
				slide,
				LEFT,
				listTop,
				colW,
				col1.length * itemH,
				col1,
				tpl,
				font,
				0,
			);
			group.push(box1);

			if (col2.length) {
				const box2 = buildAgendaListBox_(
					slide,
					LEFT + colW + colGap,
					listTop,
					colW,
					col2.length * itemH,
					col2,
					tpl,
					font,
					col1.length,
				);
				group.push(box2);
			}
		} else {
			const box = buildAgendaListBox_(
				slide,
				LEFT,
				listTop,
				usableW,
				items.length * itemH,
				items,
				tpl,
				font,
				0,
			);
			group.push(box);
		}

		if (group.length > 1) slide.group(group);

		return { success: true };
	} catch (e) {
		console.error("Error inserting agenda: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Creates and styles one list text box for the agenda. Numbered templates prefix
 * each item with "N. " (continuing from startNumber); bulleted templates use the
 * Slides bullet preset.
 *
 * @param {Slide} slide
 * @param {number} x
 * @param {number} y
 * @param {number} w
 * @param {number} h
 * @param {Array<string>} items
 * @param {Object} tpl - resolved template
 * @param {string} font
 * @param {number} startNumber - number of preceding items (for numbered offset)
 * @return {Shape} the created text box
 */
function buildAgendaListBox_(slide, x, y, w, h, items, tpl, font, startNumber) {
	const box = slide.insertShape(
		SlidesApp.ShapeType.TEXT_BOX,
		x,
		y,
		w,
		Math.max(h, 30),
	);

	const isNumbered = tpl.layout !== "bulleted";
	const lines = [];
	for (let i = 0; i < items.length; i++) {
		if (isNumbered) {
			lines.push(startNumber + i + 1 + ". " + items[i]);
		} else {
			lines.push(items[i]);
		}
	}
	const text = box.getText();
	text.setText(lines.join("\n"));

	text
		.getTextStyle()
		.setForegroundColor(tpl.itemColor)
		.setFontSize(16)
		.setFontFamily(font);
	text
		.getParagraphStyle()
		.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

	// Bulleted layout: apply a bullet preset and color the markers via the item
	// color (Slides has no separate marker color, so the whole line shares it).
	if (!isNumbered) {
		try {
			text.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
		} catch (e) {
			// ignore — fall back to plain lines
		}
	}

	// Numbered layout: color the "N. " marker at the start of each line with the
	// template's marker color so the numbers stand out.
	if (isNumbered) {
		let charIndex = 0;
		for (let i = 0; i < lines.length; i++) {
			const markerLen = (startNumber + i + 1 + ". ").length;
			try {
				text
					.getRange(charIndex, charIndex + markerLen)
					.getTextStyle()
					.setForegroundColor(tpl.markerColor)
					.setBold(true);
			} catch (e) {
				// ignore range errors
			}
			charIndex += lines[i].length + 1; // +1 for the newline
		}
	}

	return box;
}

/**
 * Groq system prompt for turning free-form context into agenda items. Built
 * from an array .join("\n") (mirrors KPI_AI_SYSTEM_PROMPT). The model must
 * output ONLY the items, one per line, exactly as parseAgendaItems_() reads.
 */
const AGENDA_AI_SYSTEM_PROMPT = [
	"You extract 3-8 agenda / section items from the user's content.",
	"Output ONLY the items, one item per line.",
	"Strict rules:",
	"- No numbering, no bullets, no preamble, no explanation, no code fences.",
	"- Output between 3 and 8 lines.",
	"- Keep each item under ~40 characters.",
].join("\n");

/**
 * Generates agenda items from free-form context via Groq. Called from the
 * Auto Minter flow (and usable from dialogs) through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateAgendaFromContext(context) {
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

	return callGroq_(AGENDA_AI_SYSTEM_PROMPT, text, {
		maxTokens: 300,
		temperature: 0.3,
	});
}

/**
 * Auto Minter adapter: turns generated one-item-per-line text into an
 * insertAgendaIntoSlide payload via parseAgendaItems_().
 *
 * @param {string} generatedText - Agenda item lines from the AI step.
 * @param {{templateId?: string}} hints
 * @return {Object|null} insert payload, or null when nothing parseable
 */
function autoBuildAgendaPayload_(generatedText, hints) {
	var h = hints || {};
	const items = parseAgendaItems_(generatedText);
	if (!items.length) return null;

	const templates = buildAgendaTemplates_();
	let templateId = templates[0].id;
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === h.templateId) templateId = h.templateId;
	}

	return { items: items, templateId: templateId };
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "agenda",
	label: "議程/目錄",
	emoji: "📋",
	order: 90,
	whenToUse:
		"an agenda, outline or table of contents: a flat list of 3-8 section titles",
	hintsSpec: "",
	generate: "generateAgendaFromContext",
	buildPayload: "autoBuildAgendaPayload_",
	insert: "insertAgendaIntoSlide",
	previewPartial: "src/components/agenda-minter/preview",
	previewKind: "list",
	precheck: "",
	options: [
		{
			name: "templateId",
			label: "範本",
			type: "select",
			choicesFrom: "getAgendaTemplates",
		},
	],
});
