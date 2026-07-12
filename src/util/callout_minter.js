// TODO(shared-migration): replace the typeof *_color fallback block with
//   getThemeColors() (shared/theme_colors.js); see kpi_minter.js for the pattern.
//   If/when this minter moves to batch requests, use shared/shape_requests.js builders.
/**
 * Server-side core for the Callout Minter dialog.
 *
 * Replaces the old convertShapeToCallout() one-shot menu action with a
 * minter-style dialog (template picker + header/body inputs + HTML preview).
 *
 * Two modes, auto-detected on insert:
 *  - Convert: a shape/text box is selected → restyle it as the callout body and
 *    add the header banner + left accent bar around it.
 *  - Insert: nothing selected → drop a brand-new callout on the current slide.
 *
 * Built with the SlidesApp service (not the batch API) because both modes need
 * to read the element's geometry and group the pieces afterwards.
 */

/**
 * Single source of truth for callout templates. Colors resolve from the global
 * palette (main_color / accent_color) so they track the configured theme.
 * style: 'banner' = top header + left bar; 'quote' = left bar only, no header.
 *
 * @return {Array<{id,name,headerLabel,style,headerFill,headerText,barColor,bodyFill,bodyBorder,bodyText}>}
 */
function buildCalloutTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const green = "#2E7D32";
	const red = "#C0392B";
	return [
		{
			id: "info",
			name: "INFO",
			headerLabel: "INFO",
			style: "banner",
			headerFill: "#E7EAE7",
			headerText: main,
			barColor: main,
			bodyFill: "#FFFFFF",
			bodyBorder: main,
			bodyText: "#000000",
		},
		{
			id: "note",
			name: "NOTE",
			headerLabel: "NOTE",
			style: "banner",
			headerFill: main,
			headerText: "#FFFFFF",
			barColor: main,
			bodyFill: "#FFFFFF",
			bodyBorder: main,
			bodyText: "#000000",
		},
		{
			id: "tip",
			name: "TIP",
			headerLabel: "TIP",
			style: "banner",
			headerFill: "#E7F4EC",
			headerText: green,
			barColor: green,
			bodyFill: "#FFFFFF",
			bodyBorder: green,
			bodyText: "#000000",
		},
		{
			id: "success",
			name: "SUCCESS",
			headerLabel: "SUCCESS",
			style: "banner",
			headerFill: green,
			headerText: "#FFFFFF",
			barColor: green,
			bodyFill: "#FFFFFF",
			bodyBorder: green,
			bodyText: "#000000",
		},
		{
			id: "warning",
			name: "WARNING",
			headerLabel: "WARNING",
			style: "banner",
			headerFill: "#FCEAD2",
			headerText: "#B26A00",
			barColor: accent,
			bodyFill: "#FFFFFF",
			bodyBorder: accent,
			bodyText: "#000000",
		},
		{
			id: "danger",
			name: "DANGER",
			headerLabel: "DANGER",
			style: "banner",
			headerFill: red,
			headerText: "#FFFFFF",
			barColor: red,
			bodyFill: "#FFFFFF",
			bodyBorder: red,
			bodyText: "#000000",
		},
		{
			id: "quote",
			name: "QUOTE",
			headerLabel: "",
			style: "quote",
			headerFill: "",
			headerText: "",
			barColor: "#999999",
			bodyFill: "#F7F7F7",
			bodyBorder: "#F7F7F7",
			bodyText: "#333333",
		},
	];
}

/**
 * Returns the callout templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getCalloutTemplates() {
	return buildCalloutTemplates_();
}

/**
 * Groq system prompt for turning free-form context into a single callout. The
 * model must reply with EXACTLY two lines — a tiny HEADER label and a concise
 * BODY — and nothing else, so the dialog can map them straight to its fields.
 */
var CALLOUT_AI_SYSTEM_PROMPT = [
	"You convert the user's content into ONE callout (highlight box).",
	"Output EXACTLY two lines and nothing else:",
	"First line:  HEADER: <a very short label, 1-3 words>",
	"Second line: BODY: <one or two concise sentences>",
	"Strict rules:",
	"- No preamble, no explanation, no closing remarks, no code fences.",
	"- Keep the header tiny (e.g. 重點 / 注意 / INFO).",
	"- Keep the body concise; do not invent facts not implied by the input.",
].join("\n");

/**
 * Generates a callout (HEADER + BODY) from free-form context via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateCalloutFromContext(context) {
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

	return callGroq_(CALLOUT_AI_SYSTEM_PROMPT, text, {
		maxTokens: 500,
		temperature: 0.3,
	});
}

/**
 * Inserts (or converts a selection into) a styled callout.
 *
 * @param {{templateId: string, header: string, body: string}} payload
 * @return {{success: boolean, error?: string, mode?: string}}
 */
function insertCalloutIntoSlide(payload) {
	try {
		const p = payload || {};
		const templates = buildCalloutTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}
		const header = (p.header != null ? p.header : tpl.headerLabel) || "";
		const body = p.body || "";
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();

		// Convert mode: a shape or text box is selected. Skipped entirely when a
		// batch caller targets an explicit slide via payload.pageObjectId.
		let mainElement = null;
		let slide = null;
		let mode = "insert";
		if (
			!p.pageObjectId &&
			selection.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT
		) {
			const els = selection.getPageElementRange().getPageElements();
			const el = els && els[0];
			if (el) {
				const t = el.getPageElementType();
				if (t === SlidesApp.PageElementType.SHAPE) {
					mainElement = el.asShape();
				} else if (t === SlidesApp.PageElementType.TEXT_BOX) {
					mainElement = el.asTextBox();
				}
				if (mainElement) {
					slide = el.getParentPage();
					mode = "convert";
				}
			}
		}

		// Insert mode: no usable selection → create a new body text box.
		if (!mainElement) {
			slide = resolveMinterTargetSlide_(presentation, p.pageObjectId);
			if (!slide) return { success: false, error: "No slide available." };
			const W = 320;
			const H = 120;
			const X = Math.max((presentation.getPageWidth() - W) / 2, 0);
			const Y = 140;
			mainElement = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, X, Y, W, H);
			mainElement.getText().setText(body || "Your text here…");
		} else if (body) {
			// Convert mode: only overwrite text when the user typed some.
			mainElement.getText().setText(body);
		}

		// Style the main (body) element.
		mainElement.getFill().setSolidFill(tpl.bodyFill);
		mainElement.getBorder().getLineFill().setSolidFill(tpl.bodyBorder);
		mainElement.getBorder().setWeight(1);
		try {
			const mt = mainElement.getText();
			mt.getTextStyle().setForegroundColor(tpl.bodyText).setFontFamily(font);
			mt.getParagraphStyle().setParagraphAlignment(
				SlidesApp.ParagraphAlignment.START,
			);
		} catch (e) {
			// no text — ignore
		}

		const mx = mainElement.getLeft();
		const my = mainElement.getTop();
		const mw = mainElement.getWidth();
		const mh = mainElement.getHeight();

		const group = [];

		// Header banner (banner style with a non-empty header only).
		const headerH = 20;
		const showHeader = tpl.style === "banner" && header;
		const topOffset = showHeader ? headerH : 0;
		if (showHeader) {
			const headerShape = slide.insertShape(
				SlidesApp.ShapeType.RECTANGLE,
				mx,
				my - headerH,
				mw,
				headerH,
			);
			headerShape.getFill().setSolidFill(tpl.headerFill);
			headerShape.getBorder().getLineFill().setSolidFill(tpl.bodyBorder);
			headerShape.getBorder().setWeight(1);
			headerShape.getText().setText(header);
			headerShape
				.getText()
				.getTextStyle()
				.setForegroundColor(tpl.headerText)
				.setBold(true)
				.setFontSize(12)
				.setFontFamily(font);
			headerShape
				.getText()
				.getParagraphStyle()
				.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
			headerShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
			group.push(headerShape);
		}

		// Left accent bar (spans the header + body).
		const barW = 3;
		const barShape = slide.insertShape(
			SlidesApp.ShapeType.RECTANGLE,
			mx - barW,
			my - topOffset,
			barW,
			mh + topOffset,
		);
		barShape.getFill().setSolidFill(tpl.barColor);
		barShape.getBorder().getLineFill().setSolidFill(tpl.barColor);
		group.push(barShape);

		group.push(mainElement);
		if (group.length > 1) slide.group(group);

		return { success: true, mode: mode };
	} catch (e) {
		console.error("Error inserting callout: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Auto Minter adapter: turns the two-line "HEADER: …" / "BODY: …" AI reply
 * (see CALLOUT_AI_SYSTEM_PROMPT) into the payload insertCalloutIntoSlide()
 * accepts. Parsing mirrors the dialog client
 * (src/components/callout-minter/scripts.html): code fences are stripped,
 * extra lines after a matched HEADER are appended to the body, and when the
 * format isn't followed the first line becomes the header and the rest the
 * body. When the AI gave no header, the template's headerLabel is used
 * (matching the insert fn's own null fallback).
 *
 * @param {string} generatedText - raw LLM output from generateCalloutFromContext()
 * @param {{templateId?: string}} [hints] - optional router hints
 * @return {{templateId: string, header: string, body: string}|null}
 *   null when the parsed body is empty
 */
function autoBuildCalloutPayload_(generatedText, hints) {
	const h = hints || {};

	// Strip code-fence lines, then split into non-empty lines.
	const text = String(generatedText == null ? "" : generatedText)
		.split(/\r?\n/)
		.filter(function (l) {
			return !/^```/.test(l.trim());
		})
		.join("\n")
		.trim();
	const lines = text.split(/\r?\n/).filter(function (l) {
		return l.trim().length;
	});

	let header = "";
	let body = "";
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i];
		const hm = line.match(/^\s*HEADER:\s*(.*)$/i);
		const bm = line.match(/^\s*BODY:\s*(.*)$/i);
		if (hm && !header) {
			header = hm[1].trim();
		} else if (bm && !body) {
			body = bm[1].trim();
		} else if (header && !bm && !hm) {
			body = (body ? body + " " : "") + line.trim();
		}
	}
	// Tolerant fallback: format not followed → first line = header, rest = body.
	if (!header && !body) {
		header = (lines[0] || "").trim();
		body = lines.slice(1).join(" ").trim();
	}
	if (!body) return null;

	const templates = buildCalloutTemplates_();
	let tpl = templates[0];
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === h.templateId) tpl = templates[i];
	}

	return {
		templateId: tpl.id,
		header: header || tpl.headerLabel,
		body: body,
	};
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "callout",
	label: "Callout 重點框",
	emoji: "📌",
	order: 30,
	whenToUse:
		"one single important message/warning/tip/quote to highlight; not for lists of items",
	hintsSpec: '{"templateId":string}',
	generate: "generateCalloutFromContext",
	buildPayload: "autoBuildCalloutPayload_",
	insert: "insertCalloutIntoSlide",
	previewPartial: "src/components/callout-minter/preview",
	previewKind: "callout",
	precheck: "",
	options: [
		{
			name: "templateId",
			label: "範本",
			type: "select",
			choicesFrom: "getCalloutTemplates",
		},
	],
});
