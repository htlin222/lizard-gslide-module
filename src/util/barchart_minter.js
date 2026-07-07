// TODO(shared-migration): replace the typeof *_color fallback block with
//   getThemeColors() (shared/theme_colors.js); see kpi_minter.js for the pattern.
//   If/when this minter moves to batch requests, use shared/shape_requests.js builders.
/**
 * Server-side core for the Bar Chart Minter (簡易長條圖鑄造器) dialog.
 *
 * Draws a lightweight, shapes-based bar chart directly onto the current slide —
 * no embedded Google Sheets chart. Each bar is a RECTANGLE scaled
 * proportionally to its value within a fixed chart area, with category labels
 * and (optionally) value labels. All pieces are grouped at the end.
 *
 * Mirrors the Callout Minter pattern (src/util/callout_minter.js): built with
 * the SlidesApp service (not the batch API) so the inserted shapes can be read
 * back and grouped afterwards.
 *
 * Input format (one bar per line): `label | value`
 *   2022 | 12
 *   2023 | 18
 *   2024 | 27
 */

/**
 * Single source of truth for bar-chart color templates. Colors resolve from the
 * global palette (main_color / accent_color / sub1_color …) so they track the
 * configured theme.
 *
 * - main / accent: a single fill color applied to every bar.
 * - multi: cycles a small palette per bar.
 *
 * @return {Array<{id,name,mode,fill?,palette?}>}
 */
function buildBarChartTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const sub1 = (typeof sub1_color !== "undefined" && sub1_color) || "#E7EAE7";
	return [
		{ id: "main", name: "主色 Main", mode: "single", fill: main },
		{ id: "accent", name: "強調 Accent", mode: "single", fill: accent },
		{
			id: "multi",
			name: "多彩 Multi",
			mode: "multi",
			palette: [main, accent, "#2E7D32", "#C0392B", "#2D6CDF", "#8E44AD", sub1],
		},
	];
}

/**
 * Returns the bar-chart templates for client-side preview. Called from the
 * dialog through google.script.run.
 * @return {Array<Object>}
 */
function getBarChartTemplates() {
	return buildBarChartTemplates_();
}

/**
 * Groq system prompt for turning free-form context into bar-chart data lines.
 * Output is ONLY `label | value` lines (value numeric), nothing else.
 */
var BARCHART_AI_SYSTEM_PROMPT = [
	"You convert the user's content into bar-chart data.",
	"Output ONLY data lines, one bar per line, in EXACTLY this format:",
	"label | value",
	"Example:",
	"2022 | 12",
	"2023 | 18",
	"2024 | 27",
	"Strict rules:",
	"- No preamble, no explanation, no closing remarks, no code fences.",
	"- Produce between 3 and 8 bars.",
	"- The value MUST be a number. Strip units, currency symbols, and % signs; keep only the number.",
	"- Keep labels short.",
	"- Do not invent data that isn't implied by the input.",
].join("\n");

/**
 * Generates bar-chart data (`label | value` lines) from free-form context via
 * Groq. Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateBarChartFromContext(context) {
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

	return callGroq_(BARCHART_AI_SYSTEM_PROMPT, text, {
		maxTokens: 1500,
		temperature: 0.3,
	});
}

/**
 * Parses `label | value` lines into {label, value} bars. Non-numeric values are
 * treated as 0 (and the line is kept, so the category still shows). Lines
 * without a separator use the whole line as the label with value 0.
 *
 * @param {Array<{label:string, value:(number|string)}>|string} data
 * @return {Array<{label:string, value:number}>}
 */
function parseBarChartData_(data) {
	const out = [];

	// Accept either an already-structured array (from the dialog) or raw text.
	if (Array.isArray(data)) {
		for (let i = 0; i < data.length; i++) {
			const row = data[i] || {};
			const label = String(row.label != null ? row.label : "").trim();
			let value = Number(row.value);
			if (!isFinite(value)) value = 0;
			if (label || value) out.push({ label: label, value: value });
		}
		return out;
	}

	const text = String(data || "").replace(/\r\n/g, "\n");
	const lines = text.split("\n");
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i].trim();
		if (!line) continue;
		const parts = line.split("|");
		const label = (parts[0] || "").trim();
		let value = 0;
		if (parts.length > 1) {
			const raw = (parts[1] || "").trim().replace(/[, ]+/g, "");
			const n = parseFloat(raw);
			value = isFinite(n) ? n : 0;
		}
		out.push({ label: label, value: value });
	}
	return out;
}

/**
 * Inserts a shapes-based bar chart onto the current slide.
 *
 * Chart geometry: a fixed chart area (width 600, height 260) starting at
 * X ~60, vertically centered. Bars are scaled proportionally to the max value.
 * Vertical bars grow upward from a baseline; horizontal bars grow rightward
 * from a left axis. Category labels sit under (vertical) or beside (horizontal)
 * each bar, and value labels sit at the bar end when enabled. All shapes are
 * grouped.
 *
 * @param {{data: Array<{label:string,value:(number|string)}>, templateId: string,
 *   orientation: string, showValues: boolean}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertBarChartIntoSlide(payload) {
	try {
		const p = payload || {};
		const bars = parseBarChartData_(p.data);
		if (!bars.length) {
			return {
				success: false,
				error: "No data. Use `label | value` per line.",
			};
		}

		const templates = buildBarChartTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}
		const palette =
			tpl.mode === "multi" && tpl.palette && tpl.palette.length
				? tpl.palette
				: [tpl.fill || "#3D6869"];
		const colorFor = (i) => palette[i % palette.length];

		const orientation =
			p.orientation === "horizontal" ? "horizontal" : "vertical";
		const showValues = p.showValues !== false;
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";
		const axisColor = "#999999";

		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		let slide = null;
		try {
			slide = selection.getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		// Fixed chart area, vertically centered on the slide.
		const chartW = 600;
		const chartH = 260;
		const startX = 60;
		const pageH = presentation.getPageHeight();
		const startY = Math.max((pageH - chartH) / 2, 30);

		// Max value drives the scale. Guard against all-zero / negative data.
		let maxVal = 0;
		for (let i = 0; i < bars.length; i++) {
			if (bars[i].value > maxVal) maxVal = bars[i].value;
		}
		if (maxVal <= 0) maxVal = 1;

		const n = bars.length;
		const labelH = 18; // space reserved for category labels
		const valueH = 16; // space reserved for value labels at the bar end
		const group = [];

		if (orientation === "vertical") {
			// Baseline sits above the category labels; bars grow upward from it.
			const baselineY = startY + chartH - labelH;
			const plotH = chartH - labelH - valueH; // usable bar height
			const slot = chartW / n;
			const barW = Math.max(8, slot * 0.6);

			// Baseline (a thin horizontal rectangle).
			const baseline = slide.insertShape(
				SlidesApp.ShapeType.RECTANGLE,
				startX,
				baselineY,
				chartW,
				1.5,
			);
			baseline.getFill().setSolidFill(axisColor);
			baseline.getBorder().getLineFill().setSolidFill(axisColor);
			group.push(baseline);

			for (let i = 0; i < n; i++) {
				const v = Math.max(0, bars[i].value);
				const h = (v / maxVal) * plotH;
				const cx = startX + slot * i + slot / 2;
				const barX = cx - barW / 2;
				const barY = baselineY - h;

				if (h > 0) {
					const bar = slide.insertShape(
						SlidesApp.ShapeType.RECTANGLE,
						barX,
						barY,
						barW,
						h,
					);
					bar.getFill().setSolidFill(colorFor(i));
					bar.getBorder().getLineFill().setSolidFill(colorFor(i));
					group.push(bar);
				}

				// Value label above the bar.
				if (showValues) {
					const valShape = slide.insertShape(
						SlidesApp.ShapeType.TEXT_BOX,
						cx - slot / 2,
						barY - valueH,
						slot,
						valueH,
					);
					valShape.getText().setText(formatBarValue_(bars[i].value));
					styleBarLabel_(valShape, font, 9, "#333333", true);
					group.push(valShape);
				}

				// Category label below the baseline.
				const catShape = slide.insertShape(
					SlidesApp.ShapeType.TEXT_BOX,
					cx - slot / 2,
					baselineY + 2,
					slot,
					labelH,
				);
				catShape.getText().setText(bars[i].label || "");
				styleBarLabel_(catShape, font, 9, "#555555", false);
				group.push(catShape);
			}
		} else {
			// Horizontal bars: vertical axis on the left, bars grow rightward.
			const labelW = 90; // space reserved for category labels on the left
			const axisX = startX + labelW;
			const plotW = chartW - labelW - 40; // leave room for value labels
			const slot = chartH / n;
			const barH = Math.max(8, slot * 0.6);

			// Vertical axis.
			const axis = slide.insertShape(
				SlidesApp.ShapeType.RECTANGLE,
				axisX,
				startY,
				1.5,
				chartH,
			);
			axis.getFill().setSolidFill(axisColor);
			axis.getBorder().getLineFill().setSolidFill(axisColor);
			group.push(axis);

			for (let i = 0; i < n; i++) {
				const v = Math.max(0, bars[i].value);
				const w = (v / maxVal) * plotW;
				const cy = startY + slot * i + slot / 2;
				const barY = cy - barH / 2;

				if (w > 0) {
					const bar = slide.insertShape(
						SlidesApp.ShapeType.RECTANGLE,
						axisX,
						barY,
						w,
						barH,
					);
					bar.getFill().setSolidFill(colorFor(i));
					bar.getBorder().getLineFill().setSolidFill(colorFor(i));
					group.push(bar);
				}

				// Category label to the left of the axis.
				const catShape = slide.insertShape(
					SlidesApp.ShapeType.TEXT_BOX,
					startX,
					cy - slot / 2,
					labelW - 4,
					slot,
				);
				catShape.getText().setText(bars[i].label || "");
				styleBarLabel_(catShape, font, 9, "#555555", false);
				catShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
				try {
					catShape
						.getText()
						.getParagraphStyle()
						.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
				} catch (e) {}
				group.push(catShape);

				// Value label at the bar end.
				if (showValues) {
					const valShape = slide.insertShape(
						SlidesApp.ShapeType.TEXT_BOX,
						axisX + w + 2,
						cy - valueH / 2,
						60,
						valueH,
					);
					valShape.getText().setText(formatBarValue_(bars[i].value));
					styleBarLabel_(valShape, font, 9, "#333333", true);
					valShape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
					try {
						valShape
							.getText()
							.getParagraphStyle()
							.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
					} catch (e) {}
					group.push(valShape);
				}
			}
		}

		if (group.length > 1) slide.group(group);

		return { success: true };
	} catch (e) {
		console.error("Error inserting bar chart: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Applies common label styling (font, size, color, optional bold, center
 * alignment) to a label text box. Wrapped in try/catch so empty text boxes
 * don't throw.
 *
 * @param {GoogleAppsScript.Slides.Shape} shape
 * @param {string} font
 * @param {number} size
 * @param {string} color
 * @param {boolean} bold
 */
function styleBarLabel_(shape, font, size, color, bold) {
	try {
		const t = shape.getText();
		t.getTextStyle()
			.setForegroundColor(color)
			.setFontFamily(font)
			.setFontSize(size)
			.setBold(!!bold);
		t.getParagraphStyle().setParagraphAlignment(
			SlidesApp.ParagraphAlignment.CENTER,
		);
	} catch (e) {
		// empty text box — ignore
	}
	try {
		shape.getBorder().setTransparent();
	} catch (e) {}
}

/**
 * Formats a numeric value for the value label: integers stay whole, decimals
 * are trimmed to at most one place.
 *
 * @param {number} v
 * @return {string}
 */
function formatBarValue_(v) {
	const n = Number(v);
	if (!isFinite(n)) return "0";
	if (Math.round(n) === n) return String(n);
	return String(Math.round(n * 10) / 10);
}

/**
 * Auto Minter adapter: turns AI-generated data lines (`label | value`) into the
 * payload insertBarChartIntoSlide() accepts. Orientation defaults to "vertical"
 * (mirroring the insert fn, where anything but "horizontal" renders vertical).
 *
 * @param {string} generatedText - raw LLM output from generateBarChartFromContext()
 * @param {{templateId?: string, orientation?: string}} [hints] - optional router hints
 * @return {{data: Array<{label:string,value:number}>, templateId: string,
 *   orientation: string, showValues: boolean}|null} null when parsing yields zero bars
 */
function autoBuildBarChartPayload_(generatedText, hints) {
	const h = hints || {};
	const data = parseBarChartData_(generatedText);
	if (!data.length) return null;

	const templates = buildBarChartTemplates_();
	let templateId = templates[0].id;
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === h.templateId) templateId = templates[i].id;
	}

	return {
		data: data,
		templateId: templateId,
		orientation: h.orientation === "horizontal" ? "horizontal" : "vertical",
		showValues: h.showValues !== false,
	};
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "barchart",
	label: "長條圖",
	emoji: "📈",
	order: 50,
	whenToUse:
		"one numeric series across 3-8 categories worth comparing visually as bars",
	hintsSpec: '{"orientation":"vertical|horizontal"}',
	generate: "generateBarChartFromContext",
	buildPayload: "autoBuildBarChartPayload_",
	insert: "insertBarChartIntoSlide",
	previewPartial: "src/components/barchart-minter/preview",
	previewKind: "bars",
	precheck: "",
	options: [
		{
			name: "templateId",
			label: "範本",
			type: "select",
			choicesFrom: "getBarChartTemplates",
		},
		{
			name: "orientation",
			label: "方向",
			type: "select",
			default: "vertical",
			choices: [
				{ value: "vertical", label: "垂直" },
				{ value: "horizontal", label: "水平" },
			],
		},
		{
			name: "showValues",
			label: "顯示數值",
			type: "checkbox",
			default: true,
		},
	],
});
