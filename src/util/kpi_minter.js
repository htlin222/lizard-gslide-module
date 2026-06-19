/**
 * Server-side core for the KPI / Big Number Minter dialog.
 *
 * Turns a few lines of "value | label | trend" into a row of big-number stat
 * cards laid out across the page width (auto-grid by count, equal cells like the
 * Grid Minter) and inserts them onto the current slide via the SlidesApp service.
 *
 * Mirrors the Callout Minter pattern (src/util/callout_minter.js):
 *  - buildKpiTemplates_() is the single source of truth for color themes; colors
 *    resolve from the global palette so they track the configured theme.
 *  - getKpiTemplates() exposes templates to the dialog for client-side preview.
 *  - insertKpiIntoSlide(payload) reads each card's geometry and renders one text
 *    box per card (big value + colored trend arrow, label below).
 *
 * Input line format (also documented in the dialog), one stat per line:
 *   value | label | trend
 * where `trend` is optional and one of up/down/flat (also ↑/↓/→). Examples:
 *   87% | 疾病控制率 | up
 *   HR 0.62 | 主要終點 | down
 *   n = 1,204 | 收案人數
 */

/**
 * Single source of truth for KPI card color themes. Value/label colors resolve
 * from the global palette (main_color / accent_color / text_color …) so they
 * track the configured theme. The trend arrow is colored independently of the
 * theme (up=green, down=red, flat=gray) by trendColor_().
 *
 * @return {Array<{id,name,valueColor,labelColor,cardFill,cardBorder}>}
 */
function buildKpiTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const text = (typeof text_color !== "undefined" && text_color) || "#333333";
	const sub1 = (typeof sub1_color !== "undefined" && sub1_color) || "#E7EAE7";
	return [
		{
			id: "main",
			name: "MAIN",
			valueColor: main,
			labelColor: text,
			cardFill: "#FFFFFF",
			cardBorder: main,
		},
		{
			id: "accent",
			name: "ACCENT",
			valueColor: accent,
			labelColor: text,
			cardFill: "#FFFFFF",
			cardBorder: accent,
		},
		{
			id: "neutral",
			name: "NEUTRAL",
			valueColor: text,
			labelColor: "#666666",
			cardFill: sub1,
			cardBorder: sub1,
		},
	];
}

/**
 * Resolves the color for a trend arrow. Independent of the theme.
 * @param {string} trend - 'up' | 'down' | 'flat' (or arrow chars)
 * @return {string} hex color
 */
function trendColor_(trend) {
	const t = normalizeTrend_(trend);
	if (t === "up") return "#2E7D32"; // green
	if (t === "down") return "#C0392B"; // red
	if (t === "flat") return "#7F8C8D"; // gray
	return "";
}

/**
 * Maps the trend keyword/symbol to a normalized token.
 * @param {string} trend
 * @return {string} 'up' | 'down' | 'flat' | ''
 */
function normalizeTrend_(trend) {
	const t = String(trend == null ? "" : trend)
		.trim()
		.toLowerCase();
	if (!t) return "";
	if (t === "up" || t === "↑" || t === "u") return "up";
	if (t === "down" || t === "↓" || t === "d") return "down";
	if (t === "flat" || t === "→" || t === "f" || t === "-") return "flat";
	return "";
}

/**
 * Returns the arrow glyph for a trend token, or '' for none.
 * @param {string} trend
 * @return {string}
 */
function trendArrow_(trend) {
	const t = normalizeTrend_(trend);
	if (t === "up") return "↑";
	if (t === "down") return "↓";
	if (t === "flat") return "→";
	return "";
}

/**
 * Returns the KPI templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getKpiTemplates() {
	return buildKpiTemplates_();
}

/**
 * Parses KPI lines into an array of {value, label, trend}. Each non-empty line
 * is `value | label | trend` (trend optional). Tolerant of full-width "｜".
 *
 * @param {string} text
 * @return {Array<{value: string, label: string, trend: string}>}
 */
function parseKpiLines_(text) {
	const raw = String(text == null ? "" : text).replace(/\r\n/g, "\n");
	const items = [];
	const lines = raw.split("\n");
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i].trim();
		if (!line) continue;
		const parts = line.split(/\s*[|｜]\s*/);
		const value = (parts[0] || "").trim();
		const label = (parts[1] || "").trim();
		const trend = normalizeTrend_(parts[2] || "");
		if (value || label) {
			items.push({ value: value, label: label, trend: trend });
		}
	}
	return items;
}

/**
 * Inserts a row of KPI / big-number stat cards onto the current slide. Cards use
 * grid-style equal positioning across the page width (start Y ~140, margin 30,
 * gap 15). Each card is a text box: big value (with colored trend arrow) above a
 * smaller label.
 *
 * @param {{items: Array<{value:string,label:string,trend:string}>, templateId: string}} payload
 * @return {{success: boolean, error?: string, count?: number}}
 */
function insertKpiIntoSlide(payload) {
	try {
		const p = payload || {};

		// Accept either pre-parsed items or a raw `text` blob.
		let items = p.items || [];
		if ((!items || !items.length) && p.text) {
			items = parseKpiLines_(p.text);
		}
		if (!items || !items.length) {
			return { success: false, error: "No KPI lines to insert." };
		}

		const templates = buildKpiTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}

		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();
		const pageW = presentation.getPageWidth();

		// Resolve the target slide (fall back to the first slide).
		let slide = null;
		try {
			slide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		// Grid-style equal positioning across the page width.
		const margin = 30;
		const gap = 15;
		const top = 140;
		const cardH = 110;
		const n = items.length;
		const usableW = pageW - 2 * margin;
		const cardW = (usableW - (n - 1) * gap) / n;

		for (let i = 0; i < n; i++) {
			const item = items[i];
			const x = margin + i * (cardW + gap);
			renderKpiCard_(slide, x, top, cardW, cardH, item, tpl, font);
		}

		return { success: true, count: n };
	} catch (e) {
		console.error("Error inserting KPI: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Renders a single KPI card as a text box (value + trend arrow, then label).
 *
 * @param {Slide} slide
 * @param {number} x
 * @param {number} y
 * @param {number} w
 * @param {number} h
 * @param {{value:string,label:string,trend:string}} item
 * @param {{valueColor:string,labelColor:string,cardFill:string,cardBorder:string}} tpl
 * @param {string} font
 */
function renderKpiCard_(slide, x, y, w, h, item, tpl, font) {
	const value = item.value || "";
	const label = item.label || "";
	const arrow = trendArrow_(item.trend);
	const arrowColor = trendColor_(item.trend);

	const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, w, h);

	// Subtle card background + border per template.
	box.getFill().setSolidFill(tpl.cardFill);
	box.getBorder().getLineFill().setSolidFill(tpl.cardBorder);
	box.getBorder().setWeight(1);
	box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

	// Build "value ↑" + newline + label. Track ranges so we can color/size each.
	const valueLine = arrow ? value + " " + arrow : value;
	const combined = label ? valueLine + "\n" + label : valueLine;

	const textRange = box.getText();
	textRange.setText(combined);

	// Base styling for the whole box.
	textRange.getParagraphStyle().setParagraphAlignment(
		SlidesApp.ParagraphAlignment.CENTER,
	);

	// Value line: bold, big, theme valueColor.
	const valueRange = textRange.getRange(0, valueLine.length);
	if (valueRange) {
		valueRange
			.getTextStyle()
			.setBold(true)
			.setFontSize(38)
			.setForegroundColor(tpl.valueColor)
			.setFontFamily(font);
	}

	// Trend arrow: colored independently (last `arrow.length` chars of the value
	// line — the arrow plus the leading space). Slightly smaller than the value,
	// matching the preview proportion (preview value 24px / arrow 20px).
	if (arrow && arrowColor) {
		const arrowStart = value.length + 1; // skip "value "
		const arrowRange = textRange.getRange(arrowStart, valueLine.length);
		if (arrowRange) {
			arrowRange.getTextStyle().setForegroundColor(arrowColor).setFontSize(32);
		}
	}

	// Label line: smaller, labelColor.
	if (label) {
		const labelStart = valueLine.length + 1; // skip the "\n"
		const labelEnd = labelStart + label.length;
		const labelRange = textRange.getRange(labelStart, labelEnd);
		if (labelRange) {
			labelRange
				.getTextStyle()
				.setBold(false)
				.setFontSize(17)
				.setForegroundColor(tpl.labelColor)
				.setFontFamily(font);
		}
	}
}
