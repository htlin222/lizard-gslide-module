/**
 * Server-side core for the Timeline / Roadmap Minter dialog (時間軸鑄造器).
 *
 * Turns a list of milestones (one "date | label" per line) into a styled
 * timeline drawn on the current slide:
 *  - A straight bar (line) across (horizontal) or down (vertical) the slide.
 *  - Evenly spaced node circles on the bar.
 *  - For each node, a bold date (theme color) and a label text near it.
 *  - Everything is grouped into one object.
 *
 * Built with the SlidesApp service (not the batch API) because it needs to
 * read page geometry, insert shapes/lines, and group the pieces afterwards.
 *
 * Mirrors the Callout Minter pattern (src/util/callout_minter.js):
 *  - getTimelineTemplates() feeds the dialog's client-side preview.
 *  - insertTimelineIntoSlide(payload) does the actual drawing.
 *  - Colors resolve from the global palette (main_color / accent_color / …).
 */

/**
 * Single source of truth for timeline templates. Colors resolve from the global
 * palette so they track the configured theme. `node` is the node-circle style:
 * 'filled' = solid theme fill; 'outlined' = white fill with a theme border.
 *
 * @return {Array<{id,name,lineColor,nodeFill,nodeBorder,nodeStyle,dateColor,labelColor}>}
 */
function buildTimelineTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const text = (typeof text_color !== "undefined" && text_color) || "#333333";
	return [
		{
			id: "main-filled",
			name: "MAIN · 實心",
			lineColor: main,
			nodeFill: main,
			nodeBorder: main,
			nodeStyle: "filled",
			dateColor: main,
			labelColor: text,
		},
		{
			id: "main-outlined",
			name: "MAIN · 空心",
			lineColor: main,
			nodeFill: "#FFFFFF",
			nodeBorder: main,
			nodeStyle: "outlined",
			dateColor: main,
			labelColor: text,
		},
		{
			id: "accent-filled",
			name: "ACCENT · 實心",
			lineColor: accent,
			nodeFill: accent,
			nodeBorder: accent,
			nodeStyle: "filled",
			dateColor: accent,
			labelColor: text,
		},
		{
			id: "accent-outlined",
			name: "ACCENT · 空心",
			lineColor: accent,
			nodeFill: "#FFFFFF",
			nodeBorder: accent,
			nodeStyle: "outlined",
			dateColor: accent,
			labelColor: text,
		},
	];
}

/**
 * Returns the timeline templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getTimelineTemplates() {
	return buildTimelineTemplates_();
}

/**
 * Resolves a template object from a templateId, falling back to the first.
 * @param {string} templateId
 * @return {Object}
 */
function resolveTimelineTemplate_(templateId) {
	const templates = buildTimelineTemplates_();
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === templateId) return templates[i];
	}
	return templates[0];
}

/**
 * Normalizes incoming milestone items into [{date,label}] with trimmed strings,
 * dropping entries that are entirely empty.
 * @param {Array<{date:string,label:string}>} items
 * @return {Array<{date:string,label:string}>}
 */
function normalizeTimelineItems_(items) {
	const out = [];
	const list = items || [];
	for (let i = 0; i < list.length; i++) {
		const it = list[i] || {};
		const date = (it.date != null ? String(it.date) : "").trim();
		const label = (it.label != null ? String(it.label) : "").trim();
		if (date || label) out.push({ date: date, label: label });
	}
	return out;
}

/**
 * Inserts a styled timeline (horizontal or vertical) onto the current slide.
 *
 * @param {{items: Array<{date:string,label:string}>, templateId: string,
 *   orientation: string}} payload - orientation is 'horizontal' (default) or 'vertical'.
 * @return {{success: boolean, error?: string, count?: number}}
 */
function insertTimelineIntoSlide(payload) {
	try {
		const p = payload || {};
		const items = normalizeTimelineItems_(p.items);
		if (!items.length) {
			return { success: false, error: "No milestones to insert." };
		}

		const tpl = resolveTimelineTemplate_(p.templateId);
		const orientation = p.orientation === "vertical" ? "vertical" : "horizontal";
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();
		const pageW = presentation.getPageWidth();
		const pageH = presentation.getPageHeight();

		// Resolve the target slide (fall back to the first slide).
		let slide = null;
		try {
			slide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		const n = items.length;
		const group = [];

		const nodeSize = 14; // node circle diameter, PT
		const lineWeight = 3;

		if (orientation === "horizontal") {
			const margin = 60;
			const midY = pageH / 2;
			const x0 = margin;
			const x1 = pageW - margin;
			const usableW = Math.max(x1 - x0, 1);

			// The horizontal bar (a thin line) across the slide.
			const line = slide.insertLine(
				SlidesApp.LineCategory.STRAIGHT,
				x0,
				midY,
				x1,
				midY,
			);
			line.getLineFill().setSolidFill(tpl.lineColor);
			line.setWeight(lineWeight);
			group.push(line);

			// Node X positions: single node centered, otherwise evenly spaced.
			for (let i = 0; i < n; i++) {
				const cx = n === 1 ? (x0 + x1) / 2 : x0 + (usableW * i) / (n - 1);
				const cy = midY;

				// Node circle.
				const node = slide.insertShape(
					SlidesApp.ShapeType.ELLIPSE,
					cx - nodeSize / 2,
					cy - nodeSize / 2,
					nodeSize,
					nodeSize,
				);
				node.getFill().setSolidFill(tpl.nodeFill);
				node.getBorder().getLineFill().setSolidFill(tpl.nodeBorder);
				node.getBorder().setWeight(2);
				group.push(node);

				// Alternate date above / label below to reduce overlap.
				const above = i % 2 === 0;
				const textW = Math.max(usableW / Math.max(n, 1), 70);
				const textH = 28;
				const dateY = above ? cy - nodeSize / 2 - textH - 4 : cy + nodeSize / 2 + 4;
				const labelY = above
					? cy - nodeSize / 2 - 2 * textH - 6
					: cy + nodeSize / 2 + textH + 6;

				if (items[i].date) {
					group.push(
						addTimelineText_(
							slide,
							items[i].date,
							cx - textW / 2,
							above ? dateY : dateY,
							textW,
							textH,
							tpl.dateColor,
							true,
							12,
							font,
							SlidesApp.ParagraphAlignment.CENTER,
						),
					);
				}
				if (items[i].label) {
					group.push(
						addTimelineText_(
							slide,
							items[i].label,
							cx - textW / 2,
							above ? labelY : labelY,
							textW,
							textH,
							tpl.labelColor,
							false,
							10,
							font,
							SlidesApp.ParagraphAlignment.CENTER,
						),
					);
				}
			}
		} else {
			// Vertical: bar down the left side, nodes evenly spaced.
			const marginY = 50;
			const lineX = 90;
			const y0 = marginY;
			const y1 = pageH - marginY;
			const usableH = Math.max(y1 - y0, 1);

			const line = slide.insertLine(
				SlidesApp.LineCategory.STRAIGHT,
				lineX,
				y0,
				lineX,
				y1,
			);
			line.getLineFill().setSolidFill(tpl.lineColor);
			line.setWeight(lineWeight);
			group.push(line);

			for (let i = 0; i < n; i++) {
				const cy = n === 1 ? (y0 + y1) / 2 : y0 + (usableH * i) / (n - 1);
				const cx = lineX;

				const node = slide.insertShape(
					SlidesApp.ShapeType.ELLIPSE,
					cx - nodeSize / 2,
					cy - nodeSize / 2,
					nodeSize,
					nodeSize,
				);
				node.getFill().setSolidFill(tpl.nodeFill);
				node.getBorder().getLineFill().setSolidFill(tpl.nodeBorder);
				node.getBorder().setWeight(2);
				group.push(node);

				// Date to the left of the line, label to the right.
				const dateW = lineX - 16;
				const textH = 22;
				if (items[i].date) {
					group.push(
						addTimelineText_(
							slide,
							items[i].date,
							4,
							cy - textH / 2,
							Math.max(dateW, 40),
							textH,
							tpl.dateColor,
							true,
							12,
							font,
							SlidesApp.ParagraphAlignment.END,
						),
					);
				}
				if (items[i].label) {
					const labelX = cx + nodeSize / 2 + 12;
					group.push(
						addTimelineText_(
							slide,
							items[i].label,
							labelX,
							cy - textH / 2,
							Math.max(pageW - labelX - 30, 60),
							textH,
							tpl.labelColor,
							false,
							12,
							font,
							SlidesApp.ParagraphAlignment.START,
						),
					);
				}
			}
		}

		if (group.length > 1) slide.group(group);

		return { success: true, count: n };
	} catch (e) {
		console.error("Error inserting timeline: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Inserts a single text box for a timeline date/label and returns it so the
 * caller can add it to the group.
 *
 * @param {Slide} slide
 * @param {string} text
 * @param {number} x
 * @param {number} y
 * @param {number} w
 * @param {number} h
 * @param {string} color - hex foreground color
 * @param {boolean} bold
 * @param {number} fontSize
 * @param {string} font
 * @param {ParagraphAlignment} alignment
 * @return {Shape}
 */
function addTimelineText_(slide, text, x, y, w, h, color, bold, fontSize, font, alignment) {
	const box = slide.insertShape(
		SlidesApp.ShapeType.TEXT_BOX,
		Math.max(x, 0),
		Math.max(y, 0),
		w,
		h,
	);
	box.getText().setText(text);
	box
		.getText()
		.getTextStyle()
		.setForegroundColor(color)
		.setBold(!!bold)
		.setFontSize(fontSize)
		.setFontFamily(font);
	box.getText().getParagraphStyle().setParagraphAlignment(alignment);
	box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
	return box;
}
