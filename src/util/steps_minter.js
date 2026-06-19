/**
 * Server-side core for the Steps Minter (步驟鑄造器) dialog.
 *
 * Turns a list of steps (one per line: `title | desc`) into a numbered process
 * diagram on the current slide: each step is a numbered circle (1, 2, 3…) plus
 * a text box with a bold title and an optional description, arranged evenly
 * across the page (horizontal) or stacked down it (vertical), with small arrow
 * connectors between consecutive steps. Everything is grouped on insert.
 *
 * Built with the SlidesApp service (like the Callout Minter) because it needs to
 * read page geometry and group the resulting pieces afterwards.
 *
 * Exposes exactly two server functions for the dialog:
 *   - getStepsTemplates()             → templates for client-side preview
 *   - insertStepsIntoSlide(payload)   → renders + groups the steps
 */

/**
 * Single source of truth for steps templates. Colors resolve from the global
 * palette (main_color / accent_color) so they track the configured theme.
 *
 * @return {Array<{id,name,fill,numberText,titleColor,descColor,connector}>}
 */
function buildStepsTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const text =
		(typeof text_color !== "undefined" && text_color) || "#000000";
	// Description text must be a readable mid-gray (matches the dialog preview's
	// #666). Do NOT use sub1_color — that's a near-white background tint
	// (#E7EAE7) and renders almost invisible on a white slide.
	const sub = "#666666";
	return [
		{
			id: "main",
			name: "MAIN",
			fill: main,
			numberText: "#FFFFFF",
			titleColor: text,
			descColor: sub,
			// Connectors are a light gray regardless of theme (read by both the
			// preview and the slide insert).
			connector: "#CCCCCC",
		},
		{
			id: "accent",
			name: "ACCENT",
			fill: accent,
			numberText: "#FFFFFF",
			titleColor: text,
			descColor: sub,
			connector: "#CCCCCC",
		},
	];
}

/**
 * Returns the steps templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getStepsTemplates() {
	return buildStepsTemplates_();
}

/**
 * Groq system prompt for turning free-form context into step lines.
 * Output is ONLY step lines, one per line, in EXACTLY `title | desc` format
 * (desc optional) — no preamble, explanation, or code fences.
 */
const STEPS_AI_SYSTEM_PROMPT = [
	"You convert the user's content into a numbered list of process steps.",
	"Output ONLY step lines, one per line, in EXACTLY this format:",
	"title | desc",
	"The desc part is optional (a line may be just the title).",
	"Example:",
	"收案 | 篩選符合條件的病人",
	"隨機分組 | 1:1 分配至兩組",
	"追蹤 | 每月回診評估",
	"分析 | 主要終點統計",
	"Strict rules:",
	"- No preamble, no explanation, no closing remarks, no code fences.",
	"- Output between 3 and 6 step lines.",
	"- Keep titles short (2–6 characters or words).",
	"- Keep each desc to one short phrase.",
	"- Order the steps logically.",
].join("\n");

/**
 * Generates step lines from free-form context via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateStepsFromContext(context) {
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

	return callGroq_(STEPS_AI_SYSTEM_PROMPT, text, {
		maxTokens: 1200,
		temperature: 0.3,
	});
}

/**
 * Parses the steps textarea (one step per line, `title | desc`) into an array of
 * {title, desc}. Blank lines are ignored; the description is optional.
 *
 * @param {string} text
 * @return {Array<{title: string, desc: string}>}
 */
function parseStepsInput_(text) {
	const raw = (text || "").replace(/\r\n/g, "\n");
	const steps = [];
	raw.split("\n").forEach((line) => {
		const trimmed = line.trim();
		if (!trimmed) return;
		const parts = trimmed.split("|");
		const title = (parts[0] || "").trim();
		const desc = parts.slice(1).join("|").trim();
		if (title || desc) steps.push({ title: title, desc: desc });
	});
	return steps;
}

/**
 * Inserts a numbered steps diagram on the current slide and groups it.
 *
 * @param {{steps: Array<{title:string,desc:string}>, templateId: string,
 *   orientation: string}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertStepsIntoSlide(payload) {
	try {
		const p = payload || {};
		let steps = p.steps || [];
		// Tolerate a raw text payload too.
		if (!steps.length && typeof p.text === "string") {
			steps = parseStepsInput_(p.text);
		}
		if (!steps.length) {
			return { success: false, error: "No steps to insert." };
		}

		const templates = buildStepsTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}
		const vertical = p.orientation === "vertical";
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		const presentation = SlidesApp.getActivePresentation();
		const selection = presentation.getSelection();
		let slide = null;
		try {
			slide = selection.getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		const pageW = presentation.getPageWidth();
		const pageH = presentation.getPageHeight();
		const margin = 40;
		const n = steps.length;
		const circle = 36; // numbered-circle diameter in PT

		const group = [];

		if (!vertical) {
			// Horizontal: circles in a row near the top, title+desc below each.
			const usableW = pageW - 2 * margin;
			const slotW = usableW / n;
			const topY = Math.max((pageH - 150) / 2, margin);
			const textW = Math.min(slotW - 8, 180);
			const textH = 70;

			for (let i = 0; i < n; i++) {
				const slotX = margin + i * slotW;
				const cx = slotX + slotW / 2;
				const circleX = cx - circle / 2;

				_buildStepNumberCircle_(slide, group, tpl, font, i + 1, circleX, topY, circle);

				// Text block centered under the circle.
				const tx = cx - textW / 2;
				const ty = topY + circle + 8;
				_buildStepTextBox_(
					slide,
					group,
					tpl,
					font,
					steps[i],
					tx,
					ty,
					textW,
					textH,
					SlidesApp.ParagraphAlignment.CENTER,
				);

				// Arrow connector to the next circle, vertically centered on circles.
				if (i < n - 1) {
					const nextCx = margin + (i + 1) * slotW + slotW / 2;
					// Short arrow centered in the gap (not spanning circle-to-circle).
					const leftEdge = circleX + circle;
					const rightEdge = nextCx - circle / 2;
					const gapSpan = rightEdge - leftEdge;
					const aw = Math.max(Math.min(44, gapSpan * 0.55), 14);
					const ah = 11;
					const ax = leftEdge + (gapSpan - aw) / 2;
					const ay = topY + (circle - ah) / 2;
					_buildStepConnector_(slide, group, tpl, ax, ay, aw, ah, false);
				}
			}
		} else {
			// Vertical: circles stacked down the left, title+desc to the right.
			const slotH = Math.min((pageH - 2 * margin) / n, 90);
			const startX = margin;
			const startY = margin;
			const textX = startX + circle + 14;
			const textW = Math.min(pageW - margin - textX, 360);
			const textH = slotH - 8;

			for (let i = 0; i < n; i++) {
				const slotY = startY + i * slotH;
				const circleY = slotY + (slotH - circle) / 2;

				_buildStepNumberCircle_(
					slide,
					group,
					tpl,
					font,
					i + 1,
					startX,
					circleY,
					circle,
				);

				_buildStepTextBox_(
					slide,
					group,
					tpl,
					font,
					steps[i],
					textX,
					slotY + 4,
					textW,
					textH,
					SlidesApp.ParagraphAlignment.START,
				);

				// Downward arrow connector between circles.
				if (i < n - 1) {
					const nextCircleY = startY + (i + 1) * slotH + (slotH - circle) / 2;
					// Short arrow centered in the gap (not spanning circle-to-circle).
					const topEdge = circleY + circle;
					const gapSpan = nextCircleY - topEdge;
					const ah = Math.max(Math.min(28, gapSpan * 0.55), 12);
					const aw = 11;
					const ay = topEdge + (gapSpan - ah) / 2;
					const ax = startX + (circle - aw) / 2;
					_buildStepConnector_(slide, group, tpl, ax, ay, aw, ah, true);
				}
			}
		}

		if (group.length > 1) slide.group(group);

		return { success: true };
	} catch (e) {
		console.error("Error inserting steps: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Builds one numbered circle (ELLIPSE with a bold white index) and pushes it
 * into the group array.
 * @private
 */
function _buildStepNumberCircle_(slide, group, tpl, font, index, x, y, size) {
	const c = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, size, size);
	c.getFill().setSolidFill(tpl.fill);
	c.getBorder().getLineFill().setSolidFill(tpl.fill);
	c.getText().setText(String(index));
	c.getText()
		.getTextStyle()
		.setForegroundColor(tpl.numberText)
		.setBold(true)
		.setFontSize(16)
		.setFontFamily(font);
	c.getText()
		.getParagraphStyle()
		.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
	c.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
	group.push(c);
}

/**
 * Builds the title (bold) + desc text box for a step and pushes it into the
 * group array.
 * @private
 */
function _buildStepTextBox_(slide, group, tpl, font, step, x, y, w, h, align) {
	const box = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y, w, h);
	const title = step.title || "";
	const desc = step.desc || "";
	const combined = desc ? title + "\n" + desc : title;
	const t = box.getText();
	t.setText(combined || " ");

	const titleEnd = title.length;
	if (titleEnd > 0) {
		t.getRange(0, titleEnd)
			.getTextStyle()
			.setForegroundColor(tpl.titleColor)
			.setBold(true)
			.setFontSize(14)
			.setFontFamily(font);
		// 1.5x line height on the title paragraph.
		t.getRange(0, titleEnd).getParagraphStyle().setLineSpacing(150);
	}
	if (desc) {
		const descStart = titleEnd + 1;
		const descEnd = descStart + desc.length;
		t.getRange(descStart, descEnd)
			.getTextStyle()
			.setForegroundColor(tpl.descColor)
			.setBold(false)
			.setFontSize(11)
			.setFontFamily(font);
	}
	t.getParagraphStyle().setParagraphAlignment(align);
	box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
	group.push(box);
}

/**
 * Builds a small arrow connector between two steps and pushes it into the group
 * array. `down` selects a downward (vertical) vs. rightward (horizontal) arrow.
 * @private
 */
function _buildStepConnector_(slide, group, tpl, x, y, w, h, down) {
	const type = down
		? SlidesApp.ShapeType.DOWN_ARROW
		: SlidesApp.ShapeType.RIGHT_ARROW;
	const arrow = slide.insertShape(type, x, y, w, h);
	arrow.getFill().setSolidFill(tpl.connector);
	arrow.getBorder().getLineFill().setSolidFill(tpl.connector);
	group.push(arrow);
}
