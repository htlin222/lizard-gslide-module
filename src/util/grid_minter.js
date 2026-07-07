/**
 * Server-side core for the Grid Minter dialog.
 *
 * Turns arbitrary user "context" into a list of "units" — each unit being a
 * heading + subheading + paragraph — then lays those units out as styled cards
 * on a grid (e.g. 3×2) and inserts them onto the current slide via the Slides
 * API. Overflow units spill onto freshly appended slides.
 *
 * Mirrors the Table Minter pattern (src/util/table_minter.js):
 *  - AI generation reuses callGroq_() / hasUserApiKey() from src/util/aiKey.js.
 *  - Insertion batches every request into a single Slides.Presentations.batchUpdate.
 *  - Card colors come from getStyleDefinitions() (src/util/default_style.js).
 *
 * The unit markdown format (also documented in the dialog):
 *   # Title
 *   ## Subtitle
 *
 *   A paragraph of body text.
 *   ---
 *   # Next Title
 *   ...
 * Units are separated by a line containing only "---".
 */

/**
 * Builds the Groq system prompt for grid units. When `hasSubtitle` is true each
 * unit carries a '## ' subtitle line; otherwise the prompt explicitly tells the
 * model to omit subtitles (title + paragraph only).
 *
 * @param {boolean} hasSubtitle
 * @return {string}
 */
function buildGridSystemPrompt_(hasSubtitle) {
	const lines = [
		"You convert the user's content into a set of grid 'units' (key points).",
		"Output EXACTLY this and nothing else:",
	];
	if (hasSubtitle) {
		lines.push(
			"- For each key point, one unit made of, in this order:",
			"  1. A title line starting with '# ' (a short 1–4 word heading).",
			"  2. A subtitle line starting with '## ' (a short supporting phrase).",
			"  3. After a blank line, ONE concise paragraph (1–2 sentences).",
		);
	} else {
		lines.push(
			"- For each key point, one unit made of, in this order:",
			"  1. A title line starting with '# ' (a short 1–4 word heading).",
			"  2. After a blank line, ONE concise paragraph (1–2 sentences).",
			"- Do NOT include any '## ' subtitle lines.",
		);
	}
	lines.push(
		"- Separate consecutive units with a line that contains only '---'.",
		"Strict rules:",
		"- No explanation, no preamble, no closing remarks, no code fences.",
		"- The very first line MUST start with '# '.",
		"- Aim for between 2 and 9 units, grouping the content into logical key points.",
		"- Keep text concise; do not invent facts that aren't implied by the input.",
	);
	return lines.join("\n");
}

/**
 * Generates grid-unit Markdown from free-form context via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @param {boolean} [hasSubtitle=true] - Whether units should include '## ' subtitles.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateGridMarkdownFromContext(context, hasSubtitle) {
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

	const wantSubtitle = hasSubtitle !== false; // default to including subtitles
	return callGroq_(buildGridSystemPrompt_(wantSubtitle), text, {
		maxTokens: 2500,
		temperature: 0.3,
	});
}

/**
 * Parses unit markdown into an array of {title, subtitle, body}.
 * Units are separated by a line containing only "---". Within a unit, the first
 * "# " line is the title, the first "## " line is the subtitle, and any
 * remaining non-heading text is the body paragraph.
 *
 * @param {string} markdown
 * @return {Array<{title: string, subtitle: string, body: string}>}
 */
function parseGridUnits_(markdown) {
	const text = (markdown || "").replace(/\r\n/g, "\n").trim();
	if (!text) return [];

	// Split on a line that is only dashes (---, ----, …), tolerating whitespace.
	const blocks = text
		.split(/\n\s*-{3,}\s*\n/)
		.map((b) => b.trim())
		.filter((b) => b.length);

	const units = [];
	for (const block of blocks) {
		const lines = block.split("\n");
		let title = "";
		let subtitle = "";
		const bodyLines = [];
		for (const raw of lines) {
			const line = raw.trim();
			if (!line) continue;
			const h2 = line.match(/^##\s+(.+)$/);
			const h1 = line.match(/^#\s+(.+)$/);
			if (h2) {
				if (!subtitle) subtitle = h2[1].trim();
			} else if (h1) {
				if (!title) title = h1[1].trim();
			} else {
				bodyLines.push(line);
			}
		}
		const body = bodyLines.join(" ").trim();
		if (title || subtitle || body) {
			units.push({ title: title, subtitle: subtitle, body: body });
		}
	}
	return units;
}

/**
 * Suggests a grid layout (rows × cols) for a given number of units. Keeps grids
 * roughly landscape (cols >= rows) to match 16:9 slides. Mirrors the client-side
 * suggestLayout() in the dialog scripts.
 *
 * @param {number} n - unit count
 * @return {{rows: number, cols: number}}
 */
function suggestGridLayout_(n) {
	const table = {
		1: { rows: 1, cols: 1 },
		2: { rows: 1, cols: 2 },
		3: { rows: 1, cols: 3 },
		4: { rows: 2, cols: 2 },
		5: { rows: 2, cols: 3 },
		6: { rows: 2, cols: 3 },
		7: { rows: 3, cols: 3 },
		8: { rows: 3, cols: 3 },
		9: { rows: 3, cols: 3 },
	};
	if (n <= 0) return { rows: 1, cols: 1 };
	if (table[n]) return table[n];
	// >9: keep 3 columns, grow rows.
	return { rows: Math.ceil(n / 3), cols: 3 };
}

/**
 * Computes card rectangles (in PT) for a rows×cols grid. Columns divide the
 * usable width evenly; rows are sized to their content via `opts.rowHeight`
 * (the tallest unit) rather than stretching to fill the slide. The grid starts
 * at `opts.top` (defaulting to 100, like the Table Minter) and `opts.left`.
 *
 * @param {number} rows
 * @param {number} cols
 * @param {number} pageW - page width in PT
 * @param {number} pageH - page height in PT
 * @param {{margin?: number, gap?: number, top?: number, left?: number, rowHeight?: number}} [opts]
 * @return {Array<{x: number, y: number, w: number, h: number}>} row-major order
 */
function computeGridPositions_(rows, cols, pageW, pageH, opts) {
	const o = opts || {};
	const margin = o.margin != null ? o.margin : 30;
	const gap = o.gap != null ? o.gap : 15;
	const left = o.left != null ? o.left : margin;
	const top = o.top != null ? o.top : 100; // start like the Table Minter

	const usableW = pageW - left - margin;
	const cellW = (usableW - (cols - 1) * gap) / cols;

	// Cards are as tall as the tallest unit's content (opts.rowHeight), NOT
	// stretched to fill the slide. Clamp so the grid still fits between `top`
	// and the bottom margin; keep a small floor so sparse cards stay usable.
	const maxRowH = (pageH - top - margin - (rows - 1) * gap) / rows;
	let cellH = o.rowHeight > 0 ? o.rowHeight : maxRowH;
	cellH = Math.max(70, Math.min(cellH, maxRowH));

	const positions = [];
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			positions.push({
				x: left + c * (cellW + gap),
				y: top + r * (cellH + gap),
				w: cellW,
				h: cellH,
			});
		}
	}
	return positions;
}

/**
 * Base font sizes (title/subtitle/body) for a card. Fixed defaults — title 18,
 * subtitle 12, body 12 — that stay put unless a card has to shrink to fit, so
 * cards size to their text rather than text growing to fill an oversized card.
 *
 * @return {{title: number, subtitle: number, body: number}}
 */
function baseFontSizes_() {
	return { title: 18, subtitle: 12, body: 12 };
}

/**
 * Estimates the natural height (PT) a unit's content needs in a card of the
 * given width, at the base font sizes plus padding. Used to pick a uniform row
 * height equal to the tallest unit.
 *
 * @param {{title: string, subtitle: string, body: string}} unit
 * @param {number} cardW - card width in PT
 * @return {number}
 */
function naturalUnitHeight_(unit, cardW) {
	const pad = 8;
	const f = baseFontSizes_();
	const innerW = Math.max(20, cardW - 2 * pad);
	// Reserve the subtitle row even when blank so cards with/without a subtitle
	// take the same height.
	const sub = unit.subtitle || " ";
	return (
		estimateBlockHeight_(
			unit.title || "",
			sub,
			unit.body || "",
			f.title,
			f.subtitle,
			f.body,
			innerW,
		) +
		2 * pad
	);
}

/**
 * Picks font sizes that fit a unit's text into a card of the given width/height.
 * Starts from the base sizes and only shrinks (body → subtitle → title) when the
 * content would overflow, down to per-role minimums.
 *
 * @param {{title: string, subtitle: string, body: string}} unit
 * @param {number} cardW - card width in PT
 * @param {number} cardH - card height in PT
 * @return {{title: number, subtitle: number, body: number, pad: number}}
 */
function fitFontSizes_(unit, cardW, cardH) {
	const pad = 8;
	const f = baseFontSizes_();
	let titleSize = f.title;
	let subSize = f.subtitle;
	let bodySize = f.body;

	const innerW = Math.max(20, cardW - 2 * pad);
	const innerH = Math.max(20, cardH - 2 * pad);

	const titleText = unit.title || "";
	const subText = unit.subtitle || " "; // reserve the blank subtitle row
	const bodyText = unit.body || "";

	// Shrink body (then subtitle, then title) until it fits or we hit the floor.
	for (let i = 0; i < 14; i++) {
		const blockH = estimateBlockHeight_(
			titleText,
			subText,
			bodyText,
			titleSize,
			subSize,
			bodySize,
			innerW,
		);
		if (blockH <= innerH) break;
		if (bodySize > 8) {
			bodySize -= 1;
		} else if (subSize > 8) {
			subSize -= 1;
		} else if (titleSize > 10) {
			titleSize -= 1;
			subSize = Math.min(subSize, Math.max(8, Math.round(titleSize * 0.7)));
			bodySize = Math.min(bodySize, Math.max(8, Math.round(titleSize * 0.78)));
		} else {
			break; // already at the floor
		}
	}

	return { title: titleSize, subtitle: subSize, body: bodySize, pad: pad };
}

/**
 * Rough estimate of the total vertical space (PT) a title/subtitle/body block
 * occupies inside a card, given font sizes and available width.
 * @return {number}
 */
function estimateBlockHeight_(
	title,
	subtitle,
	body,
	titleSize,
	subSize,
	bodySize,
	innerW,
) {
	const lineH = 1.25;
	const charsPerLine = (size) =>
		Math.max(1, Math.floor(innerW / (size * 0.55)));
	const linesFor = (text, size) =>
		text ? Math.max(1, Math.ceil(text.length / charsPerLine(size))) : 0;

	let h = 0;
	if (title) h += linesFor(title, titleSize) * titleSize * lineH + 2;
	if (subtitle) h += linesFor(subtitle, subSize) * subSize * lineH + 2;
	if (body) h += linesFor(body, bodySize) * bodySize * lineH + 2;
	return h;
}

/**
 * Pushes the Slides API requests that render one unit as a styled card into the
 * shared requests array.
 *
 * @param {Array} requests - shared batch request array
 * @param {string} pageId - target slide objectId
 * @param {{x:number,y:number,w:number,h:number}} pos - card rect in PT
 * @param {{title:string,subtitle:string,body:string}} unit
 * @param {{fillColor:string,borderColor:string,textColor:string,borderWidth:number}} style
 */
function buildUnitCardRequests_(requests, pageId, pos, unit, style) {
	const shapeId = "gridminter" + Utilities.getUuid().replace(/-/g, "");
	const sizes = fitFontSizes_(unit, pos.w, pos.h);

	// 1) The card shape itself (text box so it can hold the unit's text).
	requests.push(
		createTextBoxRequest_({
			pageId: pageId,
			id: shapeId,
			x: pos.x,
			y: pos.y,
			w: pos.w,
			h: pos.h,
		}),
	);

	// 2) Card fill + outline from the chosen default style.
	requests.push(
		fillOutlineRequest_(shapeId, {
			fillColor: style.fillColor,
			borderColor: style.borderColor,
			borderWidth: style.borderWidth,
			contentAlignment: "TOP",
		}),
	);

	// 3) Assemble the text and remember each segment's char range for styling.
	// Always emit three rows (title / subtitle / body) in order. The subtitle row
	// is kept even when empty so it renders as a blank line and cards with and
	// without a subtitle line up at the same heights.
	const rolesOrder = [
		{ text: unit.title || "", role: "title" },
		{ text: unit.subtitle || "", role: "subtitle" },
		{ text: unit.body || "", role: "body" },
	];
	const segments = [];
	let combined = "";
	for (let i = 0; i < rolesOrder.length; i++) {
		if (i > 0) combined += "\n";
		const start = combined.length;
		combined += rolesOrder[i].text;
		segments.push({
			start: start,
			end: combined.length,
			role: rolesOrder[i].role,
			text: rolesOrder[i].text,
		});
	}

	if (!combined.replace(/\n/g, "").length) return; // nothing real to insert

	requests.push(insertTextRequest_(shapeId, combined, 0));

	// 4) Per-segment text styling. Title & subtitle: bold, no italic, style text
	// color. Body: regular weight, black. Blank reserved lines are skipped.
	for (const seg of segments) {
		if (!seg.text.length) continue;
		const isTitle = seg.role === "title";
		const isSub = seg.role === "subtitle";
		const isBody = seg.role === "body";
		const fontSize = isTitle
			? sizes.title
			: isSub
				? sizes.subtitle
				: sizes.body;
		requests.push(
			textStyleRequest_(
				shapeId,
				{ start: seg.start, end: seg.end },
				{
					color: isBody ? "#000000" : style.textColor,
					bold: isTitle || isSub,
					italic: false,
					fontFamily: main_font_family,
					fontSize: fontSize,
				},
			),
		);

		// Give the title and subtitle a 1.5× line height (lineSpacing is a %).
		if (isTitle || isSub) {
			requests.push(
				paragraphStyleRequest_(
					shapeId,
					{ start: seg.start, end: seg.end },
					{ lineSpacing: 150 },
					"lineSpacing",
				),
			);
		}
	}

	// 5) Left-align the whole card.
	requests.push(
		paragraphStyleRequest_(shapeId, "ALL", { alignment: "START" }, "alignment"),
	);
}

/**
 * Inserts grid-unit cards onto the current slide, spilling overflow onto newly
 * appended slides. Everything is sent in a single batchUpdate.
 *
 * @param {{units: Array<{title:string,subtitle:string,body:string}>, rows: number,
 *   cols: number, styleNumber: number, left?: number, top?: number}} payload
 * @return {{success: boolean, error?: string, warnings?: string[]}}
 */
function insertGridIntoSlide(payload) {
	try {
		const p = payload || {};
		const units = p.units || [];
		if (!units.length) {
			return { success: false, error: "No units to insert." };
		}

		let rows = p.rows > 0 ? Math.floor(p.rows) : 0;
		let cols = p.cols > 0 ? Math.floor(p.cols) : 0;
		if (!rows || !cols) {
			const sug = suggestGridLayout_(units.length);
			rows = rows || sug.rows;
			cols = cols || sug.cols;
		}
		const cap = rows * cols;

		const styleNumber = p.styleNumber > 0 ? p.styleNumber : 1;
		const styles = getStyleDefinitions();
		const style = styles[styleNumber] || styles[1];

		const presentation = SlidesApp.getActivePresentation();
		const pageW = presentation.getPageWidth();
		const pageH = presentation.getPageHeight();

		// Resolve the target slide (fall back to the first slide).
		let firstSlide = null;
		try {
			firstSlide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			firstSlide = presentation.getSlides()[0];
		}
		if (!firstSlide) return { success: false, error: "No slide available." };

		const positionOpts = {};
		if (p.left != null) positionOpts.left = p.left;
		if (p.top != null) positionOpts.top = p.top;

		// Row height = the tallest unit's natural content height (uniform cards),
		// so the grid hugs its content instead of stretching to the slide bottom.
		const margin = 30;
		const gap = 15;
		const left = p.left != null ? p.left : margin;
		const cellW = (pageW - left - margin - (cols - 1) * gap) / cols;
		let contentRowH = 0;
		for (const u of units) {
			const h = naturalUnitHeight_(u, cellW);
			if (h > contentRowH) contentRowH = h;
		}
		// Add vertical breathing room so cards aren't cramped to the text.
		positionOpts.rowHeight = contentRowH + 36;

		// Chunk units into pages of `cap`. The first chunk uses the current slide;
		// later chunks each get a fresh BLANK slide created INSIDE the same batch
		// (a createSlide request with a client-chosen objectId). Never mix
		// SlidesApp.appendSlide() with the REST batchUpdate here: SlidesApp writes
		// are buffered (and SlidesApp has no flush()), so batchUpdate would not
		// find the new page ("The page ... could not be found").
		const requests = [];
		const warnings = [];
		let extraSlides = 0;
		for (let offset = 0; offset < units.length; offset += cap) {
			const chunk = units.slice(offset, offset + cap);
			let pageId;
			if (offset === 0) {
				pageId = firstSlide.getObjectId();
			} else {
				pageId = "gridminterpage" + Utilities.getUuid().replace(/-/g, "");
				requests.push({
					createSlide: {
						objectId: pageId,
						slideLayoutReference: { predefinedLayout: "BLANK" },
					},
				});
				extraSlides++;
			}
			const positions = computeGridPositions_(
				rows,
				cols,
				pageW,
				pageH,
				positionOpts,
			);
			chunk.forEach((unit, i) => {
				buildUnitCardRequests_(requests, pageId, positions[i], unit, style);
			});
		}

		if (extraSlides > 0) {
			warnings.push(
				extraSlides +
					" 張新投影片已新增以容納多餘的 unit（每頁 " +
					cap +
					" 格）。",
			);
		}

		if (requests.length) {
			try {
				Slides.Presentations.batchUpdate({ requests }, presentation.getId());
			} catch (batchErr) {
				return { success: false, error: batchErr.message };
			}
		}

		return { success: true, warnings: warnings };
	} catch (e) {
		console.error("Error inserting grid: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Auto Minter generate wrapper: adapts the (context, hints) auto-minter call
 * signature onto generateGridMarkdownFromContext(context, hasSubtitle).
 *
 * @param {string} context - Arbitrary text to summarize into grid units.
 * @param {{hasSubtitle?: boolean}} [hints] - optional router hints
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function autoGenerateGrid_(context, hints) {
	return generateGridMarkdownFromContext(
		context,
		!!(hints && hints.hasSubtitle),
	);
}

/**
 * Auto Minter adapter: turns AI-generated grid-unit Markdown into the payload
 * insertGridIntoSlide() accepts. Rows/cols come from the hints when both are
 * valid positive integers, else from suggestGridLayout_().
 *
 * @param {string} generatedText - raw LLM output (unit markdown)
 * @param {{rows?: number, cols?: number, styleNumber?: number}} [hints]
 * @return {{units: Array<{title:string,subtitle:string,body:string}>, rows: number,
 *   cols: number, styleNumber: number}|null} null when parsing yields zero usable units
 */
function autoBuildGridPayload_(generatedText, hints) {
	const h = hints || {};
	const units = parseGridUnits_(generatedText);
	if (!units.length) return null;

	const hintRows = parseInt(h.rows, 10);
	const hintCols = parseInt(h.cols, 10);
	let rows;
	let cols;
	if (hintRows > 0 && hintCols > 0) {
		rows = hintRows;
		cols = hintCols;
	} else {
		const sug = suggestGridLayout_(units.length);
		rows = sug.rows;
		cols = sug.cols;
	}

	const sn = parseInt(h.styleNumber, 10);
	const styleNumber = sn >= 1 && sn <= 6 ? sn : 1;

	return { units: units, rows: rows, cols: cols, styleNumber: styleNumber };
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "grid",
	label: "網格卡片",
	emoji: "🔳",
	order: 20,
	whenToUse:
		"3-9 parallel concepts/categories each with a title and short description; general-purpose card grid",
	hintsSpec:
		'{"rows":number,"cols":number,"hasSubtitle":boolean,"styleNumber":1-6}',
	generate: "autoGenerateGrid_",
	buildPayload: "autoBuildGridPayload_",
	insert: "insertGridIntoSlide",
	previewPartial: "src/components/grid-minter/preview",
	previewKind: "cards",
	precheck: "",
	options: [
		{
			name: "rows",
			label: "列 Rows",
			type: "number",
			min: 1,
			max: 9,
			placeholder: "auto",
		},
		{
			name: "cols",
			label: "欄 Cols",
			type: "number",
			min: 1,
			max: 9,
			placeholder: "auto",
		},
		{
			name: "styleNumber",
			label: "卡片樣式",
			type: "select",
			default: 1,
			choices: [
				{ value: 1, label: "1 · 白底品牌框" },
				{ value: 2, label: "2 · 品牌色底白字" },
				{ value: 3, label: "3 · 淺灰底" },
				{ value: 4, label: "4 · 白底強調框" },
				{ value: 5, label: "5 · 強調色底白字" },
				{ value: 6, label: "6 · 白底無框" },
			],
		},
		{
			name: "hasSubtitle",
			label: "含副標 (##)",
			type: "checkbox",
			default: true,
			regenerate: true,
		},
	],
});
