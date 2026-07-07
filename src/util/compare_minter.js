// TODO(shared-migration): adopt the shared core like grid_minter.js/kpi_minter.js —
//   • replace the typeof *_color fallback block with getThemeColors() (shared/theme_colors.js)
//   • replace inline createShape/updateShapeProperties/updateTextStyle literals with
//     shared/shape_requests.js builders + rgbColor_ (hexToRgbColor_ is now a shared alias)
/**
 * Server-side core for the Comparison (對照) Minter dialog.
 *
 * Turns "---"-separated markdown blocks into N comparison columns, then renders
 * each column as a side-by-side card: a colored title header bar on top of a
 * body box that lists the column's bullet points. Everything is inserted onto
 * the current slide in a single Slides.Presentations.batchUpdate.
 *
 * Mirrors the Grid Minter pattern (src/util/grid_minter.js):
 *  - Columns are separated by a line containing only "---" (same parser shape).
 *  - Insertion batches every request into one batchUpdate.
 *  - Templates resolve from the global palette (main_color / accent_color) so
 *    they track the configured theme.
 *
 * The column markdown format (also documented in the dialog):
 *   # 方案 A
 *   成本低
 *   導入快
 *   ---
 *   # 方案 B
 *   效能高
 *   可擴充
 * Columns are separated by a line containing only "---". Within a column, the
 * first "# " line is the title and every remaining non-heading line is a bullet.
 */

/**
 * Groq system prompt for turning free-form context into comparison-column
 * markdown that parseCompareColumns_() understands.
 */
var COMPARE_AI_SYSTEM_PROMPT = [
	"You convert the user's content into a side-by-side comparison.",
	"Output EXACTLY this format and nothing else:",
	"- Each column is a block that starts with a title line beginning with '# '.",
	"- After the title line, put one bullet point per line (plain text, NO leading bullet character).",
	"- Separate consecutive columns with a line that contains only '---'.",
	"Example:",
	"# 方案 A",
	"成本較低",
	"導入快速",
	"---",
	"# 方案 B",
	"效能較高",
	"可擴充性佳",
	"Strict rules:",
	"- No preamble, no explanation, no closing remarks, no code fences.",
	"- The very first line MUST start with '# '.",
	"- Produce EXACTLY 2 columns unless the input clearly implies more.",
	"- Give each column 2–5 bullet points.",
	"- Keep every line concise; do not invent facts that aren't implied by the input.",
].join("\n");

/**
 * Generates comparison-column Markdown from free-form context via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateCompareFromContext(context) {
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

	return callGroq_(COMPARE_AI_SYSTEM_PROMPT, text, {
		maxTokens: 2000,
		temperature: 0.3,
	});
}

/**
 * Single source of truth for comparison templates. Each template carries the
 * per-column header fill/text and body border/text so columns share one look.
 * Colors resolve from the global palette where applicable.
 *
 * @return {Array<{id,name,headerFill,headerText,bodyFill,bodyBorder,bodyText}>}
 */
function buildCompareTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const green = "#2E7D32";
	const red = "#C0392B";
	return [
		{
			id: "neutral",
			name: "Neutral",
			// Both columns share the main color header.
			columns: [
				{
					headerFill: main,
					headerText: "#FFFFFF",
					bodyFill: "#FFFFFF",
					bodyBorder: main,
					bodyText: "#000000",
				},
			],
		},
		{
			id: "pros-cons",
			name: "Pros / Cons",
			// First column green (pros), second column red (cons), then repeat.
			columns: [
				{
					headerFill: green,
					headerText: "#FFFFFF",
					bodyFill: "#FFFFFF",
					bodyBorder: green,
					bodyText: "#000000",
				},
				{
					headerFill: red,
					headerText: "#FFFFFF",
					bodyFill: "#FFFFFF",
					bodyBorder: red,
					bodyText: "#000000",
				},
			],
		},
		{
			id: "a-vs-b",
			name: "A vs B",
			// First column main, second column accent, then repeat.
			columns: [
				{
					headerFill: main,
					headerText: "#FFFFFF",
					bodyFill: "#FFFFFF",
					bodyBorder: main,
					bodyText: "#000000",
				},
				{
					headerFill: accent,
					headerText: "#FFFFFF",
					bodyFill: "#FFFFFF",
					bodyBorder: accent,
					bodyText: "#000000",
				},
			],
		},
	];
}

/**
 * Returns the comparison templates for client-side preview. Called from the
 * dialog through google.script.run.
 * @return {Array<Object>}
 */
function getCompareTemplates() {
	return buildCompareTemplates_();
}

/**
 * Resolves the per-column style for column index `i` from a template. Templates
 * may define fewer column styles than there are columns; the list cycles.
 *
 * @param {Object} tpl - a template from buildCompareTemplates_()
 * @param {number} i - zero-based column index
 * @return {{headerFill,headerText,bodyFill,bodyBorder,bodyText}}
 */
function compareColumnStyle_(tpl, i) {
	const cols = (tpl && tpl.columns) || [];
	if (!cols.length) {
		return {
			headerFill: "#3D6869",
			headerText: "#FFFFFF",
			bodyFill: "#FFFFFF",
			bodyBorder: "#3D6869",
			bodyText: "#000000",
		};
	}
	return cols[i % cols.length];
}

/**
 * Parses comparison markdown into an array of {title, points}.
 * Columns are separated by a line containing only "---" (---, ----, …). Within
 * a column, the first "# " line is the title and every remaining non-heading,
 * non-blank line is a bullet point.
 *
 * @param {string} markdown
 * @return {Array<{title: string, points: string[]}>}
 */
function parseCompareColumns_(markdown) {
	const text = (markdown || "").replace(/\r\n/g, "\n").trim();
	if (!text) return [];

	// Split on a line that is only dashes (---, ----, …), tolerating whitespace.
	const blocks = text
		.split(/\n\s*-{3,}\s*\n/)
		.map((b) => b.trim())
		.filter((b) => b.length);

	const columns = [];
	for (const block of blocks) {
		const lines = block.split("\n");
		let title = "";
		const points = [];
		for (const raw of lines) {
			const line = raw.trim();
			if (!line) continue;
			const h1 = line.match(/^#\s+(.+)$/);
			if (h1) {
				if (!title) title = h1[1].trim();
			} else {
				// Strip a leading bullet marker if the user typed one.
				points.push(line.replace(/^[•\-*]\s+/, ""));
			}
		}
		if (title || points.length) {
			columns.push({ title: title, points: points });
		}
	}
	return columns;
}

/**
 * Pushes the Slides API requests that render one comparison column (a colored
 * title header bar + a body box of bullet points) into the shared requests array.
 *
 * @param {Array} requests - shared batch request array
 * @param {string} pageId - target slide objectId
 * @param {{x:number,y:number,w:number,headerH:number,bodyH:number}} pos
 * @param {{title:string, points:string[]}} column
 * @param {{headerFill,headerText,bodyFill,bodyBorder,bodyText}} style
 */
function buildCompareColumnRequests_(requests, pageId, pos, column, style) {
	const uid = Utilities.getUuid().replace(/-/g, "");
	const headerId = "cmphead" + uid;
	const bodyId = "cmpbody" + uid;
	const font =
		(typeof main_font_family !== "undefined" && main_font_family) ||
		"Source Sans Pro";

	// --- Header bar -----------------------------------------------------------
	requests.push({
		createShape: {
			objectId: headerId,
			shapeType: "RECTANGLE",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: pos.w, unit: "PT" },
					height: { magnitude: pos.headerH, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: pos.x,
					translateY: pos.y,
					unit: "PT",
				},
			},
		},
	});
	requests.push({
		updateShapeProperties: {
			objectId: headerId,
			shapeProperties: {
				shapeBackgroundFill: {
					solidFill: { color: { rgbColor: hexToRgbColor_(style.headerFill) } },
				},
				outline: {
					outlineFill: {
						solidFill: {
							color: { rgbColor: hexToRgbColor_(style.headerFill) },
						},
					},
					weight: { magnitude: 1, unit: "PT" },
					dashStyle: "SOLID",
				},
				contentAlignment: "MIDDLE",
			},
			fields:
				"shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color,outline.weight,outline.dashStyle,contentAlignment",
		},
	});
	const titleText = column.title || "";
	if (titleText) {
		requests.push({
			insertText: { objectId: headerId, insertionIndex: 0, text: titleText },
		});
		requests.push({
			updateTextStyle: {
				objectId: headerId,
				textRange: { type: "ALL" },
				style: {
					foregroundColor: {
						opaqueColor: { rgbColor: hexToRgbColor_(style.headerText) },
					},
					bold: true,
					fontFamily: font,
					fontSize: { magnitude: 16, unit: "PT" },
				},
				fields: "foregroundColor,bold,fontFamily,fontSize",
			},
		});
		requests.push({
			updateParagraphStyle: {
				objectId: headerId,
				textRange: { type: "ALL" },
				style: { alignment: "CENTER" },
				fields: "alignment",
			},
		});
	}

	// --- Body box -------------------------------------------------------------
	const bodyY = pos.y + pos.headerH;
	requests.push({
		createShape: {
			objectId: bodyId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: pos.w, unit: "PT" },
					height: { magnitude: pos.bodyH, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: pos.x,
					translateY: bodyY,
					unit: "PT",
				},
			},
		},
	});
	requests.push({
		updateShapeProperties: {
			objectId: bodyId,
			shapeProperties: {
				shapeBackgroundFill: {
					solidFill: { color: { rgbColor: hexToRgbColor_(style.bodyFill) } },
				},
				outline: {
					outlineFill: {
						solidFill: {
							color: { rgbColor: hexToRgbColor_(style.bodyBorder) },
						},
					},
					weight: { magnitude: 1, unit: "PT" },
					dashStyle: "SOLID",
				},
				contentAlignment: "TOP",
			},
			fields:
				"shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color,outline.weight,outline.dashStyle,contentAlignment",
		},
	});

	const points = column.points || [];
	if (points.length) {
		const bodyText = points.map((p) => "• " + p).join("\n");
		requests.push({
			insertText: { objectId: bodyId, insertionIndex: 0, text: bodyText },
		});
		requests.push({
			updateTextStyle: {
				objectId: bodyId,
				textRange: { type: "ALL" },
				style: {
					foregroundColor: {
						opaqueColor: { rgbColor: hexToRgbColor_(style.bodyText) },
					},
					fontFamily: font,
					fontSize: { magnitude: 13, unit: "PT" },
				},
				fields: "foregroundColor,fontFamily,fontSize",
			},
		});
		requests.push({
			updateParagraphStyle: {
				objectId: bodyId,
				textRange: { type: "ALL" },
				style: {
					alignment: "START",
					lineSpacing: 130,
					spaceBelow: { magnitude: 4, unit: "PT" },
				},
				fields: "alignment,lineSpacing,spaceBelow",
			},
		});
	}
}

/**
 * Inserts comparison columns as side-by-side cards onto the current slide.
 * Columns divide the usable page width evenly (auto-width by count), start at
 * Y ≈ 120, and each card is a colored title header bar over a bullet body box.
 * Everything is sent in a single batchUpdate.
 *
 * @param {{columns: Array<{title:string, points:string[]}>, templateId: string}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertCompareIntoSlide(payload) {
	try {
		const p = payload || {};
		const columns = p.columns || [];
		if (!columns.length) {
			return { success: false, error: "No columns to insert." };
		}

		const templates = buildCompareTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}

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
		const pageId = slide.getObjectId();

		// Layout: even columns across the usable width, start near the top.
		const margin = 30;
		const gap = 18;
		const top = 120;
		const headerH = 36;
		const n = columns.length;
		const usableW = pageW - 2 * margin;
		const colW = (usableW - (n - 1) * gap) / n;
		const bodyH = Math.max(80, pageH - top - headerH - margin);

		const requests = [];
		for (let i = 0; i < n; i++) {
			const x = margin + i * (colW + gap);
			const style = compareColumnStyle_(tpl, i);
			buildCompareColumnRequests_(
				requests,
				pageId,
				{ x: x, y: top, w: colW, headerH: headerH, bodyH: bodyH },
				columns[i],
				style,
			);
		}

		if (requests.length) {
			try {
				Slides.Presentations.batchUpdate({ requests }, presentation.getId());
			} catch (batchErr) {
				return { success: false, error: batchErr.message };
			}
		}

		return { success: true };
	} catch (e) {
		console.error("Error inserting comparison: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Auto Minter adapter: turns AI-generated comparison markdown ("---"-separated
 * column blocks) into the payload insertCompareIntoSlide() accepts.
 *
 * @param {string} generatedText - raw LLM output from generateCompareFromContext()
 * @param {{templateId?: string}} [hints] - optional router hints
 * @return {{columns: Array<{title:string, points:string[]}>, templateId: string}|null}
 *   null when parsing yields zero usable columns
 */
function autoBuildComparePayload_(generatedText, hints) {
	const h = hints || {};
	const columns = parseCompareColumns_(generatedText);
	if (!columns.length) return null;

	const templates = buildCompareTemplates_();
	let templateId = templates[0].id;
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === h.templateId) templateId = templates[i].id;
	}

	return { columns: columns, templateId: templateId };
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "compare",
	label: "對照欄",
	emoji: "⚖",
	order: 60,
	whenToUse:
		"comparing exactly 2-3 alternatives/options/arms side by side, each with bullet points",
	hintsSpec: "",
	generate: "generateCompareFromContext",
	buildPayload: "autoBuildComparePayload_",
	insert: "insertCompareIntoSlide",
	previewPartial: "src/components/compare-minter/preview",
	previewKind: "columns",
	precheck: "",
	options: [
		{
			name: "templateId",
			label: "範本",
			type: "select",
			choicesFrom: "getCompareTemplates",
		},
	],
});
