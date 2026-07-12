// TODO(shared-migration): adopt the shared core like grid_minter.js/kpi_minter.js —
//   • replace the typeof *_color fallback block with getThemeColors() (shared/theme_colors.js)
//   • replace inline createShape/updateShapeProperties/updateTextStyle literals with
//     shared/shape_requests.js builders + rgbColor_ (hexToRgbColor_ is now a shared alias)
/**
 * Server-side core for the Takeaways Minter dialog (重點摘要 / Key Takeaways 鑄造器).
 *
 * Turns a heading + a list of key points (one per line, `title | desc` with the
 * description optional) into a styled "key takeaways" block on the current
 * slide: a bold, larger heading in the theme color, followed by the points
 * rendered either as numbered cards (a grid of cards, each with a number badge +
 * title + description) or as a vertical checkmark list (✓ + title + description).
 *
 * Mirrors the Grid Minter pattern (src/util/grid_minter.js):
 *  - Card layout reuses grid-style positioning (computeGridPositions_-style math).
 *  - Insertion batches every request into a single Slides.Presentations.batchUpdate.
 *  - Colors resolve from the global palette (main_color / accent_color) so they
 *    track the configured theme.
 *
 * The point line format (also documented in the dialog):
 *   First takeaway | a short supporting description
 *   Second takeaway
 *   ...
 * One point per line; the description after the first "|" is optional.
 */

/**
 * Single source of truth for takeaways templates. Each template chooses a render
 * style (numbered cards vs. checkmark list) and a color theme (main / accent /
 * green) that drives the accent color, card fill/border, and the number/check
 * badge styling. Colors resolve from the global palette so they track the theme.
 *
 * style: 'numbered' = grid of cards with a number badge; 'checklist' = vertical
 * list with a ✓ marker per row.
 *
 * @return {Array<{id,name,style,accent,cardFill,cardBorder,badgeFill,badgeText,titleText,descText}>}
 */
function buildTakeawaysTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const green = "#2E7D32";

	function theme(idBase, nameBase, style, color, cardFill, cardBorder) {
		return {
			id: idBase,
			name: nameBase,
			style: style,
			accent: color,
			cardFill: cardFill,
			cardBorder: cardBorder,
			badgeFill: color,
			badgeText: "#FFFFFF",
			titleText: color,
			descText: "#333333",
		};
	}

	return [
		// Numbered cards — three color themes.
		theme(
			"numbered-main",
			"Numbered · Main",
			"numbered",
			main,
			"#FFFFFF",
			main,
		),
		theme(
			"numbered-accent",
			"Numbered · Accent",
			"numbered",
			accent,
			"#FFFFFF",
			accent,
		),
		theme(
			"numbered-green",
			"Numbered · Green",
			"numbered",
			green,
			"#FFFFFF",
			green,
		),
		// Colored cards (filled background) — numbered, three themes.
		{
			id: "card-main",
			name: "Colored Card · Main",
			style: "numbered",
			accent: main,
			cardFill: "#E7EAE7",
			cardBorder: main,
			badgeFill: main,
			badgeText: "#FFFFFF",
			titleText: main,
			descText: "#333333",
		},
		{
			id: "card-accent",
			name: "Colored Card · Accent",
			style: "numbered",
			accent: accent,
			cardFill: "#FCEAD2",
			cardBorder: accent,
			badgeFill: accent,
			badgeText: "#FFFFFF",
			titleText: "#B26A00",
			descText: "#333333",
		},
		{
			id: "card-green",
			name: "Colored Card · Green",
			style: "numbered",
			accent: green,
			cardFill: "#E7F4EC",
			cardBorder: green,
			badgeFill: green,
			badgeText: "#FFFFFF",
			titleText: green,
			descText: "#333333",
		},
		// Checkmark list — three color themes.
		theme("check-main", "Checklist · Main", "checklist", main, "#FFFFFF", main),
		theme(
			"check-accent",
			"Checklist · Accent",
			"checklist",
			accent,
			"#FFFFFF",
			accent,
		),
		theme(
			"check-green",
			"Checklist · Green",
			"checklist",
			green,
			"#FFFFFF",
			green,
		),
	];
}

/**
 * Returns the takeaways templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getTakeawaysTemplates() {
	return buildTakeawaysTemplates_();
}

/**
 * Groq system prompt for turning free-form context into takeaway points.
 * Output is ONLY key-point lines, one per line, in the exact `title | desc`
 * format the points textarea expects (the description after "|" is optional).
 */
var TAKEAWAYS_AI_SYSTEM_PROMPT = [
	"You convert the user's content into a short list of key takeaways.",
	"Output ONLY the key-point lines and nothing else.",
	"Each line is ONE takeaway in the EXACT format: title | desc",
	"- 'title' is a short label (1–4 words).",
	"- ' | desc' is ONE concise supporting phrase; the description is optional.",
	"Example output:",
	"更快 | 交付時間從數週縮短到數天",
	"更穩 | 自動化檢查減少回歸錯誤",
	"可擴充 | 不需重寫即可承載十倍流量",
	"Strict rules:",
	"- No preamble, no explanation, no closing remarks, no code fences.",
	"- Produce between 3 and 5 takeaways.",
	"- Keep titles short; give each at most one concise description phrase.",
	"- Do not invent facts that aren't implied by the input.",
].join("\n");

/**
 * Generates takeaway point lines (`title | desc`, one per line) from free-form
 * context via Groq. Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateTakeawaysFromContext(context) {
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

	return callGroq_(TAKEAWAYS_AI_SYSTEM_PROMPT, text, {
		maxTokens: 1500,
		temperature: 0.3,
	});
}

/**
 * Parses the points textarea into an array of {title, desc}. One point per line;
 * the title and an optional description are separated by the first "|".
 *
 * @param {string} text
 * @return {Array<{title: string, desc: string}>}
 */
function parseTakeawayPoints_(text) {
	const raw = (text || "").replace(/\r\n/g, "\n");
	const points = [];
	const lines = raw.split("\n");
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i].trim();
		if (!line) continue;
		const idx = line.indexOf("|");
		let title;
		let desc;
		if (idx >= 0) {
			title = line.slice(0, idx).trim();
			desc = line.slice(idx + 1).trim();
		} else {
			title = line;
			desc = "";
		}
		if (title || desc) points.push({ title: title, desc: desc });
	}
	return points;
}

/**
 * Suggests a column count for numbered cards given a point count. Keeps grids
 * roughly landscape to match 16:9 slides.
 *
 * @param {number} n
 * @return {number} column count
 */
function suggestTakeawaysCols_(n) {
	if (n <= 1) return 1;
	if (n === 2) return 2;
	if (n === 4) return 2;
	if (n <= 9) return 3;
	return 3;
}

/**
 * Pushes the Slides API requests for the heading text box into the shared
 * requests array, and returns the heading's bottom Y (PT) so the points block
 * can start beneath it.
 *
 * @param {Array} requests - shared batch request array
 * @param {string} pageId
 * @param {string} heading
 * @param {Object} tpl - resolved template
 * @param {{left:number, top:number, width:number}} box
 * @return {number} the Y (PT) just below the heading
 */
function buildHeadingRequests_(requests, pageId, heading, tpl, box) {
	const font =
		(typeof main_font_family !== "undefined" && main_font_family) ||
		"Source Sans Pro";
	const headingH = 40;
	const shapeId = "takeawayhead" + Utilities.getUuid().replace(/-/g, "");

	requests.push({
		createShape: {
			objectId: shapeId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: box.width, unit: "PT" },
					height: { magnitude: headingH, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: box.left,
					translateY: box.top,
					unit: "PT",
				},
			},
		},
	});
	requests.push({
		insertText: { objectId: shapeId, insertionIndex: 0, text: heading },
	});
	requests.push({
		updateTextStyle: {
			objectId: shapeId,
			textRange: { type: "ALL" },
			style: {
				foregroundColor: {
					opaqueColor: { rgbColor: hexToRgbColor_(tpl.accent) },
				},
				bold: true,
				fontFamily: font,
				fontSize: { magnitude: 24, unit: "PT" },
			},
			fields: "foregroundColor,bold,fontFamily,fontSize",
		},
	});
	requests.push({
		updateParagraphStyle: {
			objectId: shapeId,
			textRange: { type: "ALL" },
			style: { alignment: "START" },
			fields: "alignment",
		},
	});

	return box.top + headingH + 12;
}

/**
 * Pushes the Slides API requests rendering one numbered card (number badge +
 * title + description) into the shared requests array.
 *
 * @param {Array} requests
 * @param {string} pageId
 * @param {{x:number,y:number,w:number,h:number}} pos
 * @param {number} number - 1-based card number
 * @param {{title:string,desc:string}} point
 * @param {Object} tpl - resolved template
 */
function buildNumberedCardRequests_(requests, pageId, pos, number, point, tpl) {
	const font =
		(typeof main_font_family !== "undefined" && main_font_family) ||
		"Source Sans Pro";
	const uid = Utilities.getUuid().replace(/-/g, "");
	const cardId = "takeawaycard" + uid;
	const badgeId = "takeawaybadge" + uid;

	// 1) Card shape (text box holding the title + description).
	requests.push({
		createShape: {
			objectId: cardId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: pos.w, unit: "PT" },
					height: { magnitude: pos.h, unit: "PT" },
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
			objectId: cardId,
			shapeProperties: {
				shapeBackgroundFill: {
					solidFill: { color: { rgbColor: hexToRgbColor_(tpl.cardFill) } },
				},
				outline: {
					outlineFill: {
						solidFill: { color: { rgbColor: hexToRgbColor_(tpl.cardBorder) } },
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

	// Card text: title (bold, theme color) + description (regular). Leave room at
	// the top for the number badge that floats above-left of the card.
	const title = point.title || "";
	const desc = point.desc || "";
	let combined = title;
	const titleEnd = combined.length;
	if (desc) combined += "\n" + desc;

	if (combined.length) {
		requests.push({
			insertText: { objectId: cardId, insertionIndex: 0, text: combined },
		});
		if (titleEnd > 0) {
			requests.push({
				updateTextStyle: {
					objectId: cardId,
					textRange: { type: "FIXED_RANGE", startIndex: 0, endIndex: titleEnd },
					style: {
						foregroundColor: {
							opaqueColor: { rgbColor: hexToRgbColor_(tpl.titleText) },
						},
						bold: true,
						fontFamily: font,
						fontSize: { magnitude: 14, unit: "PT" },
					},
					fields: "foregroundColor,bold,fontFamily,fontSize",
				},
			});
		}
		if (desc) {
			requests.push({
				updateTextStyle: {
					objectId: cardId,
					textRange: {
						type: "FIXED_RANGE",
						startIndex: titleEnd + 1,
						endIndex: combined.length,
					},
					style: {
						foregroundColor: {
							opaqueColor: { rgbColor: hexToRgbColor_(tpl.descText) },
						},
						bold: false,
						fontFamily: font,
						fontSize: { magnitude: 11, unit: "PT" },
					},
					fields: "foregroundColor,bold,fontFamily,fontSize",
				},
			});
		}
		requests.push({
			updateParagraphStyle: {
				objectId: cardId,
				textRange: { type: "ALL" },
				style: { alignment: "START" },
				fields: "alignment",
			},
		});
	}

	// 2) Number badge — a small circle floating at the card's top-left corner.
	const badgeD = 24;
	requests.push({
		createShape: {
			objectId: badgeId,
			shapeType: "ELLIPSE",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: badgeD, unit: "PT" },
					height: { magnitude: badgeD, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: pos.x - badgeD / 2,
					translateY: pos.y - badgeD / 2,
					unit: "PT",
				},
			},
		},
	});
	requests.push({
		updateShapeProperties: {
			objectId: badgeId,
			shapeProperties: {
				shapeBackgroundFill: {
					solidFill: { color: { rgbColor: hexToRgbColor_(tpl.badgeFill) } },
				},
				outline: {
					outlineFill: {
						solidFill: { color: { rgbColor: hexToRgbColor_(tpl.badgeFill) } },
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
	requests.push({
		insertText: { objectId: badgeId, insertionIndex: 0, text: String(number) },
	});
	requests.push({
		updateTextStyle: {
			objectId: badgeId,
			textRange: { type: "ALL" },
			style: {
				foregroundColor: {
					opaqueColor: { rgbColor: hexToRgbColor_(tpl.badgeText) },
				},
				bold: true,
				fontFamily: font,
				fontSize: { magnitude: 12, unit: "PT" },
			},
			fields: "foregroundColor,bold,fontFamily,fontSize",
		},
	});
	requests.push({
		updateParagraphStyle: {
			objectId: badgeId,
			textRange: { type: "ALL" },
			style: { alignment: "CENTER" },
			fields: "alignment",
		},
	});

	return [cardId, badgeId];
}

/**
 * Pushes the Slides API requests rendering one checkmark list row (✓ + title +
 * description) into the shared requests array.
 *
 * @param {Array} requests
 * @param {string} pageId
 * @param {{x:number,y:number,w:number,h:number}} pos
 * @param {{title:string,desc:string}} point
 * @param {Object} tpl - resolved template
 */
function buildChecklistRowRequests_(requests, pageId, pos, point, tpl) {
	const font =
		(typeof main_font_family !== "undefined" && main_font_family) ||
		"Source Sans Pro";
	const uid = Utilities.getUuid().replace(/-/g, "");
	const checkId = "takeawaycheck" + uid;
	const textId = "takeawayrow" + uid;
	const checkW = 24;
	const gap = 8;

	// 1) Check marker (✓ in the accent color).
	requests.push({
		createShape: {
			objectId: checkId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: checkW, unit: "PT" },
					height: { magnitude: pos.h, unit: "PT" },
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
		insertText: { objectId: checkId, insertionIndex: 0, text: "✓" },
	});
	requests.push({
		updateTextStyle: {
			objectId: checkId,
			textRange: { type: "ALL" },
			style: {
				foregroundColor: {
					opaqueColor: { rgbColor: hexToRgbColor_(tpl.accent) },
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
			objectId: checkId,
			textRange: { type: "ALL" },
			style: { alignment: "START" },
			fields: "alignment",
		},
	});

	// 2) Title + description text box to the right of the check.
	const textX = pos.x + checkW + gap;
	const textW = Math.max(40, pos.w - checkW - gap);
	requests.push({
		createShape: {
			objectId: textId,
			shapeType: "TEXT_BOX",
			elementProperties: {
				pageObjectId: pageId,
				size: {
					width: { magnitude: textW, unit: "PT" },
					height: { magnitude: pos.h, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: textX,
					translateY: pos.y,
					unit: "PT",
				},
			},
		},
	});

	const title = point.title || "";
	const desc = point.desc || "";
	let combined = title;
	const titleEnd = combined.length;
	if (desc) combined += "\n" + desc;

	if (combined.length) {
		requests.push({
			insertText: { objectId: textId, insertionIndex: 0, text: combined },
		});
		if (titleEnd > 0) {
			requests.push({
				updateTextStyle: {
					objectId: textId,
					textRange: { type: "FIXED_RANGE", startIndex: 0, endIndex: titleEnd },
					style: {
						foregroundColor: {
							opaqueColor: { rgbColor: hexToRgbColor_(tpl.titleText) },
						},
						bold: true,
						fontFamily: font,
						fontSize: { magnitude: 14, unit: "PT" },
					},
					fields: "foregroundColor,bold,fontFamily,fontSize",
				},
			});
		}
		if (desc) {
			requests.push({
				updateTextStyle: {
					objectId: textId,
					textRange: {
						type: "FIXED_RANGE",
						startIndex: titleEnd + 1,
						endIndex: combined.length,
					},
					style: {
						foregroundColor: {
							opaqueColor: { rgbColor: hexToRgbColor_(tpl.descText) },
						},
						bold: false,
						fontFamily: font,
						fontSize: { magnitude: 11, unit: "PT" },
					},
					fields: "foregroundColor,bold,fontFamily,fontSize",
				},
			});
		}
		requests.push({
			updateParagraphStyle: {
				objectId: textId,
				textRange: { type: "ALL" },
				style: { alignment: "START" },
				fields: "alignment",
			},
		});
	}

	return [checkId, textId];
}

/**
 * Inserts a "key takeaways" block (heading + points) onto the current slide.
 * Everything is sent in a single batchUpdate. The pieces are grouped per row /
 * card and the whole block (heading + points) is grouped together at the end.
 *
 * @param {{heading: string, points: Array<{title:string,desc:string}>,
 *   templateId: string}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertTakeawaysIntoSlide(payload) {
	try {
		const p = payload || {};
		const points = (p.points || []).filter(function (pt) {
			return pt && (pt.title || pt.desc);
		});
		if (!points.length) {
			return { success: false, error: "No takeaway points to insert." };
		}

		const templates = buildTakeawaysTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}

		const heading =
			(p.heading != null ? String(p.heading).trim() : "") || "Key Takeaways";

		const presentation = SlidesApp.getActivePresentation();
		const pageW = presentation.getPageWidth();

		// Resolve the target slide (payload.pageObjectId → selection → first).
		let slide = resolveMinterTargetSlide_(presentation, p.pageObjectId);
		if (!slide) return { success: false, error: "No slide available." };
		const pageId = slide.getObjectId();

		const margin = 40;
		const top = 70;
		const blockLeft = margin;
		const blockWidth = pageW - 2 * margin;

		const requests = [];
		const allIds = [];

		// Heading.
		const pointsTop = buildHeadingRequests_(requests, pageId, heading, tpl, {
			left: blockLeft,
			top: top,
			width: blockWidth,
		});

		if (tpl.style === "checklist") {
			// Vertical list of rows.
			const rowH = 46;
			const rowGap = 8;
			points.forEach(function (point, i) {
				const pos = {
					x: blockLeft,
					y: pointsTop + i * (rowH + rowGap),
					w: blockWidth,
					h: rowH,
				};
				const ids = buildChecklistRowRequests_(
					requests,
					pageId,
					pos,
					point,
					tpl,
				);
				ids.forEach(function (id) {
					allIds.push(id);
				});
			});
		} else {
			// Numbered cards laid out on a grid (grid-style positioning).
			const cols = suggestTakeawaysCols_(points.length);
			const gap = 16;
			const cardW = (blockWidth - (cols - 1) * gap) / cols;
			const cardH = 90;
			points.forEach(function (point, i) {
				const r = Math.floor(i / cols);
				const c = i % cols;
				const pos = {
					x: blockLeft + c * (cardW + gap),
					y: pointsTop + 12 + r * (cardH + gap),
					w: cardW,
					h: cardH,
				};
				const ids = buildNumberedCardRequests_(
					requests,
					pageId,
					pos,
					i + 1,
					point,
					tpl,
				);
				ids.forEach(function (id) {
					allIds.push(id);
				});
			});
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
		console.error("Error inserting takeaways: " + e.message);
		return { success: false, error: e.message };
	}
}

/**
 * Auto Minter adapter: turns AI-generated takeaway lines (`title | desc`) into
 * the payload insertTakeawaysIntoSlide() accepts. The heading defaults to
 * "Key Takeaways", mirroring the insert fn's own fallback.
 *
 * @param {string} generatedText - raw LLM output from generateTakeawaysFromContext()
 * @param {{templateId?: string, heading?: string}} [hints] - optional router hints
 * @return {{heading: string, points: Array<{title:string,desc:string}>,
 *   templateId: string}|null} null when parsing yields zero usable points
 */
function autoBuildTakeawaysPayload_(generatedText, hints) {
	const h = hints || {};
	const points = parseTakeawayPoints_(generatedText);
	if (!points.length) return null;

	const templates = buildTakeawaysTemplates_();
	let templateId = templates[0].id;
	for (let i = 0; i < templates.length; i++) {
		if (templates[i].id === h.templateId) templateId = templates[i].id;
	}

	const heading =
		(h.heading != null ? String(h.heading).trim() : "") || "Key Takeaways";

	return { heading: heading, points: points, templateId: templateId };
}

// ── Auto Minter registration ─────────────────────────────────────────────
// Self-contained guarded push: GAS file load order is unspecified, so this
// block must not call functions defined in other files at the top level.
// The registry variable MUST be declared `var` + typeof guard (never
// const/let — a const AUTO_MINTERS anywhere would break the whole project).
var AUTO_MINTERS = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
AUTO_MINTERS.push({
	key: "takeaways",
	label: "重點摘要",
	emoji: "🎯",
	order: 100,
	whenToUse:
		"3-5 key conclusions or take-home messages summarizing the content",
	hintsSpec: '{"heading":string}',
	generate: "generateTakeawaysFromContext",
	buildPayload: "autoBuildTakeawaysPayload_",
	insert: "insertTakeawaysIntoSlide",
	previewPartial: "src/components/takeaways-minter/preview",
	previewKind: "list",
	precheck: "",
	options: [
		{
			name: "templateId",
			label: "範本",
			type: "select",
			choicesFrom: "getTakeawaysTemplates",
		},
		{
			name: "heading",
			label: "標題",
			type: "text",
			placeholder: "Key Takeaways",
		},
	],
});
