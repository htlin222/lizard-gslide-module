/**
 * Server-side AI helper for the Table Minter dialog.
 *
 * Turns arbitrary user "context" (notes, data, a paragraph, CSV, …) into a
 * clean Markdown table using Groq. The key never leaves the server: this calls
 * callGroq_() in aiKey.js, which reads the user's stored key server-side.
 *
 * The client (src/components/table-minter) is responsible for robustly
 * extracting just the table block from the returned text — LLMs sometimes wrap
 * it in prose like "Here is your markdown:" or code fences.
 */

const TABLE_MINTER_SYSTEM_PROMPT = [
	"You convert the user's content into a titled Markdown table.",
	"Output EXACTLY this, in order, and nothing else:",
	"1. One title line that starts with '# Table: ' followed by a short descriptive title.",
	"2. Then ONE GitHub-Flavored Markdown table (header row, separator row, data rows).",
	"Strict rules:",
	"- No explanation, no preamble, no closing remarks, no code fences.",
	"- The first line MUST start with '# Table: '.",
	"- The first character of the table MUST be a pipe '|'.",
	"- Put labels/categories in the first column when the data has that shape.",
	"- Keep cell text concise; do not invent data that isn't implied by the input.",
	"- Use a single space of padding inside cells.",
].join("\n");

/**
 * Generates a Markdown table from free-form context via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean, generatedText?: string, error?: string, needKey?: boolean}}
 */
function generateTableMarkdownFromContext(context) {
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

	return callGroq_(TABLE_MINTER_SYSTEM_PROMPT, text, {
		maxTokens: 2000,
		temperature: 0.2,
	});
}

/**
 * Converts "#RRGGBB" to an API rgbColor object (components in 0..1).
 * @param {string} hex
 * @return {{red: number, green: number, blue: number}}
 */
function hexToRgbColor_(hex) {
	const h = (hex || "#000000").replace("#", "");
	return {
		red: parseInt(h.substring(0, 2), 16) / 255,
		green: parseInt(h.substring(2, 4), 16) / 255,
		blue: parseInt(h.substring(4, 6), 16) / 255,
	};
}

/**
 * Builds one updateTableBorderProperties request.
 * @param {string} tableId
 * @param {Object|null} range - TableRange, or null for the whole table.
 * @param {string} position - ALL | INNER_HORIZONTAL | TOP | BOTTOM | ...
 * @param {string} hex - border color
 * @param {number} weight - PT
 * @param {number} alpha - 0..1 (0 hides the border)
 * @return {Object}
 */
function tableBorderRequest_(tableId, range, position, hex, weight, alpha) {
	const req = {
		updateTableBorderProperties: {
			objectId: tableId,
			borderPosition: position,
			tableBorderProperties: {
				tableBorderFill: {
					solidFill: { color: { rgbColor: hexToRgbColor_(hex) }, alpha: alpha },
				},
				weight: { magnitude: weight, unit: "PT" },
				dashStyle: "SOLID",
			},
			fields: "tableBorderFill,weight,dashStyle",
		},
	};
	if (range) req.updateTableBorderProperties.tableRange = range;
	return req;
}

/**
 * Inserts a styled table onto the current slide at an exact position via the
 * Slides API — the alternative to clipboard paste when position matters.
 *
 * Mirrors the clipboard look: transparent fill, left-aligned text, header in
 * the accent color with thick accent strokes top & bottom, thin gray lines
 * between/under body rows, all body text black.
 *
 * @param {{header: string[], body: string[][], theme: string, fontSize: number,
 *   widthPx: number, left: number, top: number, title: string}} payload
 * @return {{success: boolean, error?: string}}
 */
function insertTableIntoSlide(payload) {
	try {
		const p = payload || {};
		const header = p.header || [];
		const body = p.body || [];
		if (!header.length) {
			return { success: false, error: "No table to insert." };
		}

		const theme = p.theme || main_color;
		const fontSize = p.fontSize > 0 ? p.fontSize : 14;
		const leftPt = p.left != null ? p.left : 25;
		const topPt = p.top != null ? p.top : 100;
		// Pasted px widths are read by Slides at 96dpi; convert to points to match.
		const widthPt = p.widthPx > 0 ? p.widthPx * 0.75 : 636;
		const rowHeightPt = fontSize + 14;

		const numRows = body.length + 1;
		const numCols = header.length;
		const totalHeight = numRows * rowHeightPt;

		// Resolve the target slide (fall back to the first slide).
		const presentation = SlidesApp.getActivePresentation();
		let slide = null;
		try {
			slide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		// Optional title text box above the table; the table is pushed down by
		// the title's height + a small gap so they don't overlap.
		const titleText = (p.title || "").trim();
		let tableTop = topPt;
		if (titleText) {
			const titleFontSize = fontSize;
			const titleHeight = Math.round(titleFontSize * 1.6) + 4;
			const gap = 8;
			const box = slide.insertTextBox(
				titleText,
				leftPt,
				topPt,
				widthPt,
				titleHeight,
			);
			try {
				box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
				const tt = box.getText();
				tt.getTextStyle()
					.setFontFamily(main_font_family)
					.setFontSize(titleFontSize)
					.setBold(true)
					.setForegroundColor("#000000");
				tt.getParagraphStyle().setParagraphAlignment(
					SlidesApp.ParagraphAlignment.START,
				);
			} catch (titleErr) {
				// Styling failure shouldn't block the table insert.
			}
			tableTop = topPt + titleHeight + gap;
		}

		// Build EVERYTHING — table, text, borders, column widths — into one
		// batchUpdate so it applies as a single instant change instead of
		// styling visibly step by step. (Created objects are available to later
		// requests within the same batch, so createTable can come first.)
		const pageId = slide.getObjectId();
		const tableId = "tblminter" + Utilities.getUuid().replace(/-/g, "");
		const headerRange = {
			location: { rowIndex: 0, columnIndex: 0 },
			rowSpan: 1,
			columnSpan: numCols,
		};
		const gray = "#CCCCCC";
		const rows = [header].concat(body);

		const requests = [
			{
				createTable: {
					objectId: tableId,
					elementProperties: {
						pageObjectId: pageId,
						size: {
							width: { magnitude: widthPt, unit: "PT" },
							height: { magnitude: totalHeight, unit: "PT" },
						},
						transform: {
							scaleX: 1,
							scaleY: 1,
							translateX: leftPt,
							translateY: tableTop,
							unit: "PT",
						},
					},
					rows: numRows,
					columns: numCols,
				},
			},
			{
				updateTableCellProperties: {
					objectId: tableId,
					tableRange: {
						location: { rowIndex: 0, columnIndex: 0 },
						rowSpan: numRows,
						columnSpan: numCols,
					},
					tableCellProperties: { contentAlignment: "MIDDLE" },
					fields: "contentAlignment",
				},
			},
		];

		// Cell text + per-cell styling.
		for (let r = 0; r < numRows; r++) {
			const isHeader = r === 0;
			for (let c = 0; c < numCols; c++) {
				const value = (rows[r] && rows[r][c]) || "";
				if (!value) continue;
				const cellLocation = { rowIndex: r, columnIndex: c };
				requests.push({
					insertText: {
						objectId: tableId,
						cellLocation: cellLocation,
						insertionIndex: 0,
						text: value,
					},
				});
				requests.push({
					updateTextStyle: {
						objectId: tableId,
						cellLocation: cellLocation,
						textRange: { type: "ALL" },
						style: {
							foregroundColor: {
								opaqueColor: {
									rgbColor: hexToRgbColor_(isHeader ? theme : "#000000"),
								},
							},
							bold: isHeader,
							fontFamily: main_font_family,
							fontSize: { magnitude: fontSize, unit: "PT" },
						},
						fields: "foregroundColor,bold,fontFamily,fontSize",
					},
				});
				requests.push({
					updateParagraphStyle: {
						objectId: tableId,
						cellLocation: cellLocation,
						textRange: { type: "ALL" },
						style: { alignment: "START" },
						fields: "alignment",
					},
				});
			}
		}

		// Borders: hide all, then thin gray row lines everywhere, and ONLY the
		// top of the header gets the thick accent stroke. The header's bottom
		// (2nd line) stays thin gray like every other row separator.
		requests.push(tableBorderRequest_(tableId, null, "ALL", "#FFFFFF", 1, 0));
		requests.push(
			tableBorderRequest_(tableId, null, "INNER_HORIZONTAL", gray, 0.75, 1),
		);
		requests.push(tableBorderRequest_(tableId, null, "BOTTOM", gray, 0.75, 1));
		requests.push(tableBorderRequest_(tableId, headerRange, "TOP", theme, 1.5, 1));

		// Proportional column widths summing to the requested total width.
		const weights = [];
		for (let c = 0; c < numCols; c++) {
			let max = 1;
			for (let r = 0; r < numRows; r++) {
				const len = ((rows[r] && rows[r][c]) || "").length;
				if (len > max) max = len;
			}
			weights.push(max);
		}
		const weightSum = weights.reduce((a, b) => a + b, 0);
		for (let c = 0; c < numCols; c++) {
			requests.push({
				updateTableColumnProperties: {
					objectId: tableId,
					columnIndices: [c],
					tableColumnProperties: {
						columnWidth: {
							magnitude: Math.max(40, (widthPt * weights[c]) / weightSum),
							unit: "PT",
						},
					},
					fields: "columnWidth",
				},
			});
		}

		try {
			Slides.Presentations.batchUpdate({ requests }, presentation.getId());
		} catch (batchErr) {
			return { success: false, error: batchErr.message };
		}

		return { success: true };
	} catch (e) {
		console.error("Error inserting table: " + e.message);
		return { success: false, error: e.message };
	}
}
