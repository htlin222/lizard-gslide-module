/**
 * shared/shape_requests.js
 *
 * Declarative Slides API request builders. These return plain request objects
 * (no side effects) to be pushed onto a `requests` array and sent in a single
 * batchUpdate — the "minter-standard" pattern exemplified by grid_minter.js.
 *
 * Extracted from grid_minter.js's buildUnitCardRequests_ so every minter can
 * build cards/text the same way instead of re-inlining the same nested request
 * literals. Depends on rgbColor_() from shared/color_utils.js.
 */

/**
 * Builds a Slides API textRange object.
 * @param {{start:number, end:number}|"ALL"} range - char range, or the string "ALL"
 * @returns {Object} textRange
 */
function textRange_(range) {
	if (range === "ALL") return { type: "ALL" };
	return { type: "FIXED_RANGE", startIndex: range.start, endIndex: range.end };
}

/**
 * createShape request for a box (default TEXT_BOX) at an exact PT rect.
 * @param {{pageId:string, id:string, x:number, y:number, w:number, h:number, shapeType?:string}} o
 * @returns {Object} createShape request
 */
function createTextBoxRequest_(o) {
	return {
		createShape: {
			objectId: o.id,
			shapeType: o.shapeType || "TEXT_BOX",
			elementProperties: {
				pageObjectId: o.pageId,
				size: {
					width: { magnitude: o.w, unit: "PT" },
					height: { magnitude: o.h, unit: "PT" },
				},
				transform: {
					scaleX: 1,
					scaleY: 1,
					translateX: o.x,
					translateY: o.y,
					unit: "PT",
				},
			},
		},
	};
}

/**
 * updateShapeProperties request for solid fill + solid outline.
 * @param {string} id - shape objectId
 * @param {{fillColor:string, borderColor:string, borderWidth?:number, dashStyle?:string, contentAlignment?:string}} o
 * @returns {Object} updateShapeProperties request
 */
function fillOutlineRequest_(id, o) {
	return {
		updateShapeProperties: {
			objectId: id,
			shapeProperties: {
				shapeBackgroundFill: {
					solidFill: { color: rgbColor_(o.fillColor) },
				},
				outline: {
					outlineFill: {
						solidFill: { color: rgbColor_(o.borderColor) },
					},
					weight: { magnitude: o.borderWidth || 1, unit: "PT" },
					dashStyle: o.dashStyle || "SOLID",
				},
				contentAlignment: o.contentAlignment || "TOP",
			},
			fields:
				"shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color,outline.weight,outline.dashStyle,contentAlignment",
		},
	};
}

/**
 * insertText request.
 * @param {string} id - shape objectId
 * @param {string} text
 * @param {number} [index=0] - insertion index
 * @returns {Object} insertText request
 */
function insertTextRequest_(id, text, index) {
	return {
		insertText: {
			objectId: id,
			insertionIndex: index || 0,
			text: text,
		},
	};
}

/**
 * updateTextStyle request for a range.
 * @param {string} id - shape objectId
 * @param {{start:number, end:number}|"ALL"} range
 * @param {{color:string, bold?:boolean, italic?:boolean, fontFamily?:string, fontSize?:number}} s
 * @returns {Object} updateTextStyle request
 */
function textStyleRequest_(id, range, s) {
	return {
		updateTextStyle: {
			objectId: id,
			textRange: textRange_(range),
			style: {
				foregroundColor: { opaqueColor: rgbColor_(s.color) },
				bold: !!s.bold,
				italic: !!s.italic,
				fontFamily: s.fontFamily,
				fontSize: { magnitude: s.fontSize, unit: "PT" },
			},
			fields: "foregroundColor,bold,italic,fontFamily,fontSize",
		},
	};
}

/**
 * updateParagraphStyle request. Caller supplies the style object and matching
 * fields mask (e.g. style {lineSpacing:150}, fields "lineSpacing").
 * @param {string} id - shape objectId
 * @param {{start:number, end:number}|"ALL"} range
 * @param {Object} style - paragraphStyle fragment
 * @param {string} fields - fields mask matching `style`
 * @returns {Object} updateParagraphStyle request
 */
function paragraphStyleRequest_(id, range, style, fields) {
	return {
		updateParagraphStyle: {
			objectId: id,
			textRange: textRange_(range),
			style: style,
			fields: fields,
		},
	};
}
