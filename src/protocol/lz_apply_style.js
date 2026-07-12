// 🎨 LZ-PROTOCOL — apply cosmetics online from each element's injected instruction
/**
 * A python-pptx deck built with the LZ style kit injects a self-describing TOML
 * style instruction into every element's alt-text (see LZ-PROTOCOL.md). After
 * importing that .pptx into Google Slides, run `lzApplyAll()` (or the menu item)
 * to CATCH every tagged element and apply the instruction ONLINE — font, size,
 * bold/italic, color, geometry, fill, and table styling.
 *
 * The instruction travels WITH each element, so lizard is a dumb interpreter and
 * needs no synced copy of the spec. python's spec_source.py is the sole authority.
 *
 * Division of labor:
 *   local pptx  → structure + content + injected instructions (rough layout ok)
 *   lizard cloud → all cosmetics (this file) + chrome rebuild (ultra batch)
 */

function _lzAnchor(name) {
	var map = {
		MIDDLE: SlidesApp.ContentAlignment.MIDDLE,
		TOP: SlidesApp.ContentAlignment.TOP,
		BOTTOM: SlidesApp.ContentAlignment.BOTTOM,
	};
	return map[name] || null;
}

/** Apply a parsed instruction's text + geometry + fill to a Shape. */
function lzApplyToShape(shape, instr) {
	// geometry (points) — reposition/resize the "roughly placed" pptx shape
	try {
		if (typeof instr.x === "number") shape.setLeft(instr.x);
		if (typeof instr.y === "number") shape.setTop(instr.y);
		if (typeof instr.w === "number") shape.setWidth(instr.w);
		if (typeof instr.h === "number") shape.setHeight(instr.h);
	} catch (e) {
		/* some element types reject resize — ignore */
	}

	// fill
	if (instr.fill) {
		try {
			shape.getFill().setSolidFill(instr.fill);
		} catch (e) {
			/* ignore */
		}
	}

	// vertical anchor
	var anchor = _lzAnchor(instr.anchor);
	if (anchor) {
		try {
			shape.setContentAlignment(anchor);
		} catch (e) {
			/* ignore */
		}
	}

	// text style
	var ts = null;
	try {
		ts = shape.getText && shape.getText().getTextStyle();
	} catch (e) {
		ts = null;
	}
	if (ts) {
		try {
			if (instr.font) ts.setFontFamily(instr.font);
			if (typeof instr.size === "number") ts.setFontSize(instr.size);
			if (typeof instr.bold === "boolean") ts.setBold(instr.bold);
			if (typeof instr.italic === "boolean") ts.setItalic(instr.italic);
			if (instr.color) ts.setForegroundColor(instr.color);
		} catch (e) {
			/* ignore */
		}
	}
}

/** Apply a TABLE instruction: geometry + per-row (header/cell) styling. */
function lzApplyToTable(table, instr) {
	try {
		if (typeof instr.x === "number") table.setLeft(instr.x);
		if (typeof instr.y === "number") table.setTop(instr.y);
		if (typeof instr.w === "number") table.setWidth(instr.w);
	} catch (e) {
		/* ignore */
	}
	var rows = table.getNumRows();
	var cols = table.getNumColumns();
	for (var r = 0; r < rows; r++) {
		var isHead = r === 0;
		var color = isHead ? instr.header_color : instr.cell_color;
		var bold = isHead ? instr.header_bold : instr.cell_bold;
		var size = isHead ? instr.header_size : instr.cell_size;
		for (var c = 0; c < cols; c++) {
			var ts;
			try {
				ts = table.getCell(r, c).getText().getTextStyle();
			} catch (e) {
				continue;
			}
			try {
				if (instr.font) ts.setFontFamily(instr.font);
				if (typeof size === "number") ts.setFontSize(size);
				if (typeof bold === "boolean") ts.setBold(bold);
				if (color) ts.setForegroundColor(color);
			} catch (e) {
				/* ignore cell */
			}
		}
	}
}

/**
 * Walk the deck, catch every LZ-tagged element, and apply its injected
 * instruction. Horizontal paragraph alignment is flushed in one Advanced-Slides
 * batch (SlidesApp has no direct setter). Returns the count of styled elements.
 */
function lzApplyStyleAll() {
	var presentation = SlidesApp.getActivePresentation();
	var slides = presentation.getSlides();
	var touched = 0;
	var alignReqs = [];

	for (var s = 0; s < slides.length; s++) {
		var slide = slides[s];

		// shapes
		var shapes = slide.getShapes();
		for (var i = 0; i < shapes.length; i++) {
			var shape = shapes[i];
			var role = lzRoleOf(shape);
			if (!role || lzIsManaged(shape)) continue; // chrome is rebuilt, not styled
			var instr = lzInstr(shape);
			if (!instr) continue;
			lzApplyToShape(shape, instr);
			if (instr.align) {
				alignReqs.push({
					updateParagraphStyle: {
						objectId: shape.getObjectId(),
						textRange: { type: "ALL" },
						style: { alignment: instr.align },
						fields: "alignment",
					},
				});
			}
			touched++;
		}

		// tables (not returned by getShapes)
		var tables = slide.getTables();
		for (var t = 0; t < tables.length; t++) {
			var table = tables[t];
			if (lzRoleOf(table) !== LZ_ROLES.TABLE) continue;
			var tinstr = lzInstr(table);
			if (!tinstr) continue;
			lzApplyToTable(table, tinstr);
			touched++;
		}
	}

	if (alignReqs.length) {
		try {
			Slides.Presentations.batchUpdate(
				{ requests: alignReqs },
				presentation.getId(),
			);
		} catch (e) {
			/* alignment best-effort */
		}
	}
	return touched;
}

/**
 * One command: apply injected cosmetics to foreign content, THEN rebuild all
 * chrome (menu / progress / page numbers / section boxes) from live slide order.
 */
function lzApplyAll() {
	var n = lzApplyStyleAll();
	if (typeof runAllFunctionsUltraMegaBatch === "function") {
		runAllFunctionsUltraMegaBatch();
	}
	return n;
}
