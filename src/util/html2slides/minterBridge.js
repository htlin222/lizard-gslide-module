/**
 * HTML to Slides — minter bridge.
 *
 * Maps the generic content bag a <section data-minter="..."> parses into
 * (items / columns / table / images / paragraphs / options, see
 * src/components/html2slides/parser-client.html) onto each minter's
 * insert<X>IntoSlide payload, and invokes it against a specific slide via
 * payload.pageObjectId (resolveMinterTargetSlide_, src/util/minter_target.js).
 *
 * Insert function names are resolved from the AUTO_MINTERS registry (each
 * minter self-registers; see src/util/auto_minter.js descriptor schema), so
 * this file never hardcodes them.
 */

/**
 * Per-key payload builders: (bag, titleText) → payload object, or a string
 * error when required content is missing.
 */
var GSLIDE_MINTER_PAYLOADS = {
	timeline: function (bag) {
		if (!bag.items.length) return "needs a <ul> of <li data-date=...> items";
		return {
			items: bag.items.map(function (it) {
				return { date: it.date || "", label: it.label || "" };
			}),
			templateId: bag.options.templateId,
			orientation: bag.options.orientation,
		};
	},
	steps: function (bag) {
		if (!bag.items.length) return "needs a <ul> of steps";
		return {
			steps: bag.items.map(function (it) {
				return { title: it.label || "", desc: it.desc || "" };
			}),
			templateId: bag.options.templateId,
			orientation: bag.options.orientation,
		};
	},
	takeaways: function (bag, titleText) {
		if (!bag.items.length) return "needs a <ul> of takeaway points";
		return {
			heading: bag.options.heading || titleText || "",
			points: bag.items.map(function (it) {
				return { title: it.label || "", desc: it.desc || "" };
			}),
			templateId: bag.options.templateId,
		};
	},
	compare: function (bag) {
		if (bag.columns.length < 2) {
			return 'needs ≥2 elements with data-col="Column title"';
		}
		return { columns: bag.columns, templateId: bag.options.templateId };
	},
	callout: function (bag, titleText) {
		var body = bag.paragraphs.join("\n");
		if (!body) return "needs a <p> body";
		return {
			templateId: bag.options.templateId,
			header: bag.options.header || titleText || "",
			body: body,
		};
	},
	gallery: function (bag) {
		if (!bag.images.length) return "needs <img src=...> elements";
		return {
			items: bag.images,
			cols: bag.options.cols,
			templateId: bag.options.templateId,
			captions: bag.options.captions,
		};
	},
	barchart: function (bag) {
		if (!bag.items.length) return "needs a <ul> of <li data-value=...> items";
		return {
			data: bag.items.map(function (it) {
				return { label: it.label || "", value: it.value };
			}),
			templateId: bag.options.templateId,
			orientation: bag.options.orientation,
			showValues: bag.options.showValues,
		};
	},
	kpi: function (bag) {
		if (!bag.items.length) return "needs a <ul> of <li data-value=...> items";
		return {
			items: bag.items.map(function (it) {
				return { value: it.value || "", label: it.label || "", trend: it.trend || "" };
			}),
			templateId: bag.options.templateId,
		};
	},
	agenda: function (bag) {
		if (!bag.items.length) return "needs a <ul> of agenda items";
		return {
			items: bag.items.map(function (it) {
				return it.label || "";
			}),
			templateId: bag.options.templateId,
		};
	},
	table: function (bag) {
		if (!bag.table || !bag.table.header.length) return "needs a <table>";
		return {
			header: bag.table.header,
			body: bag.table.body,
			theme: bag.options.theme,
			fontSize: bag.options.fontSize,
			widthPx: bag.options.widthPx,
			left: bag.options.left,
			top: bag.options.top,
			title: bag.options.title || "",
		};
	},
	grid: function (bag) {
		if (!bag.items.length) return "needs a <ul> of card items";
		return {
			units: bag.items.map(function (it) {
				return {
					title: it.label || "",
					subtitle: it.subtitle || "",
					body: it.desc || it.body || "",
				};
			}),
			rows: bag.options.rows,
			cols: bag.options.cols,
			styleNumber: bag.options.styleNumber,
		};
	},
	icon: function (bag) {
		if (!bag.options.glyph) return 'needs data-minter-options {"glyph":"★"}';
		return {
			glyph: bag.options.glyph,
			size: bag.options.size,
			color: bag.options.color,
		};
	},
};

/**
 * Runs a parsed minter spec against an already-saved slide.
 * Called in the builder's pass 2, AFTER Presentation.saveAndClose() — the
 * REST-based minters (batchUpdate) cannot see unsaved SlidesApp pages.
 * @param {string} pageObjectId - target slide's object id
 * @param {Object} spec - slide spec with .minter bag and .slots
 * @return {string[]} warnings (empty on clean success)
 */
function gslideRunMinterOnPage(pageObjectId, spec) {
	var bag = spec.minter;
	var key = bag.key;
	var build = GSLIDE_MINTER_PAYLOADS[key];
	if (!build) return ['unknown minter "' + key + '" (bridge has no adapter).'];

	var titleText = "";
	if (spec.slots && spec.slots.TITLE && spec.slots.TITLE.length) {
		titleText = spec.slots.TITLE.map(gslideParagraphPlainText_).join(" ");
	}

	var payload = build(bag, titleText);
	if (typeof payload === "string") {
		return ['minter "' + key + '" skipped: ' + payload + "."];
	}
	payload.pageObjectId = pageObjectId;

	var insertFn = gslideResolveMinterInsert_(key);
	if (!insertFn) {
		return ['minter "' + key + '" not found in AUTO_MINTERS registry.'];
	}
	try {
		var result = insertFn(payload) || {};
		if (!result.success) {
			return ['minter "' + key + '" failed: ' + (result.error || "unknown error")];
		}
	} catch (e) {
		return ['minter "' + key + '" threw: ' + e.message];
	}
	return [];
}

/** Insert fn for a minter key, via the AUTO_MINTERS registry (lazy, load-order safe). */
function gslideResolveMinterInsert_(key) {
	var reg = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
	for (var i = 0; i < reg.length; i++) {
		if (reg[i] && reg[i].key === key) {
			var fn = globalThis[reg[i].insert];
			return typeof fn === "function" ? fn : null;
		}
	}
	return null;
}

/** Plain text of a parser paragraph ({runs:[{text,...}]}). */
function gslideParagraphPlainText_(para) {
	var s = "";
	var runs = (para && para.runs) || [];
	for (var i = 0; i < runs.length; i++) s += runs[i].text;
	return s;
}
