/**
 * HTML to Slides — server-side builder.
 *
 * Receives the structured JSON produced by the dialog's client-side parser
 * (src/components/html2slides/parser-client.html) and creates slides for it.
 * Contract spec: .claude/skills/gslide-html/reference.md
 *
 * Per-slide failures are collected as warnings, never abort the batch.
 */

/**
 * Entry point called from the dialog via google.script.run.
 * @param {string|Object} payload - parser JSON (string or object)
 * @return {{created: number, total: number, warnings: string[]}}
 */
function convertGslideJsonToSlides(payload) {
	var data = typeof payload === "string" ? JSON.parse(payload) : payload;
	if (!data || !data.slides || data.slides.length === 0) {
		return { created: 0, total: 0, warnings: ["No slides in payload."] };
	}
	var warnings = (data.warnings || []).slice();
	var presentation = SlidesApp.getActivePresentation();
	var insertIndex = getInsertIndex(presentation); // from md2slides/slideCreator.js
	var created = 0;

	// Pass 1 — create slides and fill placeholders (buffered SlidesApp writes).
	// Minters are deferred: several of them draw via the advanced Slides REST
	// service (batchUpdate), which reads the SAVED document and cannot see
	// pages SlidesApp has buffered but not yet flushed ("The page ... could
	// not be found"). See the Apps Script gotcha in CLAUDE.md.
	var minterJobs = []; // {spec, pageObjectId, label}
	for (var i = 0; i < data.slides.length; i++) {
		var spec = data.slides[i];
		var label = "Slide " + (i + 1) + " (" + spec.layout + ")";
		try {
			var res = gslideBuildOneSlide(presentation, insertIndex + created, spec);
			created++;
			for (var w = 0; w < res.warnings.length; w++) {
				warnings.push(label + ": " + res.warnings[w]);
			}
			if (spec.minter && spec.minter.key) {
				minterJobs.push({
					spec: spec,
					pageObjectId: res.pageObjectId,
					slideIndex: insertIndex + created - 1,
					label: label,
				});
			}
		} catch (e) {
			warnings.push(label + ": FAILED — " + e.message);
		}
		for (var pw = 0; pw < (spec.warnings || []).length; pw++) {
			warnings.push(label + ": " + spec.warnings[pw]);
		}
	}

	// Minter jobs are NOT run here: SlidesApp's buffered writes only hit the
	// saved document when this execution ENDS (saveAndClose() mid-execution
	// does not make the new pages visible — verified empirically; see the
	// Apps Script gotcha in CLAUDE.md). The dialog issues a SECOND
	// google.script.run call (runGslideMinterJobs) right after this one
	// returns; by then Google has flushed and every minter can find its page.
	debugLog(
		"html2slides",
		"convertGslideJsonToSlides",
		"created " + created + "/" + data.slides.length + " slides, " +
			minterJobs.length + " minter jobs deferred",
	);
	return {
		created: created,
		total: data.slides.length,
		warnings: warnings,
		minterJobs: minterJobs,
	};
}

/**
 * Second-stage entry point called by the dialog AFTER the slide-creation
 * execution has ended (and its buffered writes were flushed): draws minter
 * content onto the now-saved pages.
 * @param {string|Object} jobsPayload - array of {spec, pageObjectId, slideIndex, label}
 * @return {{warnings: string[]}}
 */
function runGslideMinterJobs(jobsPayload) {
	var jobs = typeof jobsPayload === "string" ? JSON.parse(jobsPayload) : jobsPayload;
	var warnings = [];
	if (!jobs || !jobs.length) return { warnings: warnings };

	var presentation = SlidesApp.getActivePresentation();
	var slides = presentation.getSlides();

	for (var j = 0; j < jobs.length; j++) {
		var job = jobs[j];
		var targetId = job.pageObjectId;
		var found = null;
		try {
			found = presentation.getSlideById(targetId);
		} catch (e) {
			found = null;
		}
		if (!found && slides[job.slideIndex]) {
			// provisional id didn't survive the flush — fall back to position
			// (only meaningful when the position actually holds a DIFFERENT id)
			var freshId = slides[job.slideIndex].getObjectId();
			if (freshId !== targetId) {
				warnings.push(
					job.label + ": id " + targetId + " not found after save; using slide at index " +
						job.slideIndex + " (" + freshId + ").",
				);
				targetId = freshId;
			}
		}
		var minterWarnings = gslideRunMinterOnPage(targetId, job.spec);
		for (var mw = 0; mw < minterWarnings.length; mw++) {
			warnings.push(job.label + ": " + minterWarnings[mw]);
		}
	}
	return { warnings: warnings };
}

/** Creates one slide from a spec and fills its slots. Returns {warnings}. */
function gslideBuildOneSlide(presentation, index, spec) {
	var warnings = [];
	var slide = gslideInsertSlideWithLayout(presentation, index, spec.layout, warnings);

	var placeholders = gslideCollectPlaceholders(slide);
	var slotNames = gslideOrderedSlotNames(spec.slots || {});

	for (var s = 0; s < slotNames.length; s++) {
		var name = slotNames[s];
		var paragraphs = spec.slots[name];
		if (!paragraphs || paragraphs.length === 0) continue;
		var shape = gslideClaimPlaceholder(placeholders, name);
		if (!shape) {
			shape = gslideFallbackTextBox(slide, name);
			warnings.push(
				'no placeholder for slot "' + name + '"; used a plain text box instead.',
			);
		}
		try {
			gslideSetRichText(shape.getText(), paragraphs);
		} catch (e) {
			warnings.push('failed to fill slot "' + name + '": ' + e.message);
		}
	}

	if (spec.notes) {
		try {
			slide.getNotesPage().getSpeakerNotesShape().getText().setText(spec.notes);
		} catch (e) {
			warnings.push("failed to set speaker notes: " + e.message);
		}
	}
	// minter execution is deferred to pass 2 (see convertGslideJsonToSlides)
	return { warnings: warnings, pageObjectId: slide.getObjectId() };
}

/**
 * Insert a slide with the requested predefined layout.
 * Order: PredefinedLayout enum → lzLayoutType() match over the deck's own
 * layouts (locale/import-safe, src/protocol/lz_layouts.js) → TITLE_AND_BODY.
 */
function gslideInsertSlideWithLayout(presentation, index, layoutType, warnings) {
	try {
		if (SlidesApp.PredefinedLayout[layoutType]) {
			return presentation.insertSlide(index, SlidesApp.PredefinedLayout[layoutType]);
		}
	} catch (e) {
		// fall through: master may not carry this predefined layout
	}
	try {
		var layouts = presentation.getLayouts();
		for (var i = 0; i < layouts.length; i++) {
			if (lzLayoutType(layouts[i]) === layoutType) {
				return presentation.insertSlide(index, layouts[i]);
			}
		}
	} catch (e2) {
		// fall through
	}
	warnings.push(
		'layout "' + layoutType + '" not available in this deck; used TITLE_AND_BODY.',
	);
	return presentation.insertSlide(index, SlidesApp.PredefinedLayout.TITLE_AND_BODY);
}

/** Enumerate the slide's placeholder shapes as [{shape, type}]. */
function gslideCollectPlaceholders(slide) {
	var out = [];
	var shapes = slide.getShapes();
	for (var i = 0; i < shapes.length; i++) {
		try {
			var type = shapes[i].getPlaceholderType();
			if (type && type !== SlidesApp.PlaceholderType.NONE) {
				out.push({ shape: shapes[i], type: String(type), used: false });
			}
		} catch (e) {
			// non-placeholder shape
		}
	}
	return out;
}

/** Fill order: TITLE → SUBTITLE → BODY → BODY_2 → rest, so BODY_2 gets the 2nd BODY. */
function gslideOrderedSlotNames(slots) {
	var preferred = ["TITLE", "SUBTITLE", "BODY", "BODY_2"];
	var names = [];
	for (var i = 0; i < preferred.length; i++) {
		if (slots.hasOwnProperty(preferred[i])) names.push(preferred[i]);
	}
	for (var key in slots) {
		if (slots.hasOwnProperty(key) && names.indexOf(key) === -1) names.push(key);
	}
	return names;
}

/** Slot name → acceptable placeholder type names, in preference order. */
var GSLIDE_SLOT_TYPES = {
	TITLE: ["TITLE", "CENTERED_TITLE"],
	SUBTITLE: ["SUBTITLE", "BODY"],
	BODY: ["BODY", "SUBTITLE"],
	BODY_2: ["BODY"],
};

/** Claim the first unused placeholder acceptable for the slot (marks it used). */
function gslideClaimPlaceholder(placeholders, slotName) {
	var accepted = GSLIDE_SLOT_TYPES[slotName] || ["BODY", "SUBTITLE", "TITLE"];
	for (var a = 0; a < accepted.length; a++) {
		for (var i = 0; i < placeholders.length; i++) {
			var ph = placeholders[i];
			if (!ph.used && ph.type === accepted[a]) {
				ph.used = true;
				return ph.shape;
			}
		}
	}
	return null;
}

/** Fallback text box roughly where the slot would sit. */
function gslideFallbackTextBox(slide, slotName) {
	var w = slide.getWidth ? slide.getWidth() : 720;
	var h = slide.getHeight ? slide.getHeight() : 405;
	if (slotName === "TITLE") return slide.insertTextBox("", w * 0.08, h * 0.08, w * 0.84, h * 0.15);
	if (slotName === "BODY_2") return slide.insertTextBox("", w * 0.52, h * 0.3, w * 0.4, h * 0.6);
	return slide.insertTextBox("", w * 0.08, h * 0.3, w * 0.4 + (slotName === "BODY" ? w * 0.44 : 0), h * 0.6);
}

/**
 * Write paragraphs into a text range preserving run styles and list nesting.
 * List nesting uses leading tabs + applyListPreset (Slides treats leading
 * tabs as nesting levels when a preset is applied).
 */
function gslideSetRichText(textRange, paragraphs) {
	var fullText = "";
	var runOffsets = []; // {start, end, run}
	var listRanges = []; // {start, end, type} — contiguous list blocks
	var currentList = null;

	for (var p = 0; p < paragraphs.length; p++) {
		var para = paragraphs[p];
		var prefix = para.listLevel != null && para.listLevel > 0
			? new Array(para.listLevel + 1).join("\t")
			: "";
		if (p > 0) fullText += "\n";
		fullText += prefix;
		var paraStart = fullText.length;
		for (var r = 0; r < para.runs.length; r++) {
			var run = para.runs[r];
			runOffsets.push({
				start: fullText.length,
				end: fullText.length + run.text.length,
				run: run,
			});
			fullText += run.text;
		}
		if (para.listLevel != null) {
			if (currentList && currentList.type === (para.listType || "bullet")) {
				currentList.end = fullText.length;
			} else {
				currentList = {
					start: paraStart - prefix.length,
					end: fullText.length,
					type: para.listType || "bullet",
				};
				listRanges.push(currentList);
			}
		} else {
			currentList = null;
		}
	}

	textRange.setText(fullText);

	for (var i = 0; i < runOffsets.length; i++) {
		var ro = runOffsets[i];
		var run2 = ro.run;
		if (!run2.bold && !run2.italic && !run2.strikethrough && !run2.link) continue;
		try {
			var style = textRange.getRange(ro.start, ro.end).getTextStyle();
			if (run2.bold) style.setBold(true);
			if (run2.italic) style.setItalic(true);
			if (run2.strikethrough) style.setStrikethrough(true);
			if (run2.link) style.setLinkUrl(run2.link);
		} catch (e) {
			Logger.log("gslideSetRichText run style error: " + e.message);
		}
	}

	for (var L = 0; L < listRanges.length; L++) {
		var lr = listRanges[L];
		try {
			var preset =
				lr.type === "number"
					? SlidesApp.ListPreset.DIGIT_ALPHA_ROMAN
					: SlidesApp.ListPreset.DISC_CIRCLE_SQUARE;
			textRange.getRange(lr.start, lr.end).getListStyle().applyListPreset(preset);
		} catch (e) {
			Logger.log("gslideSetRichText list preset error: " + e.message);
		}
	}
}
