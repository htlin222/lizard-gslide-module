// 🗺️ LZ-PROTOCOL — canonical layout-type resolver (locale + import-source robust)
/**
 * `getLayoutName()` is unreliable as an identity: Google localizes the layout
 * DISPLAY name to the editor's UI language (zh-Hant "區段標題"), and a PPTX import
 * carries whatever the source named it (python-pptx: "Section Header"; a Google
 * export: "SECTION_HEADER"). So matching one hardcoded string misses layouts.
 *
 * LZ_LAYOUT_ALIASES maps each canonical Google predefined layout type to every
 * display name we've seen for it. `lzLayoutType(slide|layout)` returns the
 * canonical type regardless of locale or import source.
 *
 * Names are grouped by source confidence:
 *   - Google canonical enum + PowerPoint/python-pptx template names → verified.
 *   - Localized names → best-effort; run `lzDumpLayoutNames()` to read the EXACT
 *     names your environment returns and add any that are missing.
 */
var LZ_LAYOUT_ALIASES = {
	TITLE: ["TITLE", "Title Slide", "Title slide", "標題投影片", "标题幻灯片"],
	SECTION_HEADER: [
		"SECTION_HEADER",
		"Section Header",
		"Section header",
		"區段標題", // zh-Hant (verified)
		"区段标题", // zh-Hans
		"セクションの見出し", // ja
		"섹션 헤더", // ko
	],
	TITLE_AND_BODY: [
		"TITLE_AND_BODY",
		"Title and Body",
		"Title and Content", // PowerPoint / python-pptx
		"標題和內文",
		"标题和正文",
	],
	TITLE_AND_TWO_COLUMNS: [
		"TITLE_AND_TWO_COLUMNS",
		"Title and Two Columns",
		"Two Content", // PowerPoint
		"Comparison", // PowerPoint
		"標題和兩欄",
		"标题和两栏",
	],
	TITLE_ONLY: ["TITLE_ONLY", "Title Only", "Title only", "只有標題", "仅标题"],
	ONE_COLUMN_TEXT: ["ONE_COLUMN_TEXT", "One Column Text", "單欄文字", "单栏文本"],
	MAIN_POINT: ["MAIN_POINT", "Main Point", "重點", "要点"],
	SECTION_TITLE_AND_DESCRIPTION: [
		"SECTION_TITLE_AND_DESCRIPTION",
		"Section Title and Description",
		"區段標題和說明",
		"区段标题和说明",
	],
	CAPTION_ONLY: [
		"CAPTION_ONLY",
		"Caption Only",
		"Content with Caption", // PowerPoint
		"Picture with Caption", // PowerPoint
		"只有說明",
	],
	BIG_NUMBER: ["BIG_NUMBER", "Big Number", "大型數字", "大数字"],
	BLANK: ["BLANK", "Blank", "空白"],
};

// reverse index: display name -> canonical type (built once)
var _LZ_LAYOUT_BY_NAME = (function () {
	var m = {};
	for (var type in LZ_LAYOUT_ALIASES) {
		if (!LZ_LAYOUT_ALIASES.hasOwnProperty(type)) continue;
		var names = LZ_LAYOUT_ALIASES[type];
		for (var i = 0; i < names.length; i++) m[names[i]] = type;
	}
	return m;
})();

/** Raw display name from a Slide or Layout, or "". */
function lzLayoutName(slideOrLayout) {
	if (!slideOrLayout) return "";
	try {
		var layout = slideOrLayout.getLayout
			? slideOrLayout.getLayout()
			: slideOrLayout;
		return layout && layout.getLayoutName ? layout.getLayoutName() || "" : "";
	} catch (e) {
		return "";
	}
}

/**
 * Canonical layout type for a Slide or Layout, resolved across locale / import
 * source. Returns "" if the name isn't recognized (extend LZ_LAYOUT_ALIASES).
 */
function lzLayoutType(slideOrLayout) {
	var name = lzLayoutName(slideOrLayout);
	if (!name) return "";
	if (_LZ_LAYOUT_BY_NAME[name]) return _LZ_LAYOUT_BY_NAME[name];
	// tolerate case / spacing / separator drift (SECTION_HEADER vs "Section Header")
	var norm = name.toUpperCase().replace(/[\s_-]+/g, "_");
	return _LZ_LAYOUT_BY_NAME[norm] || (LZ_LAYOUT_ALIASES[norm] ? norm : "");
}

/** True if the slide's layout resolves to `type` (e.g. "SECTION_HEADER"). */
function lzIsLayout(slideOrLayout, type) {
	return lzLayoutType(slideOrLayout) === type;
}

/**
 * Diagnostic — logs each distinct layout display name in the active deck and the
 * type it resolves to (or ??? if unrecognized). Run this in your environment to
 * discover the exact localized names, then add any "???" ones to LZ_LAYOUT_ALIASES.
 */
function lzDumpLayoutNames() {
	var slides = SlidesApp.getActivePresentation().getSlides();
	var seen = {};
	for (var i = 0; i < slides.length; i++) {
		var name = lzLayoutName(slides[i]);
		if (name && !seen[name]) seen[name] = lzLayoutType(slides[i]) || "???";
	}
	var lines = [];
	for (var n in seen) if (seen.hasOwnProperty(n)) lines.push(n + "  →  " + seen[n]);
	var report = lines.length
		? lines.join("\n")
		: "(no layout names found)";
	Logger.log("LZ layout names:\n" + report);
	try {
		SlidesApp.getUi().alert("🗺️ 版面名稱對照", report, SlidesApp.getUi().ButtonSet.OK);
	} catch (e) {
		/* no UI (e.g. run from editor) — Logger has it */
	}
	return seen;
}
