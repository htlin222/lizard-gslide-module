// 🔗 LZ-PROTOCOL — shared object protocol for PPTX ⇄ Google Slides
/**
 * The PPTX import boundary discards objectId + OOXML name, but preserves
 * alt-text (title / description). LZ-Protocol formalizes alt-text as a
 * namespaced identity channel so python-pptx and lizard address the same
 * objects by ROLE. See LZ-PROTOCOL.md.
 *
 * Google Apps Script has no ES modules — these are global functions/consts.
 */

// Role vocabulary (alt-text `title` values). Kept back-compatible with the
// magic strings lizard already used (PROGRESS / PROGRESS_BG / MAIN_TITLE).
var LZ_ROLES = {
	// chrome (delete-rebuild)
	PROGRESS: "PROGRESS",
	PROGRESS_BG: "PROGRESS_BG",
	TAB: "TAB",
	PAGE_NUM: "PAGE_NUM",
	SECTION_BOX: "SECTION_BOX",
	SECTION_LABEL: "SECTION_LABEL",
	OUTLINE: "OUTLINE",
	MAIN_TITLE: "MAIN_TITLE",
	// marker (authored, detection only, never deleted)
	SECTION: "SECTION",
	// content (catch & re-apply style; never deleted)
	TITLE: "TITLE",
	TITLE_MAIN: "TITLE_MAIN",
	SUBTITLE: "SUBTITLE",
	DATE: "DATE",
	EMAIL: "EMAIL",
	BRAND_CHIP: "BRAND_CHIP",
	BODY: "BODY",
	TABLE: "TABLE",
	COL_LEFT: "COL_LEFT",
	COL_RIGHT: "COL_RIGHT",
	COL_HEAD_L: "COL_HEAD_L",
	COL_HEAD_R: "COL_HEAD_R",
	KEY_HEADLINE: "KEY_HEADLINE",
	KEY_SUB: "KEY_SUB",
	KEY_POINTS: "KEY_POINTS",
	CITATION: "CITATION",
};

// Chrome lizard owns (delete-and-rebuild). Excludes SECTION (authored content).
var LZ_MANAGED_ROLES = {
	PROGRESS: true,
	PROGRESS_BG: true,
	TAB: true,
	PAGE_NUM: true,
	SECTION_BOX: true,
	SECTION_LABEL: true,
	OUTLINE: true,
	MAIN_TITLE: true,
};

/**
 * Resolve an element's LZ role.
 * Priority: alt-text title (exact) → JSON payload in description → "".
 * @param {GoogleAppsScript.Slides.PageElement} el
 * @return {string} a LZ_ROLES value, or "" if untagged.
 */
function lzRoleOf(el) {
	if (!el || !el.getTitle) return "";
	var t = "";
	try {
		t = el.getTitle() || "";
	} catch (e) {
		t = "";
	}
	if (t && LZ_ROLES[t]) return t;

	// Fallback: {"lz":1,"role":"progress",...} in the description.
	var d = "";
	try {
		d = (el.getDescription && el.getDescription()) || "";
	} catch (e) {
		d = "";
	}
	if (d && d.charAt(0) === "{") {
		try {
			var obj = JSON.parse(d);
			if (obj && obj.lz && obj.role) {
				var r = String(obj.role).toUpperCase();
				if (LZ_ROLES[r]) return r;
			}
		} catch (e) {
			/* not JSON — ignore */
		}
	}
	return "";
}

/** True if the element is protocol-tagged chrome that lizard should reclaim. */
function lzIsManaged(el) {
	return !!LZ_MANAGED_ROLES[lzRoleOf(el)];
}

/** True if the element marks its slide as a section boundary. */
function lzIsSectionMarker(el) {
	return lzRoleOf(el) === LZ_ROLES.SECTION;
}

/**
 * Clean section title from a SECTION marker's instruction (`title = "…"`),
 * or "" if none. Lets the marker carry a tidy title independent of the shape's
 * visible text.
 */
function lzMarkerTitle(el) {
	var instr = lzInstr(el);
	return instr && instr.title ? String(instr.title).trim() : "";
}

/**
 * Parse the TOML-ish style instruction python injects into alt-text.
 * Controlled format: `[lz]` header, then `key = value` lines where value is a
 * "quoted string", a number, or true/false. Returns a flat object, or null.
 */
function lzParseInstr(text) {
	if (!text) return null;
	var out = {};
	var lines = String(text).split("\n");
	var found = false;
	for (var i = 0; i < lines.length; i++) {
		var line = lines[i].trim();
		if (!line || line.charAt(0) === "#" || line.charAt(0) === "[") continue;
		var eq = line.indexOf("=");
		if (eq < 0) continue;
		var key = line.slice(0, eq).trim();
		var raw = line.slice(eq + 1).trim();
		var val;
		if (raw.charAt(0) === '"') {
			val = raw.replace(/^"|"$/g, "");
		} else if (raw === "true" || raw === "false") {
			val = raw === "true";
		} else {
			var n = parseFloat(raw);
			val = isNaN(n) ? raw : n;
		}
		out[key] = val;
		found = true;
	}
	return found ? out : null;
}

/** The parsed style instruction carried in an element's alt-text description. */
function lzInstr(el) {
	if (!el || !el.getDescription) return null;
	try {
		return lzParseInstr(el.getDescription());
	} catch (e) {
		return null;
	}
}

/** Stamp an element with a role (lizard → pptx direction, and internal use). */
function lzTag(el, role, data) {
	if (!el || !el.setTitle) return el;
	try {
		el.setTitle(role);
		if (data && el.setDescription) {
			var payload = { lz: 1, role: String(role).toLowerCase(), v: 1 };
			for (var k in data) if (data.hasOwnProperty(k)) payload[k] = data[k];
			el.setDescription(JSON.stringify(payload));
		}
	} catch (e) {
		/* best-effort */
	}
	return el;
}
