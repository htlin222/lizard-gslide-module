/**
 * shared/theme_colors.js
 *
 * Single resolver for the deck's theme palette. Replaces the copy-pasted
 * `const main = (typeof main_color !== "undefined" && main_color) || "#3D6869"`
 * fallback blocks that several minters (kpi/steps/compare/takeaways/…) each
 * re-declared independently.
 *
 * Resolution order per color: saved user config (PropertiesService, via
 * getConfigValues) → global config variable (config.js) → hard default. The
 * hard defaults mirror config.js so a minter still renders sensibly even if the
 * globals failed to load.
 */

/** Hard fallbacks — kept in sync with the var declarations in config.js. */
var THEME_COLOR_DEFAULTS_ = {
	main: "#3D6869",
	base: "#FFFFFF",
	text: "#333333",
	accent: "#f29424",
	sub1: "#E7EAE7",
	sub2: "#E7F9F5",
};

/**
 * Reads a global var by name without throwing if it is undefined.
 * @param {string} name
 * @returns {string|undefined}
 */
function readGlobalColor_(name) {
	try {
		// `this` is the global object in Apps Script's V8 runtime.
		const v = this[name];
		return typeof v === "string" && v ? v : undefined;
	} catch (e) {
		return undefined;
	}
}

/**
 * Returns the resolved theme palette.
 * @returns {{main:string, base:string, text:string, accent:string, sub1:string, sub2:string}}
 */
function getThemeColors() {
	let config = {};
	try {
		if (typeof getConfigValues === "function") config = getConfigValues() || {};
	} catch (e) {
		config = {};
	}

	const d = THEME_COLOR_DEFAULTS_;
	return {
		main: config.mainColor || readGlobalColor_("main_color") || d.main,
		base: config.baseColor || readGlobalColor_("base_color") || d.base,
		text: config.textColor || readGlobalColor_("text_color") || d.text,
		accent: config.accentColor || readGlobalColor_("accent_color") || d.accent,
		sub1: config.sub1Color || readGlobalColor_("sub1_color") || d.sub1,
		sub2: config.sub2Color || readGlobalColor_("sub2_color") || d.sub2,
	};
}
