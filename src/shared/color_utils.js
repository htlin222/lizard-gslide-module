/**
 * shared/color_utils.js
 *
 * Canonical color helpers shared across the whole module. This is the single
 * source of truth for hex → Slides-API RGB conversion.
 *
 * Before this file there were THREE functionally-identical implementations
 * (hexToRgb in color_palette_api.js and batch/slide_utilities.js, plus
 * hexToRgbColor_ in table_minter.js). They are consolidated here so a fix or
 * tweak happens in one place. `hexToRgbColor_` is kept as a thin back-compat
 * alias for the minters that still call it (compare/takeaways/grid/table).
 *
 * Apps Script note: every function declaration here is hoisted into the global
 * namespace, so any other file can call these directly with no import.
 */

/**
 * Converts a hex color string to a Slides API rgbColor value object.
 * Accepts "#RRGGBB", "RRGGBB", "#RGB" or "RGB" (3-digit shorthand is expanded).
 * Returns black on malformed input rather than throwing.
 *
 * @param {string} hex - hex color, with or without a leading '#'
 * @returns {{red:number, green:number, blue:number}} channels in the 0–1 range
 */
function hexToRgb(hex) {
	let h = String(hex == null ? "" : hex).trim().replace(/^#/, "");

	// Expand 3-digit shorthand (e.g. "abc" -> "aabbcc").
	if (h.length === 3) {
		h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
	}

	if (!/^[0-9a-fA-F]{6}$/.test(h)) {
		return { red: 0, green: 0, blue: 0 };
	}

	return {
		red: parseInt(h.substring(0, 2), 16) / 255,
		green: parseInt(h.substring(2, 4), 16) / 255,
		blue: parseInt(h.substring(4, 6), 16) / 255,
	};
}

/**
 * Back-compat alias for hexToRgb. Retained so minters that still reference
 * hexToRgbColor_ keep working until they migrate to hexToRgb / rgbColor_.
 *
 * @param {string} hex
 * @returns {{red:number, green:number, blue:number}}
 */
function hexToRgbColor_(hex) {
	return hexToRgb(hex);
}

/**
 * Wraps a hex color as the `{ rgbColor: ... }` object Slides API requests
 * expect for solid fills, outlines and foreground colors. Saves repeating
 * `{ rgbColor: hexToRgb(hex) }` at every request site.
 *
 * @param {string} hex
 * @returns {{rgbColor:{red:number, green:number, blue:number}}}
 */
function rgbColor_(hex) {
	return { rgbColor: hexToRgb(hex) };
}
