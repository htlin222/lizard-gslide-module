/**
 * Lightweight smoke / unit tests for the minter server modules.
 *
 * This is a plain-node test (no framework — the project has no npm deps) run via
 * `node test/minters.smoke.test.js`. It lives under test/ which .claspignore
 * excludes, so it never ships to Apps Script.
 *
 * Each minter file is loaded into a fresh VM sandbox with the GAS globals
 * stubbed. Top-level `function` declarations attach to the sandbox, so we can:
 *   1. confirm the file evaluates without throwing (stronger than `node --check`),
 *   2. assert each expected server function exists,
 *   3. call the pure `get*Templates()` builders and check they return data,
 *   4. assert real behavior for the grid parser/layout helpers.
 */

const vm = require("vm");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
let passed = 0;
let failed = 0;

function ok(name, cond) {
	if (cond) {
		passed++;
	} else {
		failed++;
		console.error("  ✗ " + name);
	}
}

/**
 * Loads a util file into a fresh sandbox with stubbed GAS globals.
 * @param {string} rel - path relative to repo root
 * @return {Object} the sandbox (top-level functions are properties on it)
 */
function load(rel) {
	const sandbox = {
		// config.js palette globals
		main_color: "#3D6869",
		accent_color: "#f29424",
		base_color: "#FFFFFF",
		text_color: "#333333",
		sub1_color: "#E7EAE7",
		sub2_color: "#E7F9F5",
		main_font_family: "Source Sans Pro",
		label_font_size: 14,
		// shared helpers used by some template builders
		hexToRgbColor_: () => ({ red: 0, green: 0, blue: 0 }),
		getStyleDefinitions: () => ({}),
		getConfigValues: () => ({}),
		// GAS services — stubbed (insert* functions touch these but we don't call them)
		SlidesApp: {},
		Slides: {},
		Utilities: { getUuid: () => "uuid-stub" },
		console,
	};
	vm.createContext(sandbox);
	vm.runInContext(fs.readFileSync(path.join(ROOT, rel), "utf8"), sandbox);
	return sandbox;
}

// ── Each minter: file loads, expected functions exist, templates non-empty ──
const cases = [
	{ file: "src/util/callout_minter.js", fns: ["getCalloutTemplates", "insertCalloutIntoSlide"], templates: "getCalloutTemplates" },
	{ file: "src/util/kpi_minter.js", fns: ["getKpiTemplates", "insertKpiIntoSlide"], templates: "getKpiTemplates" },
	{ file: "src/util/timeline_minter.js", fns: ["getTimelineTemplates", "insertTimelineIntoSlide"], templates: "getTimelineTemplates" },
	{ file: "src/util/compare_minter.js", fns: ["getCompareTemplates", "insertCompareIntoSlide"], templates: "getCompareTemplates" },
	{ file: "src/util/steps_minter.js", fns: ["getStepsTemplates", "insertStepsIntoSlide"], templates: "getStepsTemplates" },
	{ file: "src/util/gallery_minter.js", fns: ["getGalleryTemplates", "insertGalleryIntoSlide"], templates: "getGalleryTemplates" },
	{ file: "src/util/agenda_minter.js", fns: ["getAgendaItems", "getAgendaTemplates", "insertAgendaIntoSlide"], templates: "getAgendaTemplates" },
	{ file: "src/util/takeaways_minter.js", fns: ["getTakeawaysTemplates", "insertTakeawaysIntoSlide"], templates: "getTakeawaysTemplates" },
	{ file: "src/util/icon_minter.js", fns: ["insertIconIntoSlide"], templates: null },
	{ file: "src/util/barchart_minter.js", fns: ["getBarChartTemplates", "insertBarChartIntoSlide"], templates: "getBarChartTemplates" },
	{ file: "src/util/grid_minter.js", fns: ["parseGridUnits_", "suggestGridLayout_", "insertGridIntoSlide"], templates: null },
];

for (const c of cases) {
	let sb;
	try {
		sb = load(c.file);
	} catch (e) {
		failed++;
		console.error("  ✗ load " + c.file + " → " + e.message);
		continue;
	}
	for (const fn of c.fns) {
		ok(c.file + " defines " + fn + "()", typeof sb[fn] === "function");
	}
	if (c.templates && typeof sb[c.templates] === "function") {
		let arr;
		try {
			arr = sb[c.templates]();
		} catch (e) {
			arr = null;
			console.error("  ✗ " + c.templates + "() threw: " + e.message);
		}
		ok(c.templates + "() returns a non-empty array", Array.isArray(arr) && arr.length > 0);
	}
}

// ── Real assertions for the grid parser/layout (deterministic, pure) ──
const grid = load("src/util/grid_minter.js");

const units = grid.parseGridUnits_(
	"# Speed\n## Fast\n\nShips quickly.\n---\n# Quality\n## Solid\n\nFewer bugs.",
);
ok("parseGridUnits_ finds 2 units", units.length === 2);
ok("parseGridUnits_ reads title", units[0] && units[0].title === "Speed");
ok("parseGridUnits_ reads subtitle", units[0] && units[0].subtitle === "Fast");
ok("parseGridUnits_ reads body", units[0] && units[0].body === "Ships quickly.");

ok("suggestGridLayout_(6) → 2x3", JSON.stringify(grid.suggestGridLayout_(6)) === JSON.stringify({ rows: 2, cols: 3 }));
ok("suggestGridLayout_(4) → 2x2", JSON.stringify(grid.suggestGridLayout_(4)) === JSON.stringify({ rows: 2, cols: 2 }));
ok("suggestGridLayout_(1) → 1x1", JSON.stringify(grid.suggestGridLayout_(1)) === JSON.stringify({ rows: 1, cols: 1 }));

// ── Callout templates shape ──
const callout = load("src/util/callout_minter.js");
const ct = callout.getCalloutTemplates();
ok("getCalloutTemplates has 7 templates", ct.length === 7);
ok("callout templates carry an id + barColor", ct.every((t) => t.id && t.barColor));

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
