#!/usr/bin/env node
/**
 * Guard against duplicate global function names across src/**.js.
 *
 * Google Apps Script flattens every .js file into one global namespace, so two
 * top-level `function foo()` declarations silently shadow each other — the last
 * one loaded wins, often with a DIFFERENT signature. That is a load-order-
 * dependent latent bug (see the refactor that introduced this check). This
 * script fails the build if any top-level function name is declared more than
 * once, so collisions are caught in CI instead of at runtime in Slides.
 *
 * Top-level only: matches lines that start with `function name(` at column 0,
 * mirroring how Apps Script global functions are written. Nested/closure
 * functions (indented) are ignored.
 *
 * Node ESM, zero dependencies. Run: `node scripts/check-dup-functions.mjs`.
 * Exit code 0 = clean, 1 = duplicates found.
 */
import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, relative } from "node:path";

const ROOT = process.cwd();
const SRC = join(ROOT, "src");

/** Recursively collect all *.js files under src/. */
function collectJsFiles(dir) {
	const out = [];
	for (const entry of readdirSync(dir)) {
		const full = join(dir, entry);
		const st = statSync(full);
		if (st.isDirectory()) out.push(...collectJsFiles(full));
		else if (entry.endsWith(".js")) out.push(full);
	}
	return out;
}

const TOP_LEVEL_FN = /^function\s+([A-Za-z0-9_$]+)\s*\(/;

// name -> [{ file, line }]
const defs = new Map();

for (const file of collectJsFiles(SRC)) {
	const lines = readFileSync(file, "utf8").split("\n");
	lines.forEach((text, i) => {
		const m = TOP_LEVEL_FN.exec(text);
		if (!m) return;
		const name = m[1];
		if (!defs.has(name)) defs.set(name, []);
		defs.get(name).push({ file: relative(ROOT, file), line: i + 1 });
	});
}

const dups = [...defs.entries()].filter(([, sites]) => sites.length > 1);

if (dups.length === 0) {
	console.log("✓ No duplicate global function names found.");
	process.exit(0);
}

console.error(
	`✗ Found ${dups.length} duplicate global function name(s) — Apps Script's flat\n` +
		`  namespace makes these silently shadow each other (last loaded wins):\n`,
);
for (const [name, sites] of dups) {
	console.error(`  ${name}`);
	for (const s of sites) console.error(`    - ${s.file}:${s.line}`);
}
console.error(
	"\n  Rename all but one definition (use a module-specific name, e.g. fooForSplit_),\n" +
		"  or remove exact duplicates. See scripts/check-dup-functions.mjs header.",
);
process.exit(1);
