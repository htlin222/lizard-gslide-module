#!/usr/bin/env node
/**
 * Builds the self-update bundle consumed by cloned decks (src/util/self_update.js).
 *
 * Output:
 *   dist/bundle.json  → { version, files: [{ name, type, source }] }
 *   dist/version.json → { version }
 *
 * `files` mirrors what clasp pushes (all src/**.js + src/**.html + appsscript.json),
 * in the Apps Script projects.updateContent shape: `name` is the repo path WITHOUT
 * extension (slashes preserved), `type` is SERVER_JS / HTML / JSON.
 *
 * Also stamps src/util/version.js with the git short SHA so a clone's
 * SCRIPT_VERSION reflects the version it runs.
 *
 * Node ESM, zero dependencies. Run: `node scripts/build-bundle.mjs`.
 */
import {
	readFileSync,
	writeFileSync,
	readdirSync,
	statSync,
	mkdirSync,
} from "node:fs";
import { join, relative, extname } from "node:path";
import { execSync } from "node:child_process";

const ROOT = process.cwd();
const SRC = join(ROOT, "src");
const DIST = join(ROOT, "dist");

function gitShortSha() {
	try {
		return execSync("git rev-parse --short HEAD", { cwd: ROOT })
			.toString()
			.trim();
	} catch {
		return "dev";
	}
}

const version = gitShortSha();

// 1) Stamp src/util/version.js BEFORE reading files, so the bundle carries it.
const versionFile = join(SRC, "util", "version.js");
writeFileSync(
	versionFile,
	"/**\n" +
		" * Installed script version — STAMPED at build time by scripts/build-bundle.mjs.\n" +
		' * Placeholder "dev" means built without stamping (always reports update available).\n' +
		" */\n" +
		'var SCRIPT_VERSION = "' +
		version +
		'";\n',
);

// 2) Walk src/ for .js/.gs/.html (this is exactly what clasp pushes from src/).
function walk(dir, out) {
	for (const entry of readdirSync(dir)) {
		const p = join(dir, entry);
		if (statSync(p).isDirectory()) walk(p, out);
		else out.push(p);
	}
	return out;
}

const files = [];
for (const abs of walk(SRC, [])) {
	const ext = extname(abs);
	let type;
	if (ext === ".js" || ext === ".gs") type = "SERVER_JS";
	else if (ext === ".html") type = "HTML";
	else continue;
	const rel = relative(ROOT, abs).split("\\").join("/");
	const name = rel.replace(/\.(js|gs|html)$/, "");
	files.push({ name, type, source: readFileSync(abs, "utf8") });
}

// 3) The manifest (mandatory; must be name "appsscript", type "JSON").
files.push({
	name: "appsscript",
	type: "JSON",
	source: readFileSync(join(ROOT, "appsscript.json"), "utf8"),
});

// 4) Emit dist/.
mkdirSync(DIST, { recursive: true });
writeFileSync(join(DIST, "bundle.json"), JSON.stringify({ version, files }));
writeFileSync(
	join(DIST, "version.json"),
	JSON.stringify({ version }, null, 2) + "\n",
);

const js = files.filter((f) => f.type === "SERVER_JS").length;
const html = files.filter((f) => f.type === "HTML").length;
console.log(
	`Built dist/ @ ${version}: ${files.length} files (${js} js, ${html} html, 1 manifest)`,
);
