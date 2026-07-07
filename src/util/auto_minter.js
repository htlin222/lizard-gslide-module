/**
 * Server-side orchestrator for the ⚡ Auto Minter dialog.
 *
 * Two-round flow, both rounds reusing the individual minters' existing code:
 *   Round 1 — autoMinterRoute(context): one Groq call picks 2-3 candidate
 *     minters (with a short reason + config hints each) from the registry.
 *   Round 2 — autoMinterGenerate(key, context, hints): runs the chosen
 *     minter's own generate<X>FromContext fn, then its autoBuild<X>Payload_
 *     adapter, returning the exact payload its insert<X>IntoSlide accepts.
 *     The dialog previews straight from that payload (insert-what-you-see).
 *
 * Registry: each minter file self-registers by pushing a descriptor into the
 * global AUTO_MINTERS array (guarded `var` declaration, safe under GAS's
 * unspecified file load order — see the registration block in any minter).
 * All descriptor function fields are NAME STRINGS resolved lazily via
 * globalThis here, so nothing depends on load order. Adding a new minter
 * requires zero changes to this file or the auto-minter dialog.
 *
 * Descriptor schema:
 *   {
 *     key: string,            // unique id, e.g. "kpi"
 *     label: string,          // zh-TW display name
 *     emoji: string,
 *     order: number,          // registry sort order
 *     whenToUse: string,      // one English line for the router prompt
 *     hintsSpec: string,      // JSON-ish description of accepted hints, or ""
 *     generate: string,       // fn name: (context, hints) →
 *                             //   {success, generatedText?, needKey?, error?}
 *     buildPayload: string,   // fn name: (generatedText, hints) → payload|null
 *     insert: string,         // fn name: (payload) → {success, error?}
 *     previewPartial: string, // template path for the dialog preview, or ""
 *     previewKind: string,    // generic fallback renderer key
 *     precheck: string,       // fn name: (context) → bool eligibility, or ""
 *     options: Array<Object>, // declarative user-facing options (see below)
 *   }
 *
 * Option spec (rendered generically by the dialog; values merge into hints):
 *   {
 *     name: string,           // hint key the adapter reads, e.g. "rows"
 *     label: string,          // zh-TW label shown in the dialog
 *     type: "number" | "select" | "checkbox" | "text",
 *     choices?: Array<{value, label}>,   // static choices (select)
 *     choicesFrom?: string,   // fn name returning templates [{id, name}] —
 *                             // resolved server-side into `choices`
 *     default?: *,            // initial value (empty/absent = auto)
 *     min?: number, max?: number, placeholder?: string,
 *     regenerate?: boolean,   // true = affects the LLM text, so changing it
 *                             // needs a re-generate (not just a re-build)
 *   }
 */

/**
 * Reads the self-registered minter registry, sorted by `order`.
 * Declared lazily so evaluation order of project files never matters.
 * @return {Array<Object>}
 */
function getAutoMinterRegistry_() {
	const reg = typeof AUTO_MINTERS === "undefined" ? [] : AUTO_MINTERS;
	return reg.slice().sort(function (a, b) {
		return (a.order || 99) - (b.order || 99);
	});
}

/**
 * Resolves one option spec for the client: a `choicesFrom` fn name becomes a
 * concrete `choices` array ([{value: t.id, label: t.name}]).
 * @param {Object} opt
 * @return {Object} client-safe option spec
 */
function resolveAutoOption_(opt) {
	const out = {
		name: opt.name,
		label: opt.label,
		type: opt.type,
		default: opt.default,
		min: opt.min,
		max: opt.max,
		placeholder: opt.placeholder,
		regenerate: !!opt.regenerate,
		choices: opt.choices || null,
	};
	if (!out.choices && opt.choicesFrom) {
		const fn = resolveAutoFn_(opt.choicesFrom);
		if (fn) {
			try {
				out.choices = (fn() || []).map(function (t) {
					return { value: t.id, label: t.name || t.id };
				});
			} catch (e) {
				out.choices = null;
			}
		}
	}
	return out;
}

/**
 * Registry view for the dialog preload: display metadata + resolved option
 * specs (no fn names leak to the client). When a minter has a template-backed
 * select (choicesFrom), the full template objects ride along as `templates`
 * so the client preview can render the SELECTED template's colors/style
 * instead of a one-look approximation.
 * @return {Array<{key, label, emoji, options: Array<Object>, templates: Array}>}
 */
function getAutoMinterPublicList_() {
	return getAutoMinterRegistry_().map(function (d) {
		const options = d.options || [];
		let templates = null;
		for (let i = 0; i < options.length; i++) {
			if (!options[i].choicesFrom) continue;
			const fn = resolveAutoFn_(options[i].choicesFrom);
			if (fn) {
				try {
					templates = fn() || null;
				} catch (e) {
					templates = null;
				}
			}
			break;
		}
		return {
			key: d.key,
			label: d.label,
			emoji: d.emoji,
			options: options.map(resolveAutoOption_),
			templates: templates,
		};
	});
}

/**
 * Looks up a registry descriptor by key.
 * @param {string} key
 * @return {Object|null}
 */
function findAutoMinter_(key) {
	const reg = getAutoMinterRegistry_();
	for (let i = 0; i < reg.length; i++) {
		if (reg[i].key === key) return reg[i];
	}
	return null;
}

/**
 * Resolves a descriptor's function-name string to the actual global function.
 * @param {string} name
 * @return {Function|null}
 */
function resolveAutoFn_(name) {
	if (!name) return null;
	const fn = globalThis[name];
	return typeof fn === "function" ? fn : null;
}

/**
 * Builds the round-1 router system prompt from the eligible registry entries.
 * One "key: whenToUse" line per minter plus shared slide-design rules of
 * thumb, ending with a strict JSON-only output contract.
 * @param {Array<Object>} entries
 * @return {string}
 */
function buildAutoRouterPrompt_(entries) {
	const lines = [
		"You are a slide-layout router. Given the user's content, pick the",
		"2-3 BEST visual layouts for a single presentation slide.",
		"Available layouts:",
	];
	for (let i = 0; i < entries.length; i++) {
		const d = entries[i];
		let line = "- " + d.key + ": " + d.whenToUse;
		if (d.hintsSpec) line += " Hints: " + d.hintsSpec;
		lines.push(line);
	}
	lines.push(
		"Slide design rules of thumb:",
		"- At most 5 KPI cards; 3-5 takeaways; a table for dense row-column data.",
		"- Sequences: timeline when dated, steps when undated.",
		"- compare for exactly 2-3 alternatives; grid for 3-9 parallel concepts.",
		"- callout for one single message; gallery ONLY when image URLs are present.",
		"- icon only as a decorative last resort.",
		'Reply with ONLY this JSON, no prose, no code fences:',
		'{"candidates":[{"key":"...","reason":"...","hints":{}}]}',
		"- 2 or 3 candidates, ranked best-first.",
		"- key MUST be one of the layout keys listed above.",
		"- reason: at most 15 characters of Traditional Chinese (繁體中文).",
		"- hints: an object following that layout's Hints spec, or {}.",
	);
	return lines.join("\n");
}

/**
 * Robustly extracts a JSON object from LLM output: strips code fences, then
 * parses the substring from the first "{" to the last "}".
 * @param {string} text
 * @return {Object|null}
 */
function extractAutoJson_(text) {
	const raw = String(text == null ? "" : text)
		.replace(/```(?:json)?/gi, "")
		.trim();
	const start = raw.indexOf("{");
	const end = raw.lastIndexOf("}");
	if (start === -1 || end === -1 || end <= start) return null;
	try {
		const parsed = JSON.parse(raw.slice(start, end + 1));
		return parsed && typeof parsed === "object" ? parsed : null;
	} catch (e) {
		return null;
	}
}

/**
 * Round 1: routes the pasted context to 2-3 candidate minters via Groq.
 * Called from the dialog through google.script.run.
 *
 * @param {string} context - Arbitrary text the user pasted.
 * @return {{success: boolean,
 *           candidates?: Array<{key, label, emoji, reason, hints}>,
 *           needKey?: boolean, error?: string}}
 */
function autoMinterRoute(context) {
	const text = (context || "").trim();
	if (!text) {
		return { success: false, error: "No context provided." };
	}
	if (!hasUserApiKey()) {
		return {
			success: false,
			needKey: true,
			error:
				"No AI key set. Run 🖖 跨頁功能 → 🔑 設定 AI 金鑰 (Groq) first, then try again.",
		};
	}

	// Eligibility pre-filter (e.g. gallery only when image URLs are present).
	const eligible = getAutoMinterRegistry_().filter(function (d) {
		const check = resolveAutoFn_(d.precheck);
		if (!check) return true;
		try {
			return !!check(text);
		} catch (e) {
			return true;
		}
	});
	if (!eligible.length) {
		return { success: false, error: "No eligible minters registered." };
	}

	const res = callGroq_(buildAutoRouterPrompt_(eligible), text, {
		maxTokens: 500,
		temperature: 0.2,
		responseFormat: { type: "json_object" },
	});
	if (!res.success) return res;

	const parsed = extractAutoJson_(res.generatedText);
	const rawCandidates =
		parsed && Object.prototype.toString.call(parsed.candidates) === "[object Array]"
			? parsed.candidates
			: null;
	if (!rawCandidates) {
		return {
			success: false,
			error: "AI 回傳格式錯誤，請再按一次「生成」。",
		};
	}

	// Keep only known eligible keys, dedupe, cap at 3, decorate from registry.
	const byKey = {};
	for (let i = 0; i < eligible.length; i++) byKey[eligible[i].key] = eligible[i];
	const seen = {};
	const candidates = [];
	for (let i = 0; i < rawCandidates.length && candidates.length < 3; i++) {
		const c = rawCandidates[i] || {};
		const d = byKey[c.key];
		if (!d || seen[c.key]) continue;
		seen[c.key] = true;
		candidates.push({
			key: d.key,
			label: d.label,
			emoji: d.emoji,
			reason: String(c.reason || "").slice(0, 30),
			hints: c.hints && typeof c.hints === "object" ? c.hints : {},
		});
	}
	if (!candidates.length) {
		return {
			success: false,
			error: "AI 無法判斷合適的鑄造器，請補充內容或改用手動鑄造器。",
		};
	}
	return { success: true, candidates: candidates };
}

/**
 * Round 2: converts the context into the chosen minter's insert payload by
 * reusing that minter's own generate fn + payload adapter.
 * Called from the dialog through google.script.run.
 *
 * @param {string} key - Registry key of the chosen minter.
 * @param {string} context - The same pasted text as round 1.
 * @param {Object} [hints] - Config hints from the router candidate.
 * @return {{success: boolean, key?: string, payload?: Object,
 *           previewKind?: string, needKey?: boolean, error?: string}}
 */
function autoMinterGenerate(key, context, hints) {
	const d = findAutoMinter_(key);
	if (!d) {
		return { success: false, error: "Unknown minter: " + key };
	}
	const gen = resolveAutoFn_(d.generate);
	const build = resolveAutoFn_(d.buildPayload);
	if (!gen || !build) {
		return {
			success: false,
			error: "Minter '" + key + "' is misregistered (missing " +
				(gen ? d.buildPayload : d.generate) + ").",
		};
	}

	const res = gen((context || "").trim(), hints || {});
	if (!res || !res.success) {
		// Pass needKey / error through untouched for the dialog to render.
		return res || { success: false, error: "Generation failed." };
	}

	let payload = null;
	try {
		payload = build(res.generatedText, hints || {});
	} catch (e) {
		console.error("Auto Minter adapter error (" + key + "): " + e.message);
	}
	if (!payload) {
		return {
			success: false,
			error:
				"AI 輸出無法解析成「" + d.label + "」，請重新生成或切換候選。",
		};
	}
	return {
		success: true,
		key: key,
		payload: payload,
		// generatedText lets the dialog re-build the payload with different
		// options (autoMinterRebuild) without another LLM call.
		generatedText: res.generatedText,
		previewKind: d.previewKind || "",
	};
}

/**
 * Re-builds the insert payload from an already generated text with new hints
 * (e.g. the user changed rows/cols/template post hoc). No LLM call.
 * Called from the dialog through google.script.run.
 *
 * @param {string} key
 * @param {string} generatedText - Text returned by autoMinterGenerate.
 * @param {Object} [hints] - Updated option values.
 * @return {{success: boolean, key?: string, payload?: Object,
 *           previewKind?: string, error?: string}}
 */
function autoMinterRebuild(key, generatedText, hints) {
	const d = findAutoMinter_(key);
	if (!d) {
		return { success: false, error: "Unknown minter: " + key };
	}
	const build = resolveAutoFn_(d.buildPayload);
	if (!build) {
		return {
			success: false,
			error: "Minter '" + key + "' has no adapter (" + d.buildPayload + ").",
		};
	}
	let payload = null;
	try {
		payload = build(generatedText || "", hints || {});
	} catch (e) {
		console.error("Auto Minter rebuild error (" + key + "): " + e.message);
	}
	if (!payload) {
		return {
			success: false,
			error: "選項套用失敗，請調整後再試或重新生成。",
		};
	}
	return {
		success: true,
		key: key,
		payload: payload,
		previewKind: d.previewKind || "",
	};
}

/**
 * Inserts a previously generated payload via the minter's own insert fn.
 * Called from the dialog through google.script.run.
 *
 * @param {string} key
 * @param {Object} payload - The payload returned by autoMinterGenerate.
 * @return {{success: boolean, error?: string}}
 */
function autoMinterInsert(key, payload) {
	const d = findAutoMinter_(key);
	if (!d) {
		return { success: false, error: "Unknown minter: " + key };
	}
	const insert = resolveAutoFn_(d.insert);
	if (!insert) {
		return {
			success: false,
			error: "Minter '" + key + "' has no insert fn (" + d.insert + ").",
		};
	}
	if (!payload) {
		return { success: false, error: "Nothing to insert — generate first." };
	}
	try {
		return insert(payload) || { success: true };
	} catch (e) {
		console.error("Auto Minter insert error (" + key + "): " + e.message);
		return { success: false, error: e.message };
	}
}
