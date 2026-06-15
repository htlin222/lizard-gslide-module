/**
 * BYOK (Bring Your Own Key) management for AI features.
 *
 * SECURITY MODEL:
 * - Each user stores THEIR OWN key in PropertiesService.getUserProperties().
 *   This store is scoped to the individual Google account, lives server-side in
 *   Google's infrastructure, and is invisible to other users of the add-on.
 * - The key is NEVER sent back to the client, never stored in localStorage, and
 *   never passed as an argument across google.script.run. All provider calls
 *   happen here, server-side, via UrlFetchApp.
 *
 * We use Groq (https://console.groq.com/docs/overview): fast, free tier, and an
 * OpenAI-compatible Chat Completions API.
 */

const AI_KEY_PROPERTY = "GROQ_API_KEY";
const GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions";
const GROQ_DEFAULT_MODEL = "llama-3.3-70b-versatile";

/**
 * Saves the current user's Groq API key.
 * Called from the sidebar via google.script.run. Returns only a boolean —
 * never echoes the key back.
 *
 * @param {string} apiKey - The user's Groq API key (starts with "gsk_").
 * @return {boolean} true on success.
 */
function saveUserApiKey(apiKey) {
	const key = (apiKey || "").trim();
	if (!/^gsk_[A-Za-z0-9]+$/.test(key)) {
		throw new Error(
			"Invalid Groq key. Get one at https://console.groq.com/keys (starts with 'gsk_').",
		);
	}
	PropertiesService.getUserProperties().setProperty(AI_KEY_PROPERTY, key);
	return true;
}

/**
 * Reports whether the current user has a key saved — WITHOUT returning the key.
 * Use this to drive sidebar UI state (e.g. "key set ✓" vs. "enter key").
 *
 * @return {boolean}
 */
function hasUserApiKey() {
	return !!PropertiesService.getUserProperties().getProperty(AI_KEY_PROPERTY);
}

/**
 * Deletes the current user's stored key.
 * @return {boolean} true on success.
 */
function clearUserApiKey() {
	PropertiesService.getUserProperties().deleteProperty(AI_KEY_PROPERTY);
	return true;
}

/**
 * Reusable gate for ANY AI-powered function across ANY module.
 *
 * Usage — put this as the first line of any function that needs AI:
 *
 *   function aiSummarizeSlide() {
 *     if (!ensureApiKey_("aiSummarizeSlide")) return;  // pops the dialog if no key
 *     const result = callGroq_(systemMsg, userMsg);
 *     ...
 *   }
 *
 * If a key exists, returns true and the caller proceeds normally.
 * If not, it opens the key-entry dialog and returns false. After the user
 * saves their key, the dialog re-runs `continueFnName` automatically, so the
 * original action completes without the user having to click again.
 *
 * @param {string} continueFnName - Name of the global function to re-run after
 *   the key is saved. Must be a no-argument global function.
 * @return {boolean} true if a key is already set (proceed); false otherwise.
 */
function ensureApiKey_(continueFnName) {
	if (hasUserApiKey()) return true;
	showApiKeyDialog(continueFnName);
	return false;
}

/**
 * Menu-safe entry point for explicit "set up my AI key" actions.
 * (Menu callbacks receive an event object, so we don't point a menu item
 * directly at showApiKeyDialog — that arg would be treated as continueFnName.)
 */
function showAiKeySetup() {
	showApiKeyDialog();
}

/**
 * Opens the reusable Groq key-entry modal dialog.
 * Can also be wired to a menu item (e.g. "🔑 設定 AI 金鑰") for explicit setup.
 *
 * @param {string} [continueFnName] - Optional function to re-run after saving.
 */
function showApiKeyDialog(continueFnName) {
	const template = HtmlService.createTemplateFromFile(
		"src/components/ai-key-dialog",
	);
	template.continueFnName = continueFnName || "";
	const html = template.evaluate().setWidth(440).setHeight(380);
	SlidesApp.getUi().showModalDialog(html, "🔑 Set up AI");
}

/**
 * Internal: read the current user's key or throw a friendly error.
 * Not exposed to the client.
 *
 * @return {string}
 */
function getUserApiKeyOrThrow_() {
	const key = PropertiesService.getUserProperties().getProperty(
		AI_KEY_PROPERTY,
	);
	if (!key) {
		throw new Error(
			"No API key set. Open the 🔑 API Key panel and paste your Groq key " +
				"(get a free one at https://console.groq.com/keys).",
		);
	}
	return key;
}

/**
 * Calls Groq's chat completions endpoint with the current user's stored key.
 * The key is read server-side here — the client never supplies it.
 *
 * @param {string} systemMessage - System role content.
 * @param {string} userMessage - User role content.
 * @param {Object} [opts] - { model, maxTokens, temperature }.
 * @return {{success: boolean, generatedText?: string, model?: string, usage?: Object, error?: string}}
 */
function callGroq_(systemMessage, userMessage, opts) {
	try {
		const apiKey = getUserApiKeyOrThrow_();
		const options = opts || {};

		const body = {
			model: options.model || GROQ_DEFAULT_MODEL,
			messages: [
				{ role: "system", content: systemMessage },
				{ role: "user", content: userMessage },
			],
			max_tokens: options.maxTokens || 1000,
			temperature: options.temperature != null ? options.temperature : 0.7,
		};

		const response = UrlFetchApp.fetch(GROQ_API_URL, {
			method: "POST",
			contentType: "application/json",
			headers: { Authorization: `Bearer ${apiKey}` },
			payload: JSON.stringify(body),
			muteHttpExceptions: true, // read error bodies without dumping the key into a stack trace
		});

		const code = response.getResponseCode();
		const data = JSON.parse(response.getContentText());

		if (code !== 200) {
			// Log status only — NEVER the request headers/payload (they contain the key).
			console.error("Groq API error: " + code);
			const msg = data && data.error ? data.error.message : "Unknown error";
			return { success: false, error: `Groq API error (${code}): ${msg}` };
		}

		if (!data.choices || data.choices.length === 0) {
			return { success: false, error: "No response generated from Groq API" };
		}

		return {
			success: true,
			generatedText: data.choices[0].message.content.trim(),
			usage: data.usage || {},
			model: data.model || body.model,
		};
	} catch (e) {
		console.error("Error calling Groq API: " + e.message);
		return { success: false, error: e.message, generatedText: "" };
	}
}
