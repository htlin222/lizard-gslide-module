/**
 * Self-update for cloned decks — "🔄 更新腳本 (Fetch latest)".
 *
 * Each clone pulls the latest source from the canonical GitHub repo and
 * overwrites its OWN Apps Script content via the Apps Script API
 * (projects.updateContent). No central registry, no clasp, no local machine.
 *
 * Flow: a GitHub Action builds dist/bundle.json ({version, files:[{name,type,
 * source}]}) on every push to main. This script fetches it and PUTs it to
 * https://script.googleapis.com/v1/projects/{scriptId}/content authorized with
 * ScriptApp.getOAuthToken(). The new code applies on the NEXT reload (the
 * running execution keeps the old code).
 *
 * Requirements (one-time, per Google account):
 *   - Enable "Google Apps Script API" at script.google.com/home/usersettings.
 *   - Approve the OAuth consent (adds the script.projects scope on first update).
 *
 * Reads SCRIPT_VERSION from src/util/version.js (stamped at build time).
 */

/** Base URL of the published bundle on GitHub raw. */
var SELF_UPDATE_RAW_BASE =
	"https://raw.githubusercontent.com/htlin222/lizard-gslide-module/main/dist";

/** @return {string} the installed version (git short SHA, or "dev"). */
function getInstalledVersion_() {
	return (typeof SCRIPT_VERSION !== "undefined" && SCRIPT_VERSION) || "dev";
}

/**
 * Checks GitHub for a newer version (fetches the tiny dist/version.json).
 * @return {{ok: boolean, installed: string, latest?: string,
 *   updateAvailable?: boolean, error?: string}}
 */
function checkForUpdate() {
	var installed = getInstalledVersion_();
	try {
		var res = UrlFetchApp.fetch(SELF_UPDATE_RAW_BASE + "/version.json", {
			muteHttpExceptions: true,
		});
		if (res.getResponseCode() !== 200) {
			return {
				ok: false,
				installed: installed,
				error: "version.json HTTP " + res.getResponseCode(),
			};
		}
		var latest = JSON.parse(res.getContentText()).version;
		return {
			ok: true,
			installed: installed,
			latest: latest,
			updateAvailable: !!latest && latest !== installed,
		};
	} catch (e) {
		return { ok: false, installed: installed, error: e.message };
	}
}

/** Menu handler: report whether an update is available. */
function menuCheckForUpdate() {
	var ui = SlidesApp.getUi();
	var r = checkForUpdate();
	if (!r.ok) {
		ui.alert("無法檢查更新：" + r.error);
		return;
	}
	if (r.updateAvailable) {
		ui.alert(
			"有新版本",
			"目前版本：" +
				r.installed +
				"\n最新版本：" +
				r.latest +
				"\n\n用「🔄 更新腳本」即可更新。",
			ui.ButtonSet.OK,
		);
	} else {
		ui.alert("已是最新版本（" + r.installed + "）");
	}
}

/**
 * Validates a downloaded bundle before writing it (abort BEFORE the PUT so the
 * clone is never left half-written).
 * @param {{version?: string, files?: Array}} bundle
 * @param {string} [expectedVersion] - version.json's version, for a cross-check
 * @return {{ok: boolean, error?: string}}
 */
function validateBundle_(bundle, expectedVersion) {
	if (!bundle || !bundle.files || !bundle.files.length) {
		return { ok: false, error: "files 為空" };
	}
	var manifests = 0;
	for (var i = 0; i < bundle.files.length; i++) {
		var f = bundle.files[i];
		if (!f || !f.name || !f.type || typeof f.source !== "string") {
			return { ok: false, error: "檔案項目不完整" };
		}
		if (f.name === "appsscript" && f.type === "JSON") manifests++;
	}
	if (manifests !== 1) {
		return { ok: false, error: "manifest 數量異常 (" + manifests + ")" };
	}
	if (expectedVersion && bundle.version && bundle.version !== expectedVersion) {
		return {
			ok: false,
			error:
				"版本不一致 (bundle " + bundle.version + " vs " + expectedVersion + ")",
		};
	}
	return { ok: true };
}

/**
 * Menu handler: pull the latest bundle and overwrite this project's content.
 */
function fetchLatestScript() {
	var ui = SlidesApp.getUi();

	var chk = checkForUpdate();
	if (chk.ok && !chk.updateAvailable) {
		ui.alert("已是最新版本（" + chk.installed + "）");
		return;
	}

	var resp = ui.alert(
		"更新腳本",
		"將更新到最新版" +
			(chk.latest ? "（" + chk.latest + "）" : "") +
			"。\n更新後請重新整理此頁面以套用。\n\n要繼續嗎？",
		ui.ButtonSet.OK_CANCEL,
	);
	if (resp !== ui.Button.OK) return;

	// Download the full bundle.
	var bundle;
	try {
		var res = UrlFetchApp.fetch(SELF_UPDATE_RAW_BASE + "/bundle.json", {
			muteHttpExceptions: true,
		});
		if (res.getResponseCode() !== 200) {
			ui.alert("下載失敗：bundle.json HTTP " + res.getResponseCode());
			return;
		}
		bundle = JSON.parse(res.getContentText());
	} catch (e) {
		ui.alert("下載或解析 bundle 失敗：" + e.message);
		return;
	}

	var v = validateBundle_(bundle, chk.latest);
	if (!v.ok) {
		ui.alert(
			"Bundle 驗證失敗：" + v.error + "\n（可能是 CDN 快取更新中，稍後再試）",
		);
		return;
	}

	// Overwrite this project's own content via the Apps Script API.
	var scriptId = ScriptApp.getScriptId();
	try {
		var put = UrlFetchApp.fetch(
			"https://script.googleapis.com/v1/projects/" + scriptId + "/content",
			{
				method: "put",
				contentType: "application/json",
				headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
				payload: JSON.stringify({ files: bundle.files }),
				muteHttpExceptions: true,
			},
		);
		var code = put.getResponseCode();
		if (code === 200) {
			ui.alert(
				"✅ 已更新到 " +
					(bundle.version || "最新版") +
					"。\n請重新整理此頁面以套用新版本。",
			);
			return;
		}
		var body = put.getContentText();
		if (code === 403 && body.indexOf("Apps Script API") !== -1) {
			// Surface Google's ACTUAL message + activation URL. The right place to
			// enable differs per clone: a default Apps-Script-managed project points
			// to script.google.com/home/usersettings, but a clone bound to a standard
			// GCP project must enable the API in THAT project's Cloud Console (the URL
			// carries the project number). Show whatever Google returns so the user is
			// never sent to the wrong toggle.
			var apiMsg = "";
			var apiUrl = "";
			try {
				var err = JSON.parse(body).error;
				apiMsg = err.message || "";
				var details = err.details || [];
				for (var d = 0; d < details.length; d++) {
					var meta = details[d].metadata;
					if (meta && meta.activationUrl) {
						apiUrl = meta.activationUrl;
						break;
					}
				}
			} catch (e) {
				// fall through with the raw body
			}
			ui.alert(
				"需要啟用 Apps Script API",
				"Google 回報：\n" +
					(apiMsg || body) +
					"\n\n啟用連結（複製到瀏覽器打開，用同一個帳號按 Enable）：\n" +
					(apiUrl ||
						"https://script.google.com/home/usersettings") +
					"\n\n啟用後等 1–2 分鐘，重新整理此頁面再跑一次更新。",
				ui.ButtonSet.OK,
			);
			return;
		}
		ui.alert("更新失敗（HTTP " + code + "）：\n" + body);
	} catch (e) {
		ui.alert("更新時發生錯誤：" + e.message);
	}
}

/**
 * Menu handler: install an on-open trigger that checks for updates. Simple
 * onOpen() can't use UrlFetchApp, so this uses an installable trigger.
 */
function enableUpdateOnOpenTrigger() {
	var ui = SlidesApp.getUi();
	try {
		var existing = ScriptApp.getProjectTriggers();
		for (var i = 0; i < existing.length; i++) {
			if (existing[i].getHandlerFunction() === "onOpenUpdateCheck_") {
				ui.alert("已經啟用開啟時檢查更新了。");
				return;
			}
		}
		ScriptApp.newTrigger("onOpenUpdateCheck_")
			.forPresentation(SlidesApp.getActivePresentation())
			.onOpen()
			.create();
		ui.alert("已啟用：之後開啟此簡報會自動檢查是否有新版本。");
	} catch (e) {
		ui.alert("無法啟用自動檢查：" + e.message);
	}
}

/** Menu handler: remove the on-open update-check trigger. */
function disableUpdateOnOpenTrigger() {
	var ui = SlidesApp.getUi();
	var triggers = ScriptApp.getProjectTriggers();
	var n = 0;
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === "onOpenUpdateCheck_") {
			ScriptApp.deleteTrigger(triggers[i]);
			n++;
		}
	}
	ui.alert(n ? "已關閉開啟時檢查更新。" : "目前沒有啟用開啟時檢查更新。");
}

/**
 * Installable on-open handler (runs with full auth, unlike simple onOpen).
 * Quietly checks for a newer version and nudges the user. Never blocks opening.
 */
function onOpenUpdateCheck_() {
	try {
		var r = checkForUpdate();
		if (r.ok && r.updateAvailable) {
			SlidesApp.getUi().alert(
				"有新版本 " +
					r.latest +
					"（目前 " +
					r.installed +
					"）\n用「⚙ 設定與批次 → 🔄 更新腳本」即可更新。",
			);
		}
	} catch (e) {
		// Never block presentation open.
	}
}
