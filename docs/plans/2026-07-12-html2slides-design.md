# HTML to Slides（html2slides）設計文件

日期：2026-07-12
狀態：已與使用者逐段確認

## 目標

比照 md2slides，新增「HTML to Slides」：接受一種**模組相容的純 HTML 格式
（gslide-html）**，對應 Google Slides 全部 11 種官方預設 layout，經解析後在目前
簡報中建立投影片。搭配兩個 agent skill：

1. 本 repo `.claude/skills/gslide-html/` — 教 agent 產生相容 HTML（規格單一事實來源）。
2. open-google-slide repo `packages/core/skills/export-gslide-html/` — 教 agent 把
   open-slide 的 React deck 語意重寫成相容 HTML（內嵌規格副本，檔頭標注上游 commit）。

## 已確認決策

| 決策點 | 結論 |
| --- | --- |
| HTML 契約 | 混合式：`data-layout`/`data-ph` 為正式契約；語意標籤推斷為 fallback |
| 解析位置 | 前端（dialog 內瀏覽器 `DOMParser`）→ 結構化 JSON → 後端建 slide |
| 輸入方式 | 貼上、上傳 .html、URL（後端 `UrlFetchApp` 抓回字串給前端解析；scope 已存在） |
| open-slide 整合 | skill + 輸出慣例（`slides/<id>/exports/gslide.html` + demo fixture） |
| 規格歸屬 | 主規格在本 repo skill 的 `reference.md`；open-slide 內嵌副本標注上游 |

## HTML 契約（gslide-html v1）

- 一份文件 = 一個 deck；每個 `<section>` = 一張 slide。
- `data-layout`：11 種 canonical 值（TITLE, SECTION_HEADER, TITLE_AND_BODY,
  TITLE_AND_TWO_COLUMNS, TITLE_ONLY, ONE_COLUMN_TEXT, MAIN_POINT,
  SECTION_TITLE_AND_DESCRIPTION, CAPTION_ONLY, BIG_NUMBER, BLANK）。
  接受別名（`section-header` 等），正規化規則同 `lzLayoutType`。
- `data-ph`：具名 slot — TITLE / SUBTITLE / BODY / BODY_2（雙欄右欄）/ NOTES。
- `data-notes` 屬性或 `<aside data-ph="NOTES">` = 講者備註。
- 推斷層（無標記時）：`<h1>` 單獨 → SECTION_HEADER；`<h1>+<p>` →
  SECTION_TITLE_AND_DESCRIPTION；`<h2>`+內容 → TITLE_AND_BODY；只有 `<h2>` →
  TITLE_ONLY；無 `<section>` 時以 h1/h2 切頁；`class="layout-*"` 等同 `data-layout`。
- 行內格式：`<strong>/<b>`、`<em>/<i>`、`<s>/<del>`、`<a href>` → 文字樣式；
  巢狀 `<ul>/<ol>` → 多層清單（深度 ≤ 3）。不支援標籤（img、table…）忽略並回報 warning。

## 模組架構

```
src/util/html2slides.js               # 進入點：dialog、fetchHtmlFromUrl()、convertGslideJsonToSlides()
src/util/html2slides/slideBuilder.js  # 後端 builder：JSON → slides
src/components/html2slides-dialog.html         # UI：貼上｜上傳｜URL 三個 tab
src/components/html2slides/parser-client.html  # 前端解析器（DOMParser），include() 共用
```

### 資料流

1. dialog 取得 HTML 字串（textarea / FileReader / `google.script.run.fetchHtmlFromUrl`）。
2. 前端 `parseGslideHtml(htmlString)` → 中介 JSON：
   `{ version: 1, slides: [{ layout, slots: { TITLE: {paragraphs}, ... }, notes, warnings }] }`
   paragraph = `{ runs: [{text, bold, italic, strikethrough, link}], listLevel, listType }`。
3. 後端 `convertGslideJsonToSlides(json)`：每頁
   `insertSlide(idx, SlidesApp.PredefinedLayout[layout])`（失敗時用 `lzLayoutType()`
   在 `presentation.getLayouts()` 找同型 layout；再不行 fallback TITLE_AND_BODY + warning）。
   Placeholder 對應動態解析：TITLE → TITLE/CENTERED_TITLE；SUBTITLE → SUBTITLE；
   BODY → 第 1 個 BODY；BODY_2 → 第 2 個 BODY；缺 placeholder 時 fallback 文字框。
   清單巢狀用段首 `\t` 縮排 + `applyListPreset`。
4. 回傳 `{created, warnings[]}`，dialog 顯示逐頁結果；單頁失敗不中斷整批。

### 測試

- 前端解析器寫成純函式，Node + linkedom 本地跑黃金範例（demo-deck.html）驗證。
- 後端照專案慣例 `clasp push` 後在簡報中手動測。

## Skill：.claude/skills/gslide-html/

- `SKILL.md`：觸發時機、工作流程（讀 reference → 選 layout 決策表 → 一律顯式
  data-* → 純 HTML 零 CSS）、驗證清單。
- `reference.md`：完整契約規格（= 規格文件本身，不另存 docs 副本）。
- `examples/demo-deck.html`：涵蓋 11 種 layout 的黃金範例（兼作解析器測試 fixture）。

## open-google-slide（第二階段，另一個 commit）

- `packages/core/skills/export-gslide-html/`：SKILL.md（React deck → 語意角色 →
  layout 對照 → 捨棄裝飾視覺 → 輸出 `slides/<id>/exports/gslide.html`）+
  reference.md 副本（標注 `Upstream: lizard-gslide-module@<commit>`）。
- `apps/demo` 加示範 deck 的人工校訂 `exports/gslide.html` 作黃金對照。
- README Export 段補一行說明可貼進本模組轉成 Google Slides。
