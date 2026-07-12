---
name: gslide-html
description: Generate lizard-gslide-module compatible pure HTML (gslide-html v1) that maps to the 11 official Google Slides layouts. Use when the user wants slide content authored as HTML for the "HTML to Slides" importer, asks for "gslide html" / "模組相容 HTML" / "轉成 Google Slides 的 HTML", or when converting an outline, document, or deck into importable slide HTML.
---

# gslide-html 產生器

把任意內容（大綱、文件、既有簡報）改寫成 lizard-gslide-module「HTML to Slides」
可匯入的純 HTML。

## 工作流程

1. **先讀規格**：本 skill 目錄的 `reference.md`（契約單一事實來源），黃金範例在
   `examples/demo-deck.html`。
2. **規劃頁面**：把內容切成投影片。先判斷是否適合 minter section
   （`reference.md` 的 minter 型錄）：時間軸、步驟、KPI、比較欄、表格、卡片網格、
   圖片牆等結構化內容一律用 `data-minter`，視覺效果遠勝純條列。其餘才選 layout——

   | 內容性質 | layout |
   | --- | --- |
   | 開場封面（deck 標題＋副標） | `TITLE` |
   | 章節開頭（只有章節名） | `SECTION_HEADER` |
   | 章節開頭＋一段介紹 | `SECTION_TITLE_AND_DESCRIPTION` |
   | 一般條列內容 | `TITLE_AND_BODY` |
   | 左右對比、並列、before/after | `TITLE_AND_TWO_COLUMNS` |
   | 單一金句、結論、call-to-action | `MAIN_POINT` |
   | 關鍵數字（KPI、統計） | `BIG_NUMBER`（TITLE 放數字，BODY 放註解） |
   | 長段落敘述文字 | `ONE_COLUMN_TEXT` |
   | 之後要手動貼大圖的頁 | `TITLE_ONLY` 或 `CAPTION_ONLY` |
   | 佔位空頁 | `BLANK` |

3. **輸出規則**：
   - 一律使用**顯式** `data-layout` ＋ `data-ph`（推斷層是給外部 HTML 的容錯，
     agent 自產不得依賴）。
   - 純 HTML：零 CSS、零 `<style>`、零 `<script>`、零行內 style。
   - 行內格式只用白名單：`<strong>` `<em>` `<s>` `<a>` `<ul>/<ol>/<li>`（巢狀 ≤ 3 層）。
   - 講者備註用 `data-notes` 或 `<aside data-ph="NOTES">`。
4. **自檢**：跑一遍 `reference.md` 末尾的驗證清單，特別是「不要整份都是
   TITLE_AND_BODY」——好簡報會混用 SECTION_HEADER / MAIN_POINT / BIG_NUMBER 做節奏。

## 匯入方式（告知使用者）

Google Slides 選單 **🔁 跨頁與匯出 → 🌐 HTML 轉換成投影片**，貼上／上傳 .html／給 URL 皆可。
