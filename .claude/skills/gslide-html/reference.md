# gslide-html v1 — 模組相容 HTML 契約規格

此文件是 lizard-gslide-module「HTML to Slides」輸入格式的**單一事實來源**。
解析器：`src/components/html2slides/parser-client.html`（前端 DOMParser）。
後端 builder：`src/util/html2slides/slideBuilder.js`。

## 總則

- 一份 HTML 文件 = 一個 deck；每個頂層 `<section>` = 一張投影片。
- **純 HTML、零 CSS、零 JavaScript**。`<style>`、`<script>`、`class` 樣式一律不需要
  （解析器會忽略樣式；只有 `class="layout-*"` 有語意）。
- 編碼 UTF-8；`<head>`/`<body>` 可有可無，解析器兩者皆容忍。

## 顯式契約（agent 產出時一律使用）

### `data-layout` — 每個 `<section>` 必填

11 種 canonical 值（等同 Google `PredefinedLayout` enum）：

| 值 | 用途 | 合法 slot |
| --- | --- | --- |
| `TITLE` | 開場封面（主標＋副標） | TITLE, SUBTITLE |
| `SECTION_HEADER` | 章節分隔頁 | TITLE, SUBTITLE |
| `TITLE_AND_BODY` | 最常用的內容頁 | TITLE, BODY |
| `TITLE_AND_TWO_COLUMNS` | 左右對比／並列 | TITLE, BODY, BODY_2 |
| `TITLE_ONLY` | 只有標題（下方自行留白） | TITLE |
| `ONE_COLUMN_TEXT` | 單窄欄文字 | TITLE, BODY |
| `MAIN_POINT` | 單一金句／重點 | TITLE |
| `SECTION_TITLE_AND_DESCRIPTION` | 章節標題＋說明 | TITLE, SUBTITLE, BODY |
| `CAPTION_ONLY` | 底部小字說明（配大圖用） | BODY |
| `BIG_NUMBER` | 關鍵數字 | TITLE（數字）, BODY（註解） |
| `BLANK` | 空白頁 | （無） |

也接受寬鬆別名：大小寫不拘、`-`/空白視同 `_`（`section-header` → `SECTION_HEADER`）。

### `data-ph` — 元素對應 placeholder slot

- `TITLE`／`SUBTITLE`／`BODY`／`BODY_2`（雙欄 layout 的右欄）。
- 掛在任何元素上皆可；元素的**內文**（含行內格式與清單）進入該 slot。
- 同一 slot 出現多次時，內容依序串接為多個段落。

### 講者備註

- `<section data-notes="一句話備註">`，或
- `<aside data-ph="NOTES">多段備註，支援段落</aside>`（`<aside>` 不會出現在投影片上）。

### 範例

```html
<section data-layout="TITLE_AND_TWO_COLUMNS" data-notes="強調左右差異">
  <h2 data-ph="TITLE">方案比較</h2>
  <div data-ph="BODY">
    <ul><li>方案 A：快</li><li>成本低</li></ul>
  </div>
  <div data-ph="BODY_2">
    <ul><li>方案 B：穩</li><li>可擴充</li></ul>
  </div>
</section>
```

## 推斷層（fallback，處理無標記的外部 HTML；agent 自產禁止依賴）

- `class="layout-big-number"`（或 id）等同 `data-layout="BIG_NUMBER"`。
- `<section>` 無任何標記時：
  - 只有 `<h1>` → `SECTION_HEADER`
  - `<h1>` ＋其他內容 → `SECTION_TITLE_AND_DESCRIPTION`
  - `<h2>`（或 `<h3>`）＋內容 → `TITLE_AND_BODY`
  - 只有 `<h2>` → `TITLE_ONLY`
  - 無標題有內容 → `TITLE_AND_BODY`（無標題）
- slot 推斷：第一個 heading → `TITLE`；其餘流內容 → 該 layout 的第一個內容 slot
  （`SECTION_HEADER`/`TITLE` 為 `SUBTITLE`，其餘為 `BODY`）。
- 整份文件沒有 `<section>` 時：以 `<h1>`/`<h2>` 為界自動切頁（同 md2slides 規則）。

## 行內格式（白名單）

| HTML | 效果 |
| --- | --- |
| `<strong>` `<b>` | 粗體 |
| `<em>` `<i>` | 斜體 |
| `<s>` `<del>` `<strike>` | 刪除線 |
| `<a href="...">` | 超連結 |
| `<br>` | 段內換行（轉為新段落） |
| `<ul>`/`<ol>`/`<li>` | 清單；巢狀深度 ≤ 3；`<ol>` 轉編號清單 |

**不支援**（v1 忽略並回報 warning）：`<img>`、`<table>`、`<pre>`/`<code>` 區塊、
`<video>`、`<iframe>`、行內 style。文字內容仍會被抽出（img/iframe 除外）。

## Minter section（v1.1 — 進階版面元件）

`<section data-minter="KEY">` 不走 11 種 layout，改由模組的 minter 在該頁繪製
styled 元件（時間軸、KPI 卡、比較欄…）。基底頁自動選 `TITLE_ONLY`（有標題時）
或 `BLANK`；也可用 `data-layout` 指定。標題、`data-notes`/`<aside data-ph="NOTES">`
照常可用。

通用語法：
- **items**：section 內的頂層 `<ul>/<ol>`，每個 `<li>` 一個項目。`<li>` 的行內文字
  = 主文字（label/title），`data-*` 屬性 = 欄位，`<li>` 內的 `<p>` = 描述（desc）。
- **選項**：`data-minter-options='{"templateId":"...","cols":3}'`（JSON）；
  捷徑屬性 `data-template`、`data-orientation`。

| KEY | 用途 | 內容來源 | 常用欄位／選項 |
| --- | --- | --- | --- |
| `timeline` | 時間軸 | `<li data-date="2024 Q1">事件</li>` | `orientation: horizontal\|vertical` |
| `steps` | 步驟流程 | `<li>步驟<p>說明</p></li>` | `orientation` |
| `takeaways` | 重點帶走卡 | `<li>重點<p>說明</p></li>` | `heading`（預設用標題） |
| `compare` | 左右比較欄 | ≥2 個 `<div data-col="欄標題"><ul>…` | `templateId` |
| `kpi` | KPI 數字卡 | `<li data-value="99.9%" data-trend="up">標籤</li>` | `templateId` |
| `barchart` | 長條圖 | `<li data-value="42">類別</li>` | `orientation`、`showValues` |
| `agenda` | 議程列表 | `<li>議程項目</li>` | `templateId` |
| `table` | 樣式表格 | 真正的 `<table>`（thead+tbody） | `theme`、`fontSize` |
| `grid` | 卡片網格 | `<li data-subtitle="副標">卡片標題<p>內文</p></li>` | `rows`、`cols`、`styleNumber` |
| `gallery` | 圖片牆 | `<img src>` 或 `<figure><img><figcaption>` | `cols`、`captions` |
| `callout` | 重點框 | 標題＋`<p>` 內文 | `header`（預設用標題）、`templateId` |
| `icon` | 單一大圖示 | 無內容，全靠選項 | `{"glyph":"★","size":96,"color":"#3D6869"}` |

範例：

```html
<section data-minter="timeline" data-orientation="horizontal" data-notes="按季講">
  <h2 data-ph="TITLE">2024 里程碑</h2>
  <ul>
    <li data-date="Q1">專案啟動</li>
    <li data-date="Q3">Beta 上線</li>
  </ul>
</section>

<section data-minter="table">
  <h2 data-ph="TITLE">季度數字</h2>
  <table>
    <thead><tr><th>季</th><th>營收</th></tr></thead>
    <tbody><tr><td>Q1</td><td>100</td></tr></tbody>
  </table>
</section>
```

選型原則：時間序列→`timeline`、流程→`steps`、量化對比→`barchart`/`kpi`、
質性對比→`compare`、資料明細→`table`、多卡片並列→`grid`——**不要把這些內容
硬塞成 TITLE_AND_BODY 的條列**。

## 產出驗證清單（agent 自檢）

1. 每個 `<section>` 都有 `data-layout`，值在 11 種 canonical 之內。
2. 每個內容元素都有 `data-ph`，slot 名在該 layout 的合法集合內（見上表）。
3. 無 `<style>`/`<script>`/行內 style；無不支援標籤。
4. 巢狀清單 ≤ 3 層；`BODY_2` 只用於 `TITLE_AND_TWO_COLUMNS`。
5. 開場用 `TITLE`、章節用 `SECTION_HEADER`、金句用 `MAIN_POINT`、
   關鍵數據用 `BIG_NUMBER` — 不要整份都是 `TITLE_AND_BODY`。
