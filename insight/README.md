# Google Slides API 研究文件 (Insight Documents)

本目錄包含對 `lizard-gslide-module` 專案程式碼的深度分析，旨在幫助開發者理解如何透過 Google Slides API 實作各種自動化功能，並以此為基礎建構自己的投影片模板系統。

## 文件索引

| 編號 | 文件 | 主題 | 核心概念 |
|------|------|------|---------|
| 01 | [01-page-number.md](./01-page-number.md) | 可更新的特殊頁碼格式 | createShape → insertText → updateTextStyle 的完整流程、Object ID 命名慣例、刪除-重建模式 |
| 02 | [02-progress-bar.md](./02-progress-bar.md) | 可更新的進度條 | 雙矩形疊加模式、進度比例計算、Outline 技巧、Batch Update 效能 |
| 03 | [03-sections.md](./03-sections.md) | 可更新的章節系統 | Section 偵測機制、FIXED_RANGE 文字樣式、可點擊的標籤導航、Outline 自動生成 |
| 04 | [04-object-positioning.md](./04-object-positioning.md) | 物件定位方法 | 座標系統 (720×405pt)、Transform 矩陣、兩套 API 對比、旋轉公式、常用位置參考表 |
| 05 | [05-template-operations.md](./05-template-operations.md) | 模板操作 | 主題匯入流程、PropertiesService 持久化、6 種預定義樣式、HTML Sidebar 模組化 |

## 適合誰閱讀

- 想要用 Google Apps Script 自動化投影片的開發者
- 想要建立自己的投影片模板系統的團隊
- 想要理解 Google Slides API (REST) 和 SlidesApp API (高階) 差異的人

## 核心設計原則

本專案的架構基於以下設計原則：

1. **Batch Update 為王** — 所有動態元素（頁碼、進度條、標籤、章節）都透過一次 `Slides.Presentations.batchUpdate()` 呼叫完成，而不是逐一操作
2. **刪除-重建模式** — 每次更新時先刪除舊元素、再建立新元素，確保內容永遠同步
3. **Object ID 命名慣例** — 使用 prefix（如 `page_num_`、`progress_`）來識別程式建立的元素
4. **快取最大化** — 預計算顏色、尺寸、GUID 池，避免在迴圈中重複計算
5. **設定持久化** — 使用 `PropertiesService` 儲存使用者偏好，全域變數作為執行期快取

## 專案檔案結構對照

```
src/
├── config.js                          → 設定管理、選單、全域變數
├── batch/
│   ├── ultra_mega_batch.js            → 批次處理主控制器
│   ├── cache_manager.js               → 快取管理（顏色、GUID、尺寸）
│   ├── element_generators.js          → 頁碼、進度條、標籤、腳註
│   ├── section_elements.js            → 章節文字框、標籤、大綱
│   ├── slide_utilities.js             → 章節偵測、元素刪除、hexToRgb
│   ├── theme.js                       → 主題匯入
│   ├── page_number.js                 → 頁碼（舊版）
│   └── toggle_watermark.js            → 浮水印
├── util/
│   ├── default_style.js               → 6 種預定義樣式
│   ├── html_service_utils.js          → HTML 模組化工具
│   ├── date.js                        → 日期更新
│   ├── numbering.js                   → 數字遞增圓圈
│   ├── boxes.js                       → 批次建立方塊
│   ├── toggle_grids.js                → 網格切換
│   ├── average_padding.js             → 智能置中
│   └── graph/
│       ├── shape_creator.js           → 子形狀建立
│       ├── split_shape.js             → 形狀分割
│       └── setGap.js                  → 間距調整
└── components/
    ├── config-sidebar/                → 設定面板 HTML
    ├── flowchart/                     → 流程圖 HTML
    └── markdown-sidebar/             → Markdown 側邊欄 HTML
```
