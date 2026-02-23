# 可更新的特殊頁碼格式 (Updatable Page Numbers)

## 概述

本專案實作了一套**可批次更新的頁碼系統**，透過 Google Slides API 的 `batchUpdate` 機制，在每次執行時先刪除舊的頁碼元素，再重新建立新的頁碼。這種「刪除-重建」的模式確保頁碼永遠與投影片的實際順序同步。

## 核心概念：為什麼不用內建頁碼？

Google Slides 內建的頁碼功能非常有限——它只能顯示頁碼數字，無法自訂格式（如 `3 / 15`）、無法自訂位置、無法自訂樣式。因此本專案選擇**用程式動態建立文字框**來取代內建頁碼。

## 關鍵檔案

| 檔案 | 角色 |
|------|------|
| `src/batch/element_generators.js` | **主要實作** — `addPageNumberUltra()` 函數（第 87-119 行） |
| `src/batch/page_number.js` | **舊版實作** — `appendPageNumberToSlide()` 函數（已標記 deprecated） |
| `src/batch/slide_utilities.js` | **刪除邏輯** — `batchDeleteAllElements()` 函數（第 145-173 行） |
| `src/batch/cache_manager.js` | **快取管理** — 預先計算尺寸和 GUID 池 |
| `src/batch/ultra_mega_batch.js` | **調度中心** — `runAllFunctionsUltraMegaBatch()` 函數 |

## 實作細節

### 1. 頁碼格式：`"當前頁 / 總頁數"`

頁碼的文字內容格式為 `${slideIndex + 1} / ${slideCache.totalSlides}`，例如 `3 / 15`。

```javascript
// src/batch/element_generators.js:101
{ insertText: { objectId: pageId, text: `${slideIndex + 1} / ${slideCache.totalSlides}` } }
```

### 2. 建立頁碼文字框的 API 請求序列

每個頁碼需要 4 個 batch request：

#### Step 1: 建立文字框形狀 (createShape)

```javascript
// src/batch/element_generators.js:91-99
{
    createShape: {
        objectId: pageId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
            pageObjectId: slideId,                              // 指定在哪一頁
            size: slideCache.sizes.pageNum,                     // { width: 70pt, height: 30pt }
            transform: {
                ...cache.transforms.identity,                   // { scaleX: 1, scaleY: 1, unit: 'PT' }
                translateX: 650,                                // 右下角 X 位置
                translateY: 370                                 // 右下角 Y 位置
            }
        }
    }
}
```

#### Step 2: 插入文字 (insertText)

```javascript
// src/batch/element_generators.js:101
{ insertText: { objectId: pageId, text: `${slideIndex + 1} / ${slideCache.totalSlides}` } }
```

#### Step 3: 設定文字樣式 (updateTextStyle)

```javascript
// src/batch/element_generators.js:102-110
{
    updateTextStyle: {
        objectId: pageId,
        textRange: { type: 'ALL' },                            // 套用到所有文字
        style: {
            bold: true,                                         // 粗體
            fontFamily: main_font_family,                       // 設定字型（如 "Source Sans Pro"）
            fontSize: { magnitude: 12, unit: 'PT' },           // 12pt 字體大小
            foregroundColor: {
                opaqueColor: { rgbColor: cache.colors.inactive } // 灰色文字 (#888888)
            }
        },
        fields: 'bold,fontFamily,fontSize,foregroundColor'      // 指定要更新的欄位
    }
}
```

#### Step 4: 設定段落對齊 (updateParagraphStyle)

```javascript
// src/batch/element_generators.js:112-116
{
    updateParagraphStyle: {
        objectId: pageId,
        textRange: { type: 'ALL' },
        style: { alignment: 'CENTER' },                        // 置中對齊
        fields: 'alignment'
    }
}
```

### 3. 唯一 ID 機制

每個頁碼元素都使用唯一的 `objectId`，確保不會衝突：

```javascript
// src/batch/element_generators.js:88
const pageId = `page_num_${slideId}_${getNextGuid()}`;
```

其中 `getNextGuid()` 來自 `cache_manager.js`，它使用**預生成的 GUID 池**（1000 個 UUID）來避免重複的 UUID 生成開銷：

```javascript
// src/batch/cache_manager.js:47-52
function getNextGuid() {
    const cache = initializeUltraCache();
    const guid = cache.guids[cache.guidIndex];
    cache.guidIndex = (cache.guidIndex + 1) % cache.guids.length;
    return guid;
}
```

### 4. 「刪除-重建」的更新機制

每次更新頁碼時，系統會先掃描所有投影片，找到以 `page_num_` 開頭的元素並刪除：

```javascript
// src/batch/slide_utilities.js:146-153
const deletePatterns = [
    'tab_', 'progress_', 'sections_', 'label_',
    'outline_', 'obj_', 'page_num_'                           // ← 頁碼的 prefix
];

// ...掃描每一頁的所有形狀
for (const shape of shapes) {
    const id = shape.getObjectId();
    const shouldDelete = deletePatterns.some(p => id.startsWith(p));
    if (shouldDelete) {
        requests.push({ deleteObject: { objectId: id } });
    }
}
```

### 5. 預計算的尺寸快取

頁碼文字框的尺寸在 `createUltraSlideCache()` 中預先計算，避免重複建立物件：

```javascript
// src/batch/cache_manager.js:69
sizes: {
    pageNum: {
        width: { magnitude: 70, unit: 'PT' },
        height: { magnitude: 30, unit: 'PT' }
    },
    // ...其他尺寸
}
```

## 你自己要做：如何實作自訂頁碼

### 範例 1：修改頁碼格式

如果你想要 `Page 3 of 15` 而不是 `3 / 15`：

```javascript
// 修改 insertText 的 text 參數
{ insertText: { objectId: pageId, text: `Page ${slideIndex + 1} of ${slideCache.totalSlides}` } }
```

### 範例 2：改變位置（左下角）

```javascript
transform: {
    scaleX: 1, scaleY: 1, unit: 'PT',
    translateX: 20,    // 左邊 20pt
    translateY: 370    // 底部
}
```

### 範例 3：跳過第一頁的頁碼

在 `generateAllElementsUltra()` 的迴圈中（`src/batch/ultra_mega_batch.js:68`），迴圈從 `i = 1` 開始，已經自動跳過第一頁（封面頁）：

```javascript
// src/batch/ultra_mega_batch.js:68
for (let i = 1; i < slideCache.totalSlides; i++) {
    // i 從 1 開始 → 第一頁 (index 0) 不會加上頁碼
}
```

### 範例 4：建立最小的頁碼系統

如果你要從零開始建立一個頁碼系統，核心邏輯如下：

```javascript
function addPageNumbers() {
    const presentation = SlidesApp.getActivePresentation();
    const presentationId = presentation.getId();
    const slides = presentation.getSlides();
    const requests = [];

    slides.forEach((slide, index) => {
        if (index === 0) return; // 跳過封面頁

        const slideId = slide.getObjectId();
        const pageId = `page_num_${slideId}_${Utilities.getUuid().slice(0, 8)}`;

        // 1. 建立文字框
        requests.push({
            createShape: {
                objectId: pageId,
                shapeType: 'TEXT_BOX',
                elementProperties: {
                    pageObjectId: slideId,
                    size: {
                        width: { magnitude: 70, unit: 'PT' },
                        height: { magnitude: 30, unit: 'PT' }
                    },
                    transform: {
                        scaleX: 1, scaleY: 1, unit: 'PT',
                        translateX: 650,
                        translateY: 370
                    }
                }
            }
        });

        // 2. 插入文字
        requests.push({
            insertText: {
                objectId: pageId,
                text: `${index + 1} / ${slides.length}`
            }
        });

        // 3. 設定文字樣式
        requests.push({
            updateTextStyle: {
                objectId: pageId,
                textRange: { type: 'ALL' },
                style: {
                    bold: true,
                    fontFamily: 'Source Sans Pro',
                    fontSize: { magnitude: 12, unit: 'PT' },
                    foregroundColor: {
                        opaqueColor: { rgbColor: { red: 0.53, green: 0.53, blue: 0.53 } }
                    }
                },
                fields: 'bold,fontFamily,fontSize,foregroundColor'
            }
        });

        // 4. 置中對齊
        requests.push({
            updateParagraphStyle: {
                objectId: pageId,
                textRange: { type: 'ALL' },
                style: { alignment: 'CENTER' },
                fields: 'alignment'
            }
        });
    });

    // 一次性送出所有請求
    if (requests.length) {
        Slides.Presentations.batchUpdate({ requests }, presentationId);
    }
}
```

## 重要的 API 觀念

### `fields` 參數

`fields` 參數是 Google Slides API 的**欄位遮罩 (field mask)**，只有在 `fields` 中列出的屬性會被更新。這是必填的，如果省略會導致 API 錯誤：

```javascript
// 正確：只更新 bold 和 fontSize
fields: 'bold,fontSize'

// 錯誤：缺少 fields 會導致 API 拒絕請求
```

### `textRange` 參數

- `{ type: 'ALL' }` — 套用到整個文字框的所有文字
- `{ type: 'FIXED_RANGE', startIndex: 0, endIndex: 5 }` — 套用到特定範圍

### Object ID 命名慣例

本專案使用 prefix 命名法來識別元素類型，方便後續刪除：

| Prefix | 元素類型 |
|--------|----------|
| `page_num_` | 頁碼 |
| `progress_` | 進度條 |
| `progress_bg_` | 進度條背景 |
| `tab_` | 標籤導航 |
| `sections_` | 章節文字框 |
| `label_` | 章節標籤 |
| `outline_` | 大綱 |
| `obj_` | 標題腳註 |

這個命名系統使得「刪除-重建」模式成為可能：只要 ID 以特定 prefix 開頭，就知道它是由程式建立的，可以安全刪除。
