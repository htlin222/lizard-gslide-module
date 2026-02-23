# 可更新的進度條 (Updatable Progress Bar)

## 概述

本專案實作了一個**視覺化進度條**，顯示在每張投影片的底部。進度條由兩個矩形組成：一個灰色背景條（代表全長）和一個主題色前景條（代表當前進度）。隨著投影片的推進，前景條的寬度會按比例增長。

## 進度條的視覺效果

```
投影片 1/5:  ██░░░░░░░░░░░░░░░░░░   (20%)
投影片 3/5:  ██████████████░░░░░░   (60%)
投影片 5/5:  ████████████████████   (100%)

█ = 主題色 (main_color, 如 #3D6869)
░ = 灰色背景 (#E0E0E0)
```

## 關鍵檔案

| 檔案 | 角色 |
|------|------|
| `src/batch/element_generators.js` | **主要實作** — `addProgressBarUltra()` 函數（第 33-82 行） |
| `src/batch/cache_manager.js` | **快取管理** — 預計算進度條 Y 座標和高度 |
| `src/config.js` | **設定** — `progressBarHeight` 變數（第 11 行） |
| `src/batch/ultra_mega_batch.js` | **調度中心** — 整合進度條到批次更新流程 |

## 實作細節

### 1. 進度比例計算

進度條的寬度基於**當前投影片在總頁數中的位置**：

```javascript
// src/batch/element_generators.js:34-35
const progressRatio = slideIndex / (slideCache.totalSlides - 1);
const barWidth = slideCache.width * progressRatio;
```

- `slideIndex = 1`（第二頁）, `totalSlides = 10` → `progressRatio = 1/9 ≈ 0.111`
- `slideIndex = 5`（第六頁）, `totalSlides = 10` → `progressRatio = 5/9 ≈ 0.556`
- `slideIndex = 9`（最後一頁）, `totalSlides = 10` → `progressRatio = 9/9 = 1.0`

注意：分母用 `totalSlides - 1` 而非 `totalSlides`，所以最後一頁的進度條是 100% 滿寬。

### 2. 完整的 API 請求序列（4 個 requests）

#### Request 1: 建立灰色背景條

```javascript
// src/batch/element_generators.js:42-50
{
    createShape: {
        objectId: bgId,
        shapeType: 'RECTANGLE',
        elementProperties: {
            pageObjectId: slideId,
            size: {
                height: slideCache.sizes.progressBar.height,  // { magnitude: 5, unit: 'PT' }
                width: { magnitude: slideCache.width, unit: 'PT' }  // 投影片全寬 (720pt)
            },
            transform: {
                ...cache.transforms.identity,     // { scaleX: 1, scaleY: 1, unit: 'PT' }
                translateX: 0,                     // 從最左邊開始
                translateY: slideCache.progressBarY // 投影片高度 - 進度條高度
            }
        }
    }
}
```

#### Request 2: 背景條樣式

```javascript
// src/batch/element_generators.js:52-59
{
    updateShapeProperties: {
        objectId: bgId,
        shapeProperties: {
            shapeBackgroundFill: {
                solidFill: { color: { rgbColor: cache.colors.gray } }  // #E0E0E0
            },
            outline: {
                weight: { magnitude: 0.1, unit: 'PT' },
                outlineFill: {
                    solidFill: { color: { rgbColor: cache.colors.gray } }
                }
            }
        },
        fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
    }
}
```

#### Request 3: 建立主題色前景條（進度指示）

```javascript
// src/batch/element_generators.js:61-70
{
    createShape: {
        objectId: progId,
        shapeType: 'RECTANGLE',
        elementProperties: {
            pageObjectId: slideId,
            size: {
                height: slideCache.sizes.progressBar.height,  // 同樣高度
                width: { magnitude: barWidth, unit: 'PT' }    // ★ 根據進度比例計算的寬度
            },
            transform: {
                ...cache.transforms.identity,
                translateX: 0,                     // 從最左邊開始（與背景條重疊）
                translateY: slideCache.progressBarY // 同樣的 Y 位置
            }
        }
    }
}
```

#### Request 4: 前景條樣式

```javascript
// src/batch/element_generators.js:72-80
{
    updateShapeProperties: {
        objectId: progId,
        shapeProperties: {
            shapeBackgroundFill: {
                solidFill: { color: { rgbColor: cache.colors.main } }  // 主題色 #3D6869
            },
            outline: {
                weight: { magnitude: 0.1, unit: 'PT' },
                outlineFill: {
                    solidFill: { color: { rgbColor: cache.colors.main } }
                }
            }
        },
        fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
    }
}
```

### 3. 進度條位置的預計算

進度條的 Y 座標在 `createUltraSlideCache()` 中預先計算：

```javascript
// src/batch/cache_manager.js:64
progressBarY: height - progressBarHeight
```

其中：
- `height` = `presentation.getPageHeight()` (標準投影片 ≈ 405pt)
- `progressBarHeight` = 來自全域設定 `src/config.js:11`（預設 5pt）

所以進度條位於投影片最底部 5pt 的範圍。

### 4. 可設定的進度條高度

使用者可以透過設定面板調整進度條高度：

```javascript
// src/config.js:11
var progressBarHeight = 5;  // 預設 5pt

// src/config.js:258-259 — 從 PropertiesService 讀取使用者自訂值
const savedProgressBarHeight = userProperties.getProperty(CONFIG_KEYS.PROGRESS_BAR_HEIGHT);
```

### 5. 顏色系統

進度條使用預先快取的顏色值，避免在迴圈中重複呼叫 `hexToRgb()`：

```javascript
// src/batch/cache_manager.js:19-24
colors: {
    main: hexToRgb(main_color),         // #3D6869 → { red: 0.239, green: 0.408, blue: 0.412 }
    inactive: hexToRgb('#888888'),       // 灰色文字
    gray: hexToRgb('#E0E0E0'),          // 進度條背景
    white: hexToRgb('#FFFFFF'),          // 白色
    bgColor: hexToRgb('#FFFFFF')
}
```

`hexToRgb()` 函數定義在 `src/batch/slide_utilities.js:178-187`：

```javascript
function hexToRgb(hex) {
    const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return m ? {
        red: parseInt(m[1], 16) / 255,     // Google API 需要 0-1 範圍
        green: parseInt(m[2], 16) / 255,
        blue: parseInt(m[3], 16) / 255
    } : { red: 0, green: 0, blue: 0 };
}
```

## 關鍵設計模式

### Outline 技巧

注意進度條矩形的 `outline` 被設為與填充色相同、且粗細只有 0.1pt。這是一個技巧：如果完全不設 outline，Google Slides 會使用預設的黑色邊框。設為相同顏色+極細粗細，視覺上就看不到邊框了。

### 背景條 + 前景條的疊加

兩個矩形在同一個位置，前景條疊在背景條上方。因為 batch request 中前景條在後面建立，所以自然會在上層。

## 你自己要做：如何實作進度條

### 範例 1：最小進度條實作

```javascript
function addProgressBars() {
    const presentation = SlidesApp.getActivePresentation();
    const presentationId = presentation.getId();
    const slides = presentation.getSlides();
    const slideWidth = presentation.getPageWidth();   // 720pt
    const slideHeight = presentation.getPageHeight(); // 405pt
    const barHeight = 5;
    const requests = [];

    const mainColor = { red: 0.239, green: 0.408, blue: 0.412 };  // #3D6869
    const grayColor = { red: 0.878, green: 0.878, blue: 0.878 };  // #E0E0E0

    slides.forEach((slide, index) => {
        if (index === 0) return;  // 跳過封面頁

        const slideId = slide.getObjectId();
        const progressRatio = index / (slides.length - 1);
        const barWidth = slideWidth * progressRatio;

        const bgId = `progress_bg_${slideId}_${Utilities.getUuid().slice(0, 8)}`;
        const progId = `progress_${slideId}_${Utilities.getUuid().slice(0, 8)}`;

        // 背景條
        requests.push(
            {
                createShape: {
                    objectId: bgId, shapeType: 'RECTANGLE',
                    elementProperties: {
                        pageObjectId: slideId,
                        size: {
                            height: { magnitude: barHeight, unit: 'PT' },
                            width: { magnitude: slideWidth, unit: 'PT' }
                        },
                        transform: {
                            scaleX: 1, scaleY: 1, unit: 'PT',
                            translateX: 0,
                            translateY: slideHeight - barHeight
                        }
                    }
                }
            },
            {
                updateShapeProperties: {
                    objectId: bgId,
                    shapeProperties: {
                        shapeBackgroundFill: { solidFill: { color: { rgbColor: grayColor } } },
                        outline: {
                            weight: { magnitude: 0.1, unit: 'PT' },
                            outlineFill: { solidFill: { color: { rgbColor: grayColor } } }
                        }
                    },
                    fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
                }
            }
        );

        // 前景條
        requests.push(
            {
                createShape: {
                    objectId: progId, shapeType: 'RECTANGLE',
                    elementProperties: {
                        pageObjectId: slideId,
                        size: {
                            height: { magnitude: barHeight, unit: 'PT' },
                            width: { magnitude: barWidth, unit: 'PT' }
                        },
                        transform: {
                            scaleX: 1, scaleY: 1, unit: 'PT',
                            translateX: 0,
                            translateY: slideHeight - barHeight
                        }
                    }
                }
            },
            {
                updateShapeProperties: {
                    objectId: progId,
                    shapeProperties: {
                        shapeBackgroundFill: { solidFill: { color: { rgbColor: mainColor } } },
                        outline: {
                            weight: { magnitude: 0.1, unit: 'PT' },
                            outlineFill: { solidFill: { color: { rgbColor: mainColor } } }
                        }
                    },
                    fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
                }
            }
        );
    });

    if (requests.length) {
        Slides.Presentations.batchUpdate({ requests }, presentationId);
    }
}
```

### 範例 2：進度條放在頂部

只需改變 `translateY`：

```javascript
translateY: 0  // 頂部
// 而不是
translateY: slideHeight - barHeight  // 底部
```

### 範例 3：分段式進度條

如果你想要每個章節一個區段，而不是連續的進度條：

```javascript
// 假設有 5 個章節，每個章節佔 1/5 的寬度
const sectionWidth = slideWidth / totalSections;
const currentSectionWidth = sectionWidth * (currentSectionIndex + 1);

// 前景條寬度 = 已完成的章節寬度
width: { magnitude: currentSectionWidth, unit: 'PT' }
```

## 效能考量

### Batch Update 的重要性

本專案將**所有投影片的所有元素**（頁碼、進度條、標籤、章節等）收集到一個 `requests` 陣列中，然後用**一次 `batchUpdate` API 呼叫**送出：

```javascript
// src/batch/ultra_mega_batch.js:44-49
if (requests.length) {
    Logger.log(`Ultra batch: ${requests.length} operations in 1 API call`);
    Slides.Presentations.batchUpdate({ requests }, presentationId);
}
```

這比逐一呼叫 API 快了 10-50 倍。一個 20 頁的投影片可能產生 200+ 個 requests，但只需要 1 次 API 呼叫。
