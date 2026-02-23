# 可更新的章節系統 (Updatable Sections & Chapters)

## 概述

本專案實作了一套完整的**章節管理系統**，包含四個可更新的組件：

1. **Section Header（章節頁）** — 在章節投影片上顯示所有章節列表，高亮當前章節
2. **Tab Navigation（標籤導航列）** — 在每一般投影片頂部顯示可點擊的章節導航條
3. **Section Label（章節標籤）** — 在章節頁左上角顯示 `Section: N`
4. **Outline（大綱頁）** — 如果第二頁標題是 "Outline"，自動生成章節列表

## 關鍵檔案

| 檔案 | 角色 |
|------|------|
| `src/batch/slide_utilities.js` | **章節偵測** — `getSectionHeadersUltra()` 函數（第 12-35 行） |
| `src/batch/section_elements.js` | **章節頁元素** — Section Box、Label、Outline（第 1-285 行） |
| `src/batch/element_generators.js` | **標籤導航** — `addTabNavigationUltra()` 函數（第 180-284 行） |
| `src/batch/ultra_mega_batch.js` | **調度邏輯** — 追蹤當前章節索引（第 64-86 行） |

## 章節偵測機制

### 如何識別「章節頁」

系統透過投影片的 **Layout 名稱** 來辨識章節頁——當 `layoutName === "SECTION_HEADER"` 時，該投影片被視為一個章節的開始：

```javascript
// src/batch/slide_utilities.js:12-35
function getSectionHeadersUltra(slides) {
    const sections = [];
    for (let i = 0; i < slides.length; i++) {
        const slide = slides[i];
        if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') {
            // 從文字框中提取章節標題
            const shapes = slide.getShapes();
            for (const shape of shapes) {
                if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
                    const text = shape.getText().asString().trim();
                    if (text) {
                        sections.push({
                            title: text,           // 章節標題
                            index: i,              // 在所有投影片中的索引
                            slideId: slide.getObjectId()  // 投影片 ID
                        });
                        break;  // 找到第一個有文字的文字框就停止
                    }
                }
            }
        }
    }
    return sections;
}
```

回傳的 `sections` 陣列範例：

```javascript
[
    { title: "Introduction",   index: 2,  slideId: "g1a2b3c4" },
    { title: "Methodology",    index: 7,  slideId: "g5d6e7f8" },
    { title: "Results",        index: 12, slideId: "g9h0i1j2" },
    { title: "Conclusion",     index: 18, slideId: "gk3l4m5n" }
]
```

### 追蹤「當前章節」

在主迴圈中，系統動態追蹤每張投影片屬於哪個章節：

```javascript
// src/batch/ultra_mega_batch.js:64-82
let currentSectionIdx = -1;  // 開始時還沒有進入任何章節

for (let i = 1; i < slideCache.totalSlides; i++) {
    // 當遇到下一個章節的起始頁時，更新章節索引
    if (currentSectionIdx + 1 < sectionsCache.length &&
        i >= sectionsCache[currentSectionIdx + 1].index) {
        currentSectionIdx++;
    }

    // 將當前章節索引傳遞給元素生成器
    generateSlideElementsUltra(slideId, slideData, i, slideCache,
        sectionsCache, currentSectionIdx, requests, cache);
}
```

## 組件 1：Section Box（章節文字框）

### 功能描述

在每一張章節頁上，建立一個文字框列出**所有章節標題**，其中：
- **之前的章節** → 灰色、30pt
- **當前章節** → 主題色、粗體、36pt
- **之後的章節** → 黑色、30pt

### 實作細節

```javascript
// src/batch/section_elements.js:56-158 — addUnifiedSectionBox()

// Step 1: 建立所有章節標題的文字（帶編號）
const lines = sectionsCache.map((s, i) => `${i + 1}. ${s.title}`);
const fullText = lines.join('\n');
// 例如: "1. Introduction\n2. Methodology\n3. Results\n4. Conclusion"

// Step 2: 建立文字框
requests.push({
    createShape: {
        objectId: boxId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
            pageObjectId: slideId,
            size: {
                width: { magnitude: 500, unit: 'PT' },     // 寬 500pt
                height: { magnitude: 300, unit: 'PT' }      // 高 300pt
            },
            transform: {
                ...cache.transforms.identity,
                translateX: 200,                              // 水平偏移 200pt
                translateY: (405 - 300) / 2                   // 垂直置中
            }
        }
    }
});

// Step 3: 插入全部文字
requests.push({ insertText: { objectId: boxId, text: fullText } });

// Step 4: 設定形狀背景為白色
requests.push({
    updateShapeProperties: {
        objectId: boxId,
        shapeProperties: {
            contentAlignment: 'MIDDLE',                       // 垂直置中
            shapeBackgroundFill: {
                solidFill: { color: { rgbColor: cache.colors.white } }
            },
            outline: {
                outlineFill: {
                    solidFill: { color: { rgbColor: cache.colors.white } }
                }
            }
        },
        fields: 'contentAlignment,shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color'
    }
});
```

### 逐行套用不同樣式（FIXED_RANGE 技巧）

這是最關鍵的技巧——用 `charIndex` 追蹤每一行的字元位置，精準地為每行文字套用不同樣式：

```javascript
// src/batch/section_elements.js:120-157
let charIndex = 0;
lines.forEach((line, i) => {
    const startIndex = charIndex;
    const endIndex = charIndex + line.length;

    // 根據相對於當前章節的位置選擇樣式
    let style;
    if (i < currentIdx) {
        style = { fontSize: 30, color: cache.colors.inactive, bold: false };  // 之前的章節
    } else if (i === currentIdx) {
        style = { fontSize: 36, color: cache.colors.main, bold: true };       // 當前章節
    } else {
        style = { fontSize: 30, color: { red: 0, green: 0, blue: 0 }, bold: false }; // 之後的章節
    }

    requests.push({
        updateTextStyle: {
            objectId: boxId,
            textRange: {
                type: 'FIXED_RANGE',
                startIndex: startIndex,        // 這一行的起始字元位置
                endIndex: endIndex             // 這一行的結束字元位置
            },
            style: {
                fontSize: { magnitude: style.fontSize, unit: 'PT' },
                fontFamily: main_font_family,
                foregroundColor: { opaqueColor: { rgbColor: style.color } },
                bold: style.bold
            },
            fields: 'fontSize,fontFamily,foregroundColor,bold'
        }
    });

    // 移動到下一行（+1 是 \n 換行字元）
    charIndex = endIndex + 1;
});
```

> **重要觀念**：Google Slides API 的 `textRange` 使用的是**字元索引**，不是行索引。所以你必須自己計算每一行在整個文字中的起始和結束位置。`\n` 也佔一個字元位置。

## 組件 2：Tab Navigation（標籤導航列）

### 功能描述

在每張一般投影片（非章節頁）的頂部，顯示一排可點擊的標籤，每個標籤代表一個章節。當前章節的標籤用**主題色背景 + 白色文字**高亮顯示。

### 實作細節

```javascript
// src/batch/element_generators.js:180-284 — addTabNavigationUltra()

// Step 1: 計算每個標籤的寬度
const estCharW = 8 * 0.75;  // 字元寬度估算 = 字體大小 × 0.75
const widths = sections.map(sec =>
    Math.max(sec.title.length * estCharW, 50) + 5  // 最小 50pt + 5pt 緩衝
);
const totalWidth = widths.reduce((a, b) => a + b, 0);
const xStart = Math.max((720 - totalWidth) / 2, 0);  // 置中

// Step 2: 建立白色背景條
requests.push({
    createShape: {
        objectId: bgId, shapeType: 'RECTANGLE',
        elementProperties: {
            pageObjectId: slideId,
            size: {
                height: { magnitude: 14, unit: 'PT' },   // 14pt 高
                width: { magnitude: 720, unit: 'PT' }     // 全寬
            },
            transform: { ...cache.transforms.identity, translateX: 0, translateY: 0 }
        }
    }
});

// Step 3: 為每個章節建立標籤
let xPos = xStart;
sections.forEach((sec, idx) => {
    const isActive = idx === currentSection;
    const tabId = `tab_${slideId}_${getNextGuid()}`;

    // 建立標籤文字框
    requests.push({
        createShape: {
            objectId: tabId, shapeType: 'TEXT_BOX',
            elementProperties: {
                pageObjectId: slideId,
                size: {
                    height: { magnitude: 14, unit: 'PT' },
                    width: { magnitude: widths[idx], unit: 'PT' }
                },
                transform: {
                    ...cache.transforms.identity,
                    translateX: xPos,        // 動態累加的 X 位置
                    translateY: 0
                }
            }
        }
    });

    // 插入章節標題文字
    requests.push({ insertText: { objectId: tabId, text: sec.title } });

    // 設定背景色（活躍章節 = 主題色，其他 = 白色）
    requests.push({
        updateShapeProperties: {
            objectId: tabId,
            shapeProperties: {
                shapeBackgroundFill: {
                    solidFill: {
                        color: {
                            rgbColor: isActive ? cache.colors.main : cache.colors.white
                        }
                    }
                },
                contentAlignment: 'MIDDLE'
            },
            fields: 'shapeBackgroundFill.solidFill.color,contentAlignment'
        }
    });

    // 設定文字樣式（含 link 連結到章節頁）
    requests.push({
        updateTextStyle: {
            objectId: tabId, textRange: { type: 'ALL' },
            style: {
                bold: true,
                fontFamily: main_font_family,
                fontSize: { magnitude: 8, unit: 'PT' },
                foregroundColor: {
                    opaqueColor: {
                        rgbColor: isActive ? cache.colors.white : cache.colors.inactive
                    }
                },
                underline: false,
                link: { pageObjectId: sec.slideId }  // ★ 點擊跳轉到章節頁
            },
            fields: 'bold,fontFamily,fontSize,foregroundColor,underline,link'
        }
    });

    xPos += widths[idx];  // 累加 X 位置
});
```

### 重要特性：可點擊的連結

標籤文字使用了 `link: { pageObjectId: sec.slideId }` 來建立**投影片內部連結**。在播放模式下，點擊標籤可以直接跳轉到對應的章節頁。同時設定 `underline: false` 來移除預設的底線樣式。

## 組件 3：Section Label（章節標籤）

```javascript
// src/batch/section_elements.js:163-219 — addSectionLabel()

// 在章節頁左上角建立一個小標籤
const labelId = `label_${slideId}_${getNextGuid()}`;

requests.push({
    createShape: {
        objectId: labelId, shapeType: 'TEXT_BOX',
        elementProperties: {
            pageObjectId: slideId,
            size: {
                width: { magnitude: 80, unit: 'PT' },
                height: { magnitude: 25, unit: 'PT' }
            },
            transform: { ...cache.transforms.identity, translateX: 50, translateY: 50 }
        }
    }
});

// 文字內容：如 "Section: 3"
requests.push({ insertText: { objectId: labelId, text: `Section: ${sectionNumber}` } });

// 背景色 = 主題色，文字 = 白色
requests.push({
    updateShapeProperties: {
        objectId: labelId,
        shapeProperties: {
            contentAlignment: 'MIDDLE',
            shapeBackgroundFill: { solidFill: { color: { rgbColor: cache.colors.main } } }
        },
        fields: 'contentAlignment,shapeBackgroundFill.solidFill.color'
    }
});
```

## 組件 4：Outline 大綱頁

如果第二張投影片的標題是 `"Outline"`，系統會自動在該頁建立一個**帶項目符號的章節列表**：

```javascript
// src/batch/section_elements.js:225-285 — addOutlineToSecondSlide()

// 檢查第二張投影片的標題
const title = getOutlineSlideTitle(secondSlide);

if (title === 'Outline') {
    const outlineTitles = sectionsCache.map(s => s.title);
    // outlineTitles = ["Introduction", "Methodology", "Results", "Conclusion"]

    // 建立文字框
    requests.push({
        createShape: {
            objectId: outlineId, shapeType: 'TEXT_BOX',
            elementProperties: {
                pageObjectId: secondSlide.getObjectId(),
                size: {
                    width: { magnitude: 400, unit: 'PT' },
                    height: { magnitude: 300, unit: 'PT' }
                },
                transform: { ...cache.transforms.identity, translateX: 280, translateY: 51 }
            }
        }
    });

    // 插入所有章節標題（換行分隔）
    requests.push({
        insertText: { objectId: outlineId, text: outlineTitles.join('\n') }
    });

    // 新增項目符號
    requests.push({
        createParagraphBullets: {
            objectId: outlineId,
            textRange: { type: 'ALL' },
            bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE'   // 圓形 → 空心圓 → 方形（多層級）
        }
    });
}
```

## 你自己要做：如何實作章節系統

### 範例 1：只要標籤導航列

```javascript
function addTabNavigation() {
    const presentation = SlidesApp.getActivePresentation();
    const slides = presentation.getSlides();
    const requests = [];

    // Step 1: 偵測章節
    const sections = [];
    slides.forEach((slide, i) => {
        if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') {
            const shapes = slide.getShapes();
            for (const shape of shapes) {
                const text = shape.getText().asString().trim();
                if (text) {
                    sections.push({ title: text, index: i, slideId: slide.getObjectId() });
                    break;
                }
            }
        }
    });

    if (sections.length === 0) return;

    // Step 2: 為每頁建立標籤列
    let currentSection = -1;
    slides.forEach((slide, i) => {
        if (i === 0) return;
        if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') return;

        // 更新當前章節
        if (currentSection + 1 < sections.length && i >= sections[currentSection + 1].index) {
            currentSection++;
        }

        // 為這一頁加上標籤列...
        // （使用上述 addTabNavigationUltra 的邏輯）
    });

    Slides.Presentations.batchUpdate({ requests }, presentation.getId());
}
```

### 範例 2：自訂章節偵測方式

如果你不用 `SECTION_HEADER` layout，可以改用其他方式偵測章節：

```javascript
// 方法 A：用特定文字前綴偵測
if (text.startsWith('##')) {
    sections.push({ title: text.replace('##', '').trim(), ... });
}

// 方法 B：用 Speaker Notes 偵測
const notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString();
if (notes.includes('[SECTION]')) {
    sections.push({ title: extractTitle(slide), ... });
}

// 方法 C：用形狀標題 (alt text) 偵測
const shapes = slide.getShapes();
for (const shape of shapes) {
    if (shape.getTitle() === 'SECTION_MARKER') {
        sections.push({ title: shape.getText().asString().trim(), ... });
    }
}
```

## 架構圖

```
onOpen()
  └─ createCustomMenu()
       ├─ "🛠 同時執行所有功能" → confirmRunAll()
       │   └─ runAllFunctionsUltraMegaBatch()
       │       ├─ getSectionHeadersUltra()     ← 偵測所有章節
       │       ├─ batchDeleteAllElements()      ← 刪除舊元素
       │       ├─ generateAllElementsUltra()    ← 為每頁生成元素
       │       │   ├─ addProgressBarUltra()
       │       │   ├─ addPageNumberUltra()
       │       │   ├─ addTitleFootnoteUltra()
       │       │   └─ addTabNavigationUltra()   ← 標籤導航
       │       └─ addSectionElementsUltra()     ← 章節頁專用元素
       │           ├─ addUnifiedSectionBox()     ← Section Box
       │           ├─ addSectionLabel()          ← Section Label
       │           └─ addOutlineToSecondSlide()  ← Outline
       │
       ├─ "📑 更新標籤頁" → runProcessTabs()
       ├─ "📚 更新 SECTION Header" → runProcessSectionBoxes()
       └─ ...
```
