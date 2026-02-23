# 物件定位方法 (Object Positioning & Coordinate System)

## 概述

要在 Google Slides 中精準定位物件，你必須理解投影片的座標系統、尺寸單位、以及兩套不同的 API（SlidesApp 高階 API 和 Slides REST API）各自的定位方式。本文件整理了專案中所有與定位相關的程式碼模式。

## 投影片座標系統

### 標準投影片尺寸

Google Slides 的標準寬螢幕 (16:9) 投影片尺寸為：

```
寬度 (Width)  = 720 pt
高度 (Height) = 405 pt
```

本專案取得尺寸的方式：

```javascript
// src/batch/cache_manager.js:58-59
const width = presentation.getPageWidth();    // 720
const height = presentation.getPageHeight();  // 405
```

其他常見比例的投影片尺寸：
- 4:3 → 720 pt × 540 pt
- 16:10 → 720 pt × 450 pt

### 座標原點與方向

```
(0, 0) ─────────────────────────── (720, 0)
  │                                      │
  │              投影片                   │
  │                                      │
  │    (x, y) = (translateX, translateY) │
  │       ┌──────────┐                   │
  │       │   形狀   │ height             │
  │       └──────────┘                   │
  │          width                        │
(0, 405) ────────────────────────── (720, 405)
```

- **原點** = 左上角 (0, 0)
- **X 軸** → 向右為正
- **Y 軸** → 向下為正
- **單位** = Points (PT)

## 尺寸單位

### Points (PT)

本專案**全程使用 Points (PT)** 作為基本單位。在 Google Slides 中：

- 1 PT = 1/72 inch
- 1 inch = 72 PT
- 1 cm ≈ 28.35 PT

```javascript
// src/batch/element_generators.js — 所有尺寸都用 PT
size: {
    width: { magnitude: 70, unit: 'PT' },
    height: { magnitude: 30, unit: 'PT' }
}
```

### EMU (English Metric Units)

Google Slides REST API 底層使用 EMU 作為內部單位，但本專案在 API 請求中直接指定 `unit: 'PT'`，由 API 自動轉換：

```
1 PT = 12,700 EMU
1 inch = 914,400 EMU
1 cm = 360,000 EMU
```

你**不需要**手動計算 EMU，只要在 API 請求中指定 `unit: 'PT'` 即可。

## 兩套定位方式

### 方式一：SlidesApp API（高階 API）

用於單個形狀的即時操作，使用 `setLeft()`, `setTop()`, `setWidth()`, `setHeight()`：

```javascript
// src/util/graph/shape_creator.js:92-98 — 插入形狀
const childShape = slide.insertShape(
    parentShape.getShapeType(),
    childLeft,         // X 位置 (PT)
    childTop,          // Y 位置 (PT)
    childWidth,        // 寬度 (PT)
    childHeight        // 高度 (PT)
);

// 明確設定位置（確保精準度）
childShape.setLeft(childLeft);
childShape.setTop(childTop);
childShape.setWidth(childWidth);
childShape.setHeight(childHeight);
```

#### 讀取物件位置

```javascript
// src/util/add_arrow.js:32-55 — logSelectedItemGeometry()
const element = pageElements[0];
const transform = element.getTransform();

const x = transform.getTranslateX();    // X 位置
const y = transform.getTranslateY();    // Y 位置
const width = element.getWidth();        // 寬度
const height = element.getHeight();      // 高度
```

#### 也可以用 getLeft/getTop

```javascript
// src/util/graph/shape_creator.js:61-65
const parentLeft = parentShape.getLeft();
const parentTop = parentShape.getTop();
const parentWidth = parentShape.getWidth();
const parentHeight = parentShape.getHeight();
const parentRotation = parentShape.getRotation();
```

### 方式二：Slides REST API（Batch Update）

用於批次操作，使用 `transform` 物件：

```javascript
// src/batch/element_generators.js — Batch API 的定位方式
elementProperties: {
    pageObjectId: slideId,           // 指定在哪一頁
    size: {
        width: { magnitude: 70, unit: 'PT' },
        height: { magnitude: 30, unit: 'PT' }
    },
    transform: {
        scaleX: 1,                   // X 縮放（1 = 不縮放）
        scaleY: 1,                   // Y 縮放（1 = 不縮放）
        translateX: 650,             // X 位移 (PT)
        translateY: 370,             // Y 位移 (PT)
        unit: 'PT'                   // 單位
    }
}
```

#### Transform 矩陣

`transform` 實際上是一個 2D 仿射變換矩陣：

```
| scaleX  shearX  translateX |
| shearY  scaleY  translateY |
| 0       0       1          |
```

本專案中預定義了兩個常用的 transform：

```javascript
// src/batch/cache_manager.js:31-33
transforms: {
    identity: { scaleX: 1, scaleY: 1, unit: 'PT' },
    rotation90: { scaleX: 0, shearX: -1, shearY: 1, scaleY: 0, unit: 'PT' }
}
```

- **identity** — 不旋轉、不變形，只需設定 `translateX/Y` 來定位
- **rotation90** — 旋轉 90 度（用於側邊標題腳註）

使用方式：

```javascript
// 不旋轉的定位
transform: { ...cache.transforms.identity, translateX: 650, translateY: 370 }

// 旋轉 90 度的定位（如側邊標題）
transform: { ...cache.transforms.rotation90, translateX: slideCache.width, translateY: boxY }
```

## 實際定位範例集

### 範例 1：進度條（底部全寬）

```javascript
// src/batch/element_generators.js:42-49
// 位置：投影片底部，全寬
size: {
    height: { magnitude: 5, unit: 'PT' },        // 高 5pt
    width: { magnitude: 720, unit: 'PT' }          // 全寬 720pt
},
transform: {
    scaleX: 1, scaleY: 1, unit: 'PT',
    translateX: 0,                                  // 最左邊
    translateY: 400                                  // 底部 (405 - 5)
}
```

### 範例 2：頁碼（右下角）

```javascript
// src/batch/element_generators.js:93-97
// 位置：右下角
size: { width: { magnitude: 70, unit: 'PT' }, height: { magnitude: 30, unit: 'PT' } },
transform: {
    scaleX: 1, scaleY: 1, unit: 'PT',
    translateX: 650,                                // 右側 (720 - 70)
    translateY: 370                                  // 底部 (405 - 30 - 5 for progress bar)
}
```

### 範例 3：標籤導航列（頂部置中）

```javascript
// src/batch/element_generators.js:193-197
// 位置：最頂部，全寬
size: {
    height: { magnitude: 14, unit: 'PT' },
    width: { magnitude: 720, unit: 'PT' }
},
transform: {
    scaleX: 1, scaleY: 1, unit: 'PT',
    translateX: 0,
    translateY: 0                                    // 最頂部
}
```

### 範例 4：章節文字框（垂直置中）

```javascript
// src/batch/section_elements.js:25-27
// 計算垂直置中
const slideHeight = 405;
const boxHeight = 300;
const y = (slideHeight - boxHeight) / 2;             // (405 - 300) / 2 = 52.5

// 位置：偏右、垂直置中
transform: { scaleX: 1, scaleY: 1, unit: 'PT', translateX: 200, translateY: y }
```

### 範例 5：浮水印（45 度旋轉置中）

```javascript
// src/batch/toggle_watermark.js:8-15
const slideWidth  = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();
const wmWidth  = 500;
const wmHeight = 100;
const dx = slideWidth / 2;                           // 水平中心
const dy = slideHeight / 2 - 2 * wmHeight;          // 垂直偏上
const cos45 = Math.SQRT1_2;                          // ≈ 0.707
const sin45 = Math.SQRT1_2;

transform: {
    scaleX:     cos45,                               // cos(45°)
    shearX:    -sin45,                               // -sin(45°)
    shearY:     sin45,                               // sin(45°)
    scaleY:     cos45,                               // cos(45°)
    translateX: dx,
    translateY: dy,
    unit:      'PT'
}
```

### 範例 6：側邊標題腳註（旋轉 90 度）

```javascript
// src/batch/element_generators.js:137-147
const boxWidth = 360;
const boxY = (slideCache.height - boxWidth) / 2;     // 垂直置中（注意：旋轉後寬變高）

transform: {
    // 90 度旋轉矩陣
    scaleX: 0, shearX: -1, shearY: 1, scaleY: 0,
    translateX: slideCache.width,                     // 投影片右邊緣
    translateY: boxY,
    unit: 'PT'
}
```

### 範例 7：網格線（用 insertLine）

```javascript
// src/util/toggle_grids.js:21-29
var width = 720, height = 405;
var coords = [
    // [x1, y1, x2, y2]
    [180, 0,   180, height],      // 垂直線 1 (1/4 位置)
    [360, 0,   360, height],      // 垂直線 2 (1/2 位置)
    [540, 0,   540, height],      // 垂直線 3 (3/4 位置)
    [0,   135, width, 135],       // 水平線 1 (1/3 位置)
    [0,   270, width, 270]        // 水平線 2 (2/3 位置)
];

// 使用 insertLine 建立線條
var line = slide.insertLine(
    SlidesApp.LineCategory.STRAIGHT,
    c[0], c[1],    // 起點 (x1, y1)
    c[2], c[3]     // 終點 (x2, y2)
);
```

### 範例 8：在父形狀中建立子形狀（網格計算）

```javascript
// src/util/graph/shape_creator.js:61-98
const parentLeft = parentShape.getLeft();
const parentTop = parentShape.getTop();
const parentWidth = parentShape.getWidth();
const parentHeight = parentShape.getHeight();

// 計算可用空間
const availableWidth = parentWidth - padding * 2;
const availableHeight = parentHeight - paddingTop - padding;

// 計算子形狀尺寸
const childWidth = (availableWidth - gap * (columns - 1)) / columns;
const childHeight = (availableHeight - gap * (rows - 1)) / rows;

// 計算子形狀位置
for (let row = 0; row < rows; row++) {
    for (let col = 0; col < columns; col++) {
        const childLeft = parentLeft + padding + col * (childWidth + gap);
        const childTop = parentTop + paddingTop + row * (childHeight + gap);

        const childShape = slide.insertShape(
            parentShape.getShapeType(),
            childLeft, childTop, childWidth, childHeight
        );
    }
}
```

### 範例 9：相鄰元素間的居中定位

```javascript
// src/util/average_padding.js:87-114
// 在左右鄰居之間居中
if (neighbors.left && neighbors.right) {
    const leftEdge = neighbors.left.right;       // 左鄰居的右邊緣
    const rightEdge = neighbors.right.left;      // 右鄰居的左邊緣
    const availableWidth = rightEdge - leftEdge;
    newX = leftEdge + (availableWidth - selectedWidth) / 2;
}

// 在上下鄰居之間居中
if (neighbors.top && neighbors.bottom) {
    const topEdge = neighbors.top.bottom;
    const bottomEdge = neighbors.bottom.top;
    const availableHeight = bottomEdge - topEdge;
    newY = topEdge + (availableHeight - selectedHeight) / 2;
}
```

## 重要的定位公式

### 置中公式

```javascript
// 水平置中
x = (slideWidth - objectWidth) / 2

// 垂直置中
y = (slideHeight - objectHeight) / 2

// 在特定區域內置中
x = regionLeft + (regionWidth - objectWidth) / 2
y = regionTop + (regionHeight - objectHeight) / 2
```

### 靠右對齊

```javascript
x = slideWidth - objectWidth - margin
// 例如：頁碼靠右
translateX: 720 - 70 = 650
```

### 靠底對齊

```javascript
y = slideHeight - objectHeight - margin
// 例如：進度條靠底
translateY: 405 - 5 = 400
```

### 網格排列

```javascript
// N 個等寬物件在可用空間中排列
const objectWidth = (availableWidth - gap * (N - 1)) / N;
const objectX = startX + i * (objectWidth + gap);
```

### 旋轉矩陣

```javascript
// 旋轉 θ 度
const rad = θ * Math.PI / 180;
transform: {
    scaleX: Math.cos(rad),
    shearX: -Math.sin(rad),
    shearY: Math.sin(rad),
    scaleY: Math.cos(rad),
    translateX: x,
    translateY: y,
    unit: 'PT'
}
```

## 常用位置參考表

以標準 16:9 投影片 (720 × 405 pt) 為基準：

| 位置描述 | translateX | translateY |
|----------|-----------|-----------|
| 左上角 | 0 | 0 |
| 右上角 | 720 - width | 0 |
| 左下角 | 0 | 405 - height |
| 右下角 | 720 - width | 405 - height |
| 水平置中 | (720 - width) / 2 | - |
| 垂直置中 | - | (405 - height) / 2 |
| 完全置中 | (720 - width) / 2 | (405 - height) / 2 |
| 頂部標籤列 | 0 | 0 |
| 底部進度條 | 0 | 400 |
| 右下頁碼 | 650 | 370 |

## SlidesApp API vs Slides REST API 對比

| 操作 | SlidesApp API | Slides REST API (Batch) |
|------|--------------|------------------------|
| 建立形狀 | `slide.insertShape(type, x, y, w, h)` | `createShape` request |
| 設定位置 | `shape.setLeft(x); shape.setTop(y)` | `transform.translateX/Y` |
| 設定大小 | `shape.setWidth(w); shape.setHeight(h)` | `size.width/height` |
| 設定旋轉 | `shape.setRotation(degrees)` | transform 矩陣 |
| 讀取位置 | `shape.getLeft()` / `transform.getTranslateX()` | N/A（需先讀取） |
| 適用場景 | 即時操作、互動式 | 批次操作、高效能 |
| 效能 | 每次呼叫都是獨立 API 請求 | 多個操作合併為 1 個 API 請求 |

## 形狀標題 (Title / Alt Text) 作為識別標記

專案中大量使用形狀的 `title` 屬性作為程式化識別標記：

```javascript
// SlidesApp API
shape.setTitle('WATERMARK');
shape.setTitle('NUMBER');
shape.setTitle('PARENT');
shape.setTitle('CHILD');
shape.setTitle('PREVIOUS_TITLE');
shape.setTitle('GRID_1');

// Slides REST API (Batch)
requests.push({
    updatePageElementAltText: {
        objectId: wmId,
        title: 'WATERMARK'
    }
});
```

這讓系統可以在後續操作中識別和管理特定元素（如刪除所有 `WATERMARK`、找到所有 `NUMBER` 元素等）。
