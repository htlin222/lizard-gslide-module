# 如何操作不同的模板 (Template Operations)

## 概述

本專案實作了一套**模板管理系統**，讓使用者可以從一個「範本投影片」匯入主題（包括 Master Slide、Layout、顏色方案），並透過設定系統（`PropertiesService`）來持久化自訂的顏色、字型等配置。本文件涵蓋三個層面：

1. **主題匯入** — 從範本投影片複製主題到當前投影片
2. **設定系統** — 持久化儲存和載入配置（顏色、字型、進度條高度等）
3. **樣式系統** — 預定義的 6 種形狀樣式，可一鍵套用

## 關鍵檔案

| 檔案 | 角色 |
|------|------|
| `src/batch/theme.js` | **主題匯入** — `applyThemeToCurrentPresentation()` 和 `applyTheme()` |
| `src/config.js` | **設定管理** — 全域變數、PropertiesService、選單系統 |
| `src/util/default_style.js` | **樣式定義** — 6 種預定義樣式 |
| `src/util/html_service_utils.js` | **Sidebar 服務** — HTML 模組化工具 |
| `src/components/config-sidebar/` | **設定面板** — 使用者介面 |

## 第一部分：主題匯入

### 核心概念

Google Slides 的「主題」包含 Master Slide（母版投影片）和 Layout（版面配置）。當你從另一個投影片 `appendSlide()` 時，Google Slides 會**自動複製**該投影片對應的 Master Slide 和 Layout 到目標投影片。

### 範本投影片 ID

```javascript
// src/config.js:12
const sourcePresentationId = "1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220";
```

這是一個固定的 Google Slides 檔案 ID，存放了團隊的標準主題模板。

### 主題匯入流程

```javascript
// src/batch/theme.js:10-30 — applyThemeToCurrentPresentation()
function applyThemeToCurrentPresentation() {
    // 1. 取得範本投影片 ID
    Logger.log('Source presentation ID: ' + sourcePresentationId);

    // 2. 取得當前投影片 ID（優先從 PropertiesService 讀取）
    const currentPresentationId =
        PropertiesService.getScriptProperties().getProperty('presentationId') ||
        SlidesApp.getActivePresentation().getId();

    // 3. 開啟兩個投影片
    const sourcePresentation = SlidesApp.openById(sourcePresentationId);
    const currentPresentation = SlidesApp.openById(currentPresentationId);

    // 4. 執行主題套用
    applyTheme(sourcePresentation, currentPresentation);
}
```

### applyTheme() 詳細步驟

```javascript
// src/batch/theme.js:37-119 — applyTheme()
function applyTheme(sourcePresentation, targetPresentation) {
    // Step 1: 從範本複製第一張和最後一張投影片
    // → 這會自動帶入 Master Slide 和 Layout
    const sourceSlides = sourcePresentation.getSlides();
    const copiedFirstSlide = targetPresentation.appendSlide(sourceSlides[0]);
    const copiedLastSlide = targetPresentation.appendSlide(sourceSlides[sourceSlides.length - 1]);

    // Step 2: 更新封面標題為當前投影片的檔案名稱
    const presentationName = targetPresentation.getName();
    const shapes = copiedFirstSlide.getShapes();
    for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
            const placeholder = shape.getPlaceholderType();
            if (placeholder === SlidesApp.PlaceholderType.TITLE || i === 0) {
                shape.getText().setText(presentationName);
                break;
            }
        }
    }

    // Step 3: 刪除原來的第一張投影片
    const allSlides = targetPresentation.getSlides();
    if (allSlides.length > 2) {
        allSlides[0].remove();  // 刪除原始空白頁
    }

    // 現在投影片有了新的主題，所有新建的投影片都會使用這個主題
}
```

### 自動觸發

當偵測到新投影片（只有 0-1 頁）時，自動匯入主題：

```javascript
// src/config.js:37-44
function onOpen() {
    loadSavedConfiguration();
    createCustomMenu();

    const slides = presentation.getSlides();
    if (slides.length <= 1) {
        applyThemeToCurrentPresentation();   // ← 自動套用主題
        Logger.log('New presentation detected - theme automatically applied');
    }
}
```

## 第二部分：設定系統 (PropertiesService)

### 持久化儲存

Google Apps Script 的 `PropertiesService` 提供了三種儲存範圍：

| 類型 | 方法 | 範圍 |
|------|------|------|
| User Properties | `PropertiesService.getUserProperties()` | 每個使用者獨立 |
| Script Properties | `PropertiesService.getScriptProperties()` | 整個腳本共享 |
| Document Properties | `PropertiesService.getDocumentProperties()` | 每個文件獨立 |

本專案使用 **User Properties** 來儲存個人配置：

```javascript
// src/config.js:14-21
const CONFIG_KEYS = {
    MAIN_COLOR: 'main_color',
    FONT_FAMILY: 'main_font_family',
    WATERMARK_TEXT: 'water_mark_text',
    FONT_SIZE: 'label_font_size',
    PROGRESS_BAR_HEIGHT: 'progress_bar_height'
};
```

### 儲存設定

```javascript
// src/config.js:320-339 — saveConfigValues()
function saveConfigValues(config) {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperties({
        [CONFIG_KEYS.MAIN_COLOR]: config.mainColor,
        [CONFIG_KEYS.FONT_FAMILY]: config.fontFamily,
        [CONFIG_KEYS.WATERMARK_TEXT]: config.watermarkText,
        [CONFIG_KEYS.FONT_SIZE]: config.fontSize,
        [CONFIG_KEYS.PROGRESS_BAR_HEIGHT]: config.progressBarHeight
    });

    // 同時更新全域變數
    main_color = config.mainColor;
    main_font_family = config.fontFamily;
    water_mark_text = config.watermarkText;
    label_font_size = parseInt(config.fontSize, 10);
    progressBarHeight = parseInt(config.progressBarHeight, 10);
}
```

### 載入設定

```javascript
// src/config.js:360-379 — loadSavedConfiguration()
function loadSavedConfiguration() {
    const userProperties = PropertiesService.getUserProperties();
    const savedMainColor = userProperties.getProperty(CONFIG_KEYS.MAIN_COLOR);
    const savedFontFamily = userProperties.getProperty(CONFIG_KEYS.FONT_FAMILY);
    // ...

    // 只在有儲存值時覆蓋預設值
    if (savedMainColor) main_color = savedMainColor;
    if (savedFontFamily) main_font_family = savedFontFamily;
    // ...
}
```

### 儲存並立即套用

```javascript
// src/config.js:345-354 — saveAndApplyConfig()
function saveAndApplyConfig(config) {
    saveConfigValues(config);    // 儲存設定
    runAllFunctions();           // 重新執行所有批次功能（套用新設定）
}
```

### 讀取設定（供 Sidebar 使用）

```javascript
// src/config.js:243-274 — getConfigValues()
function getConfigValues() {
    const userProperties = PropertiesService.getUserProperties();
    return {
        mainColor: userProperties.getProperty(CONFIG_KEYS.MAIN_COLOR) || main_color,
        baseColor: base_color,
        textColor: text_color,
        sub1Color: sub1_color,
        accentColor: accent_color,
        fontFamily: userProperties.getProperty(CONFIG_KEYS.FONT_FAMILY) || main_font_family,
        watermarkText: userProperties.getProperty(CONFIG_KEYS.WATERMARK_TEXT) || water_mark_text,
        fontSize: userProperties.getProperty(CONFIG_KEYS.FONT_SIZE) || label_font_size,
        progressBarHeight: userProperties.getProperty(CONFIG_KEYS.PROGRESS_BAR_HEIGHT) || progressBarHeight,
        availableFonts: getAvailableFonts()
    };
}
```

## 第三部分：樣式系統

### 6 種預定義樣式

```javascript
// src/util/default_style.js:12-67 — STYLE_DEFINITIONS
const STYLE_DEFINITIONS = {
    1: { name: 'Style 1', fillColor: 'base_color',   borderColor: 'main_color',   textColor: 'main_color' },
    2: { name: 'Style 2', fillColor: 'main_color',   borderColor: 'main_color',   textColor: 'base_color' },
    3: { name: 'Style 3', fillColor: 'sub1_color',   borderColor: 'main_color',   textColor: 'main_color' },
    4: { name: 'Style 4', fillColor: 'base_color',   borderColor: 'accent_color', textColor: 'accent_color' },
    5: { name: 'Style 5', fillColor: 'accent_color', borderColor: 'accent_color', textColor: 'base_color' },
    6: { name: 'Style 6', fillColor: 'base_color',   borderColor: 'base_color',   textColor: 'main_color' }
};
```

視覺上的效果（以預設色 #3D6869 為例）：

```
Style 1: [白底 + 主題色邊框 + 主題色文字]
Style 2: [主題色底 + 主題色邊框 + 白色文字]  ← 最常見的「強調」樣式
Style 3: [淺灰底 + 主題色邊框 + 主題色文字]  ← 「次要」樣式
Style 4: [白底 + 橘色邊框 + 橘色文字]        ← 「提醒」樣式
Style 5: [橘色底 + 橘色邊框 + 白色文字]      ← 「警告」樣式
Style 6: [白底 + 白色邊框 + 主題色文字]      ← 「隱形邊框」樣式
```

### 顏色變數解析

樣式定義中使用的是變數名稱（如 `'main_color'`），實際顏色值在執行時解析：

```javascript
// src/util/default_style.js:92-107 — resolveColorVariable()
function resolveColorVariable(colorVar, config) {
    switch (colorVar) {
        case 'main_color':   return config.mainColor || main_color;     // #3D6869
        case 'base_color':   return config.baseColor || base_color;     // #FFFFFF
        case 'sub1_color':   return config.sub1Color || sub1_color;     // #E7EAE7
        case 'accent_color': return config.accentColor || accent_color; // #f29424
        case 'text_color':   return config.textColor || text_color;     // #333333
        default: return colorVar;
    }
}
```

### 套用樣式到形狀

```javascript
// src/util/default_style.js:113-186 — applyDefaultStyle()
function applyDefaultStyle(styleNumber) {
    const selectedElements = selection.getPageElementRange().getPageElements();
    const styles = getStyleDefinitions();
    const style = styles[styleNumber];

    for (const element of selectedElements) {
        if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
            const shape = element.asShape();

            // 確保形狀有文字（避免錯誤）
            if (!shape.getText().asString()) {
                shape.getText().setText('TEXT_HERE');
            }

            // 套用邊框
            shape.getBorder().setWeight(style.borderWidth);         // 1pt
            shape.getBorder().getLineFill().setSolidFill(style.borderColor);

            // 套用填充色
            shape.getFill().setSolidFill(style.fillColor);

            // 套用文字色
            shape.getText().getTextStyle().setForegroundColor(style.textColor);
        }
    }
}
```

## 第四部分：HTML Sidebar 模組化

### include() 函數

Google Apps Script 不支援 ES6 import，但支援**伺服器端 HTML 包含**：

```javascript
// src/util/html_service_utils.js:8-10
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
```

在 HTML 模板中使用：

```html
<!-- src/components/config-sidebar/index.html -->
<!DOCTYPE html>
<html>
<head>
    <?!= include('src/components/config-sidebar/styles') ?>
</head>
<body>
    <?!= include('src/components/config-sidebar/config-form') ?>
    <?!= include('src/components/config-sidebar/style-buttons') ?>
    <?!= include('src/components/config-sidebar/sidebar-scripts') ?>
</body>
</html>
```

`<?!= ... ?>` 是 Google Apps Script 的**模板語法**，`!` 表示不要 HTML escape 輸出。

### 建立模組化 Sidebar

```javascript
// src/util/html_service_utils.js:31-36
function createConfigSidebar() {
    const template = createModularHtmlTemplate('src/components/config-sidebar/index');
    return template.evaluate()
        .setTitle('Lizard Slides')
        .setWidth(300);
}
```

### Sidebar 與伺服器端通訊

在 HTML 中呼叫伺服器端函數：

```html
<!-- 在 sidebar HTML 中 -->
<script>
    // 讀取設定
    google.script.run
        .withSuccessHandler(function(config) {
            document.getElementById('mainColor').value = config.mainColor;
        })
        .getConfigValues();

    // 儲存並套用設定
    function saveAndApply() {
        var config = {
            mainColor: document.getElementById('mainColor').value,
            fontFamily: document.getElementById('fontFamily').value,
            // ...
        };
        google.script.run
            .withSuccessHandler(function() { alert('Settings applied!'); })
            .saveAndApplyConfig(config);
    }
</script>
```

## 你自己要做：如何建立自己的模板系統

### 範例 1：使用不同的範本投影片

```javascript
// 方法 A：固定 ID
const myTemplateId = "YOUR_TEMPLATE_PRESENTATION_ID";

// 方法 B：從 PropertiesService 讀取（可動態切換）
const templateId = PropertiesService.getUserProperties().getProperty('template_id') || DEFAULT_TEMPLATE_ID;

// 方法 C：從 Google Drive 搜尋特定名稱的投影片
const files = DriveApp.getFilesByName('My Template');
if (files.hasNext()) {
    const templateId = files.next().getId();
}
```

### 範例 2：建立多個主題供切換

```javascript
const THEMES = {
    'lizard': '1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220',
    'ocean':  'OCEAN_TEMPLATE_ID',
    'sunset': 'SUNSET_TEMPLATE_ID'
};

function applyThemeByName(themeName) {
    const templateId = THEMES[themeName];
    const source = SlidesApp.openById(templateId);
    const target = SlidesApp.getActivePresentation();
    applyTheme(source, target);
}
```

### 範例 3：匯出當前投影片為模板

```javascript
function saveAsTemplate() {
    const presentation = SlidesApp.getActivePresentation();
    const templateId = presentation.getId();

    // 儲存為自訂模板
    PropertiesService.getUserProperties().setProperty('custom_template_id', templateId);

    Logger.log('Saved current presentation as template: ' + templateId);
}
```

### 範例 4：最小的設定系統

```javascript
// 全域設定
var CONFIG = {
    mainColor: '#3D6869',
    fontFamily: 'Source Sans Pro',
    progressBarHeight: 5
};

// 載入設定
function loadConfig() {
    const props = PropertiesService.getUserProperties();
    const saved = props.getProperty('slide_config');
    if (saved) {
        Object.assign(CONFIG, JSON.parse(saved));
    }
}

// 儲存設定
function saveConfig(newConfig) {
    Object.assign(CONFIG, newConfig);
    PropertiesService.getUserProperties().setProperty(
        'slide_config',
        JSON.stringify(CONFIG)
    );
}
```

## 全域變數 vs PropertiesService

| 特性 | 全域變數 | PropertiesService |
|------|---------|-------------------|
| 生命週期 | 每次執行重新初始化 | 永久儲存 |
| 跨執行保持 | 否 | 是 |
| 跨使用者共享 | 否 | 取決於類型 |
| 讀取速度 | 即時 | 需要 API 呼叫 |
| 適用場景 | 執行期間的暫存值 | 使用者偏好設定 |

本專案的設計模式是：**PropertiesService 作為持久儲存，全域變數作為執行期間的快取**。每次 `onOpen()` 時從 PropertiesService 載入值到全域變數。

## 選單系統架構

```javascript
// src/config.js:63-147 — createCustomMenu()
// 四個頂層選單：
ui.createMenu('📦 批次處理')      // 批次操作（主題、進度條、標籤等）
ui.createMenu('✨ 加入元素')       // 單頁美化工具
ui.createMenu('🎨 繪圖')          // 圖形操作工具
ui.createMenu('🖖 跨頁功能')       // 跨頁操作（標題鏈、Markdown 等）
```

每個選單項目對應一個全域函數，如：

```javascript
.addItem('🎨 套用蜥蜴主題', 'applyThemeToCurrentPresentation')
.addItem('🔄 更新進度條', 'runUpdateProgressBars')
```

Google Apps Script 的選單系統要求函數名稱必須是**字串形式的全域函數名**，不支援傳入函數引用或閉包。
