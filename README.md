# Lizard Google Slides Module

A Google Apps Script project that enhances Google Slides with automated formatting, styling, and content management features. This module provides a custom menu with various tools to improve slide design consistency and streamline presentation creation.

## Features

### Batch Processing

- **Apply All Updates**: Run all formatting functions at once
- **Progress Bars**: Add dynamic progress indicators at the bottom of slides
- **Tab Lists**: Format and manage tab-style navigation elements
- **Section Boxes**: Create and update section navigation boxes
- **Footer Updates**: Maintain consistent footer elements across slides
- **Watermark Toggle**: Add or remove watermarks from slides

### Single Slide Beautification

- **Date Updates**: Automatically update date elements on the first slide
- **Grid Toggle**: Add or remove grid lines for precise element positioning
- **Badge Creation**: Convert text elements into styled badges
- **Paste in Place**: Duplicate images in the exact same position across slides

### Content Creation

- **Title Propagation**: Copy titles from previous slides
- **Title-Based Slides**: Create new slides with current title
- **Theme Application**: Apply a consistent theme from a template presentation

## Quick Start Guide

1. Install clasp if you haven't already:
   ```bash
   npm install -g @google/clasp
   ```

2. Login to your Google account:
   ```bash
   clasp login
   ```

3. Clone this repository:
   ```bash
   git clone <repository-url>
   cd lizard-gslide-module
   ```

4. Create a new Google Slides project from within the cloned directory:
   ```bash
   clasp create --type slides --title "My Styled Presentation"
   ```

5. Push the code to your new project:
   ```bash
   clasp push
   ```

6. Open the newly created presentation from your Google Drive

7. Open the Apps Script editor (Extensions > Apps Script)

8. Run the `onOpen` function to initialize the custom menu and apply the theme

9. Return to your presentation and you should now see the "🛠 工具選單" (Tools Menu) in your menu bar

## Manual Installation

1. Open your Google Slides presentation
2. Go to Extensions > Apps Script
3. Delete any code in the editor
4. Copy all the files from this repository into the Apps Script editor
5. Save the project
6. Refresh your presentation

## Usage

After installation, a new menu item "🛠 工具選單" will appear in your Google Slides menu bar. If the menu doesn't appear automatically, you can refresh the page or manually run the `showMenuManually()` function from the Apps Script editor.

### Menu Structure

- **🗃️ 批次處理** (Batch Processing)
  - 🛠 同時更新所有 (Apply All Updates)
  - 🔄 更新進度條 (Update Progress Bars)
  - 📑 更新標籤頁 (Update Tabs)
  - 📚 更新章節導覽 (Update Section Boxes)
  - 🦶 更新 Footer (Update Footer)
  - 💧 切換浮水印 (Toggle Watermark)

- **🎨 單頁美化** (Single Slide Beautification)
  - 📅 更新日期 (Update Date)
  - 📏 加上網格 (Add Grid)
  - 🔰 加上badge (Add Badge)
  - 🍡 貼上在同一處 (Paste in Place)

- **🖖 新增** (Add Content)
  - 👆 取得前一頁的標題 (Get Previous Title)
  - 👇 標題加到新的下頁 (Add Title to New Slide)
  - 🎨 套用主題 (Apply Theme)

## Configuration

You can customize the module by modifying the variables in `config.gs`:

```javascript
var main_color = '#3D6869';               // Main theme color
var main_font_family = 'Source Sans Pro';  // Font family
var water_mark_text = 'ⓒ Hsieh-Ting Lin'; // Watermark text
var label_font_size = 14;                 // Font size for labels
```

The source presentation template ID can also be changed:

```javascript
const sourcePresentationId = '1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220';
```

## Requirements

- Google account with access to Google Slides
- Google Slides API enabled in the Apps Script project

## Development

This project uses [clasp](https://github.com/google/clasp) for local development. To set up the development environment:

1. Install clasp: `npm install -g @google/clasp`
2. Login to Google: `clasp login`
3. Clone the project: `clasp clone <script-id>`
4. Make changes locally
5. Push changes: `clasp push`

## License

Copyright © Hsieh-Ting Lin

## Author

Hsieh-Ting Lin