# Lizard Google Slides Module

A Google Apps Script project that enhances Google Slides with automated formatting, styling, and content management features. This module provides a custom menu with various tools to improve slide design consistency and streamline presentation creation.

![image](https://github.com/user-attachments/assets/b9930702-381d-4bc6-9458-8385bea9d7a7)

## Project Structure

The codebase is organized into the following directories:

- **src/util** - Utility functions for manipulating slides (shapes, styles, etc.)
- **src/batch** - Batch processing modules for applying changes to multiple slides
- **src/components** - HTML components for the sidebar interface
- **scripts** - Core JavaScript files including configuration and HTML service utilities

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
   git clone https://github.com/htlin222/lizard-gslide-module.git
   cd lizard-gslide-module
   ```

4. Create a new Google Slides project from within the cloned directory:

   ```bash
   clasp create --type slides --title "My Styled Presentation"
   ```

5. Overwrite the appscript.json file with the example configuration:

   ```bash
   cp appsscript.example.json appsscript.json
   ```

6. Push the code to your new project:

   ```bash
   clasp push
   ```

7. Open the Google Slides presentation directly:

   ```bash
   clasp open-container
   ```

   This command opens the container (Google Slides presentation) directly in your browser

8. In the Google Slides presentation, go to Extensions > Apps Script to enable the script

   ![image](https://github.com/user-attachments/assets/fe774867-c791-4e6f-a948-b54dfc34c693)

9. The `onOpen` function might not run automatically. If you don't see the menu:
   - In the Apps Script editor, open the `src/config.js` file
   - Find the `onOpen()` function
   - Click the Run button (▶️) to manually create the menu

![image](https://github.com/user-attachments/assets/b068f4c2-8c36-406d-9753-c89787370fe3)

11. You should now see the "🛠 工具選單" (Tools Menu) in your menu bar

![image](https://github.com/user-attachments/assets/dab56975-9be1-494b-9509-1347d9dfa9d3)

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
  - 🔰 加上 badge (Add Badge)
  - 🍡 貼上在同一處 (Paste in Place)

- **🖖 新增** (Add Content)
  - 👆 取得前一頁的標題 (Get Previous Title)
  - 👇 標題加到新的下頁 (Add Title to New Slide)
  - 🎨 套用主題 (Apply Theme)

## Configuration

You can customize the module by modifying the variables in `src/config.js`:

```javascript
var main_color = "#3D6869"; // Main theme color
var main_font_family = "Source Sans Pro"; // Font family
var water_mark_text = "ⓒ Hsieh-Ting Lin"; // Watermark text
var label_font_size = 14; // Font size for labels
```

The source presentation template ID can also be changed:

```javascript
const sourcePresentationId = "1qAZzq-uo5blLH1nqp9rbrGDlzz_Aj8eIp0XjDdmI220";
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

Hsieh-Ting Lin M.D.
