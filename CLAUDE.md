# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project that enhances Google Slides with automated formatting, styling, and content management features. The module provides a custom menu with various tools to improve slide design consistency and streamline presentation creation.

## Development Commands

### Initial Setup

```bash
# Install clasp globally if not already installed
npm install -g @google/clasp

# Login to Google account
clasp login

# Initialize new project (use the provided script)
./init.sh

# Or manually create project
clasp create --type slides --title "Your Presentation Title"
cp appsscript.example.json appsscript.json
clasp push
clasp open-container
```

### Development Workflow

```bash
# Push changes to Google Apps Script
clasp push

# Pull changes from Google Apps Script
clasp pull

# Open the Google Slides presentation
clasp open-container

# Open the Apps Script editor
clasp open
```

### Testing

- No automated testing framework is configured
- Testing is done manually in Google Slides after pushing code
- After making changes, run `clasp push` and test in the Google Slides presentation

## Architecture

### Core Files Structure

- **src/config.js** - Main configuration and menu creation logic
  - Contains `onOpen()` function that creates custom menus
  - Defines global configuration variables (colors, fonts, etc.)
  - Manages configuration persistence via PropertiesService

- **src/util/** - Utility functions for slide manipulation
  - Individual utility files with functions for specific tasks
  - No ES6 imports - functions are globally available in Google Apps Script

- **src/batch/** - Batch processing modules
  - Functions that process multiple slides at once
  - Uses Google Slides API batch update requests for efficiency

- **src/components/** - HTML components for sidebar interface
  - Modular HTML files included via `<?!= include() ?>` syntax
  - Contains configuration forms and style buttons
  - **flowchartSidebar.html** - Interactive flowchart creation interface

- **src/util/flowchart/** - Flowchart and hierarchical shape management
  - **main.js** - Main API functions for flowchart operations
  - **graphIdUtils.js** - Graph ID parsing, generation, and management
  - **childCreationUtils.js** - Child shape creation with positioning and styling
  - **siblingCreationUtils.js** - Sibling shape creation with layout consistency
  - **index.js** - Function exports and documentation

### Key Architecture Patterns

1. **Global Functions**: Google Apps Script doesn't support ES6 modules, so all functions are global
2. **Batch API Updates**: Uses `runRequestProcessors()` pattern to collect multiple API requests and send them as a batch
3. **HTML Service**: Uses server-side HTML templates with `<?!= include() ?>` for modular components
4. **Configuration Management**: Uses PropertiesService for persistent configuration storage
5. **Graph ID System**: Uses shape title (alt text) to store hierarchical graph IDs for flowchart management
6. **Flowchart Architecture**: Supports both LR (Left-Right) and TD (Top-Down) layout patterns

### Menu System

Three main menu categories are created in `src/config.js`:

- **🗃 批次處理 (Batch Processing)** - Functions that process multiple slides
- **🎨 加入元素 (Add Elements)** - Single slide beautification tools
- **🖖 跨頁功能 (Cross-page Functions)** - Functions that work across multiple slides

### Configuration System

Configuration is managed through:

- Global variables in `src/config.js` (defaults)
- PropertiesService for user-specific persistent settings
- Sidebar interface for real-time configuration updates

### Important Functions

- `onOpen()` - Automatically creates menus when presentation opens
- `runRequestProcessors(...)` - Batches multiple API requests for efficiency
- `createCustomMenu()` - Creates the custom menu structure
- `applyThemeToCurrentPresentation()` - Applies theme from template presentation

### Flowchart System Functions

- `createChildTop/Right/Bottom/Left()` - Creates child shapes in specified direction
- `createChildTopWithText()` - Creates child shapes with custom text content
- `createSiblingShape()` - Creates sibling shapes with proper positioning
- `showSelectedShapeGraphId()` - Displays Graph ID information for debugging
- `parseGraphId()` - Parses Graph ID format: `graph[parent](layout)[current][children]`
- `generateGraphId()` - Generates hierarchical Graph IDs with layout support
- `getShapeGraphId()` / `setShapeGraphId()` - Graph ID management via shape titles

## Development Notes

- This is a Google Apps Script project, not a Node.js project
- No package.json or npm dependencies
- Uses Google Slides API v1 (enabled in appsscript.json)
- HTML templates use server-side includes, not client-side frameworks
- All code runs in Google's V8 runtime (specified in appsscript.json)

## Configuration Variables

Key variables in `src/config.js`:

```javascript
var main_color = "#3D6869"; // Main theme color
var main_font_family = "Source Sans Pro"; // Font family
var water_mark_text = "ⓒ Hsieh-Ting Lin"; // Watermark text
var label_font_size = 14; // Font size for labels
const sourcePresentationId = "1qAZzq-..."; // Template presentation ID
```

## Google Apps Script Specifics

- Files must be .js or .gs extensions (both work the same)
- Uses HtmlService for UI components
- PropertiesService for data persistence
- SlidesApp and Slides API for slide manipulation
- No require() or import statements - everything is global

## Flowchart Features

### Graph ID System

- Stores hierarchical information in shape titles (alt text) instead of visible text
- Format: `graph[parent](layout)[current][children]`
- Examples:
  - Root: `graph[](TD)[A1][]`
  - Child: `graph[A1](TD)[B1][]`
  - Parent with children: `graph[A1](LR)[B1][C1,C2]`

### Supported Layouts

- **LR (Left-Right)**: Parent connects to children horizontally
- **TD (Top-Down)**: Parent connects to children vertically
- Layout consistency maintained across sibling relationships

### Child Creation Modes

1. **By Count**: Specify number of children to create (empty text)
2. **By Text**: Multi-line text input, each line becomes text for one child shape

### UI Components

- Collapsible Line Settings (default folded)
- Tab-based child creation interface
- Real-time Graph ID inspector
- Persistent settings via localStorage
