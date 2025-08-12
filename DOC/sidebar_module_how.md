# Modular Sidebar Guide for Lizard Slides

This guide explains how the modular sidebar structure works in Lizard Slides and how to maintain, extend, or modify it. The project now supports multiple modular sidebars following a consistent pattern.

## Overview

All sidebars in Lizard Slides have been modularized into smaller components for better maintainability, organization, and reusability. This approach separates concerns (HTML, CSS, JavaScript) and makes the codebase easier to navigate and update. The modular system supports multiple specialized sidebars with shared infrastructure.

## File Structure

```
lizard-gslide-module/
├── src/
│   ├── util/
│   │   └── html_service_utils.js       # Server-side utilities for HTML inclusion
│   └── components/
│       ├── sidebar.html                # Main configuration sidebar container
│       ├── flowchartSidebar.html       # Flowchart sidebar container
│       ├── markdown-sidebar.html       # Markdown sidebar container
│       ├── styles.html                 # Shared CSS styles (main sidebar)
│       ├── config-form.html            # Configuration form inputs
│       ├── style-buttons.html          # Style application buttons
│       ├── sidebar-scripts.html        # JavaScript functionality (main sidebar)
│       └── flowchart/                  # Flowchart sidebar components
│           ├── styles.html             # Flowchart-specific CSS styles
│           ├── line-settings.html      # Line configuration controls
│           ├── shape-connections.html  # Shape connection tools
│           ├── child-creation.html     # Child shape creation interface
│           ├── background-elements.html # Background rectangle tools
│           ├── stage-bar.html          # Stage bar creation tools
│           ├── graph-inspector.html    # Graph ID inspection tools
│           └── scripts.html            # Flowchart-specific JavaScript
```

## How It Works

The modular sidebar system uses Google Apps Script's HTML templating system with a custom `include()` function to load components. Here's the process:

1. **Main Container Files**: Each sidebar type has a main HTML file that serves as a container and includes all relevant components
2. **Shared Infrastructure**: The `html_service_utils.js` file provides server-side functions to handle HTML inclusion for all sidebars
3. **Component Organization**: Components are organized in subdirectories (e.g., `flowchart/`) for specialized sidebars, with shared components at the root level
4. **Consistent Pattern**: All sidebars follow the same creation and loading pattern for consistency

## Server-Side Setup

The `html_service_utils.js` file contains the core infrastructure functions:

1. **`include(filename)`**: Loads an HTML file and returns its content
2. **`createModularHtmlTemplate(filename)`**: Creates an HTML template with the include function attached
3. **Sidebar Creation Functions**:
   - `createConfigSidebar()`: Main configuration sidebar
   - `createMarkdownSidebar()`: Markdown conversion sidebar
   - `createFlowchartSidebar()`: Flowchart tools sidebar

### Sidebar Usage Pattern

All sidebars follow this consistent pattern:

```javascript
// Main configuration sidebar
function showConfigurationSidebar() {
  const sidebar = createConfigSidebar();
  SlidesApp.getUi().showSidebar(sidebar);
}

// Flowchart sidebar
function showFlowchartSidebar() {
  const sidebar = createFlowchartSidebar();
  SlidesApp.getUi().showSidebar(sidebar);
}

// Markdown sidebar
function showMarkdownSidebar() {
  const sidebar = createMarkdownSidebar();
  SlidesApp.getUi().showSidebar(sidebar);
}
```

## Component Structure

### 1. Main Container Files

Each sidebar type has a main HTML container that defines structure and includes components:

- **`src/components/sidebar.html`**: Main configuration sidebar
- **`src/components/flowchartSidebar.html`**: Flowchart tools sidebar
- **`src/components/markdown-sidebar.html`**: Markdown conversion sidebar

All use the include syntax:

```html
<?!= include('src/components/component-name'); ?>
```

### 2. Shared Components

**Styles** (`src/components/styles.html`):

- CSS styles for the main configuration sidebar

**Configuration Form** (`src/components/config-form.html`):

- Form inputs for sidebar configuration

**Style Buttons** (`src/components/style-buttons.html`):

- Style application buttons

**JavaScript** (`src/components/sidebar-scripts.html`):

- JavaScript functionality for main sidebar

### 3. Specialized Component Directories

**Flowchart Components** (`src/components/flowchart/`):

- **`styles.html`**: Flowchart-specific CSS styles
- **`line-settings.html`**: Line and arrow configuration
- **`shape-connections.html`**: Shape connection tools with quadrant-based connections
- **`child-creation.html`**: Child shape creation with count/text modes
- **`background-elements.html`**: Background rectangle creation
- **`stage-bar.html`**: Stage bar creation tools
- **`graph-inspector.html`**: Graph ID inspection and management
- **`scripts.html`**: All flowchart-specific JavaScript functionality

## How to Make Changes

### Adding a New Component

**For Main Sidebar**:

1. Create a new HTML file in `src/components/` (e.g., `new-feature.html`)
2. Add the component's HTML, CSS, or JavaScript content
3. Include in `src/components/sidebar.html`:
   ```html
   <?!= include('src/components/new-feature'); ?>
   ```

**For Specialized Sidebars** (e.g., Flowchart):

1. Create a new HTML file in the specialized directory (e.g., `src/components/flowchart/new-tool.html`)
2. Add the component content
3. Include in the main sidebar file:
   ```html
   <?!= include('src/components/flowchart/new-tool'); ?>
   ```

### Modifying an Existing Component

1. Identify the component location:
   - Main sidebar: `src/components/`
   - Flowchart sidebar: `src/components/flowchart/`
   - Other specialized sidebars: respective subdirectories
2. Make your changes
3. Save the file

### Adding New JavaScript Functions

**For Main Sidebar**:

1. Open `src/components/sidebar-scripts.html`
2. Add your new functions
3. Update initialization if needed

**For Flowchart Sidebar**:

1. Open `src/components/flowchart/scripts.html`
2. Add your new functions
3. Update event listeners in the `window.addEventListener("load")` function

### Adding New Styles

**For Main Sidebar**:

1. Open `src/components/styles.html`
2. Add your new CSS rules

**For Flowchart Sidebar**:

1. Open `src/components/flowchart/styles.html`
2. Add your flowchart-specific CSS rules

## Best Practices

1. **Keep Components Focused**: Each component should have a single responsibility
2. **Minimize Dependencies**: Make components as independent as possible
3. **Follow Consistent Patterns**: All sidebars use the same creation and loading pattern
4. **Organize by Purpose**: Use subdirectories for specialized sidebar components
5. **Use Consistent Naming**: Follow established naming conventions:
   - Main containers: `sidebar.html`, `flowchartSidebar.html`
   - Component directories: `flowchart/`, `markdown/`
   - Creation functions: `createConfigSidebar()`, `createFlowchartSidebar()`
6. **Comment Your Code**: Add comments to explain complex logic
7. **Test After Changes**: Always test sidebars after making changes
8. **Shared vs. Specialized**: Use shared components for common functionality, specialized components for unique features

## Troubleshooting

### Component Not Loading

If a component isn't loading:

1. **Check include paths**:
   - Main sidebar: `<?!= include('src/components/component-name'); ?>`
   - Flowchart sidebar: `<?!= include('src/components/flowchart/component-name'); ?>`
2. **Verify file locations**:
   - Check files exist in correct `src/components/` subdirectory
   - Ensure proper directory structure
3. **Syntax errors**: Check for HTML/CSS/JS syntax errors in component files
4. **Template usage**: Ensure using `createModularHtmlTemplate()` not `createHtmlOutputFromFile()`

### JavaScript Errors

If you encounter JavaScript errors:

1. Check the browser console for error messages
2. Verify that all required DOM elements exist before accessing them
3. Ensure event listeners are attached after the DOM is fully loaded

### Server-Side Errors

If you encounter server-side errors:

1. Check that `html_service_utils.js` is properly included in your project
2. Verify that all HTML files are properly named and located
3. Check Apps Script logs for error messages

## Example: Adding a New Specialized Sidebar

Let's say you want to add a new "Chart Tools" sidebar:

### 1. Create the Infrastructure

**Add creation function to `src/util/html_service_utils.js`**:

```javascript
function createChartSidebar() {
  const template = createModularHtmlTemplate("src/components/chartSidebar");
  return template.evaluate().setTitle("Chart Tools").setWidth(300);
}
```

### 2. Create the Main Container

**Create `src/components/chartSidebar.html`**:

```html
<!doctype html>
<html>
  <head>
    <base target="_top" />
    <?!= include('src/components/chart/styles'); ?>
  </head>
  <body>
    <?!= include('src/components/chart/chart-types'); ?> <?!=
    include('src/components/chart/data-input'); ?> <?!=
    include('src/components/chart/styling-options'); ?> <?!=
    include('src/components/chart/scripts'); ?>
  </body>
</html>
```

### 3. Create Specialized Components

**Create component directory**: `src/components/chart/`

**Create components**:

- `styles.html`: Chart-specific CSS
- `chart-types.html`: Chart type selection interface
- `data-input.html`: Data input forms
- `styling-options.html`: Chart styling controls
- `scripts.html`: Chart-specific JavaScript

### 4. Create the Show Function

**In your chart utility file**:

```javascript
function showChartSidebar() {
  try {
    // Use the modular sidebar approach
    const sidebar = createChartSidebar();
    SlidesApp.getUi().showSidebar(sidebar);
  } catch (e) {
    console.error(`Error showing chart sidebar: ${e.message}`);
    SlidesApp.getUi().alert(
      "Error",
      `Could not open chart sidebar: ${e.message}`,
    );
  }
}
```

### 5. Add to Menu

**In `src/config.js`**:

```javascript
.addItem("📊 Chart Tools", "showChartSidebar")
```

## Key Patterns Summary

### Sidebar Creation Pattern

1. **HTML Container**: Main sidebar HTML with includes
2. **Creation Function**: In `html_service_utils.js` using `createModularHtmlTemplate()`
3. **Show Function**: Calls creation function and displays sidebar
4. **Menu Integration**: Add menu item pointing to show function

### Component Organization Pattern

- **Shared components**: `src/components/` (styles, config-form, etc.)
- **Specialized components**: `src/components/[sidebar-type]/` (flowchart/, chart/, etc.)
- **Consistent naming**: `styles.html`, `scripts.html` in each directory
- **Focused responsibility**: Each component handles one specific area

### Infrastructure Pattern

- **`include()` function**: Loads HTML fragments
- **`createModularHtmlTemplate()`**: Sets up template with include capability
- **Creation functions**: One per sidebar type, following naming convention
- **Error handling**: Consistent try-catch blocks with user-friendly messages

## Conclusion

This modular approach makes sidebars more maintainable and extensible. The consistent pattern across all sidebars means:

- **Easy to understand**: Same structure for all sidebar types
- **Simple to extend**: Follow the established pattern for new sidebars
- **Maintainable**: Components are focused and independent
- **Scalable**: Can support unlimited specialized sidebars

By separating concerns and organizing code into focused components, you can easily update and expand functionality without navigating through large monolithic files.
