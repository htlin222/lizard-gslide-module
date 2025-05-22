# Google Apps Script Sidebar Integration Guide

This guide explains how to create, register, and use functions in Google Apps Script with HTML sidebars for the Lizard Slides Module.

## Table of Contents
- [Understanding the Sidebar Architecture](#understanding-the-sidebar-architecture)
- [Creating Server-Side Functions](#creating-server-side-functions)
- [Exposing Functions to the Sidebar](#exposing-functions-to-the-sidebar)
- [Calling Server Functions from the Sidebar](#calling-server-functions-from-the-sidebar)
- [Adding UI Elements to the Sidebar](#adding-ui-elements-to-the-sidebar)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Complete Examples](#complete-examples)

## Understanding the Sidebar Architecture

Google Apps Script uses a client-server model for sidebar interactions:

1. **Server-side**: JavaScript functions in `.js` or `.gs` files run on Google's servers
2. **Client-side**: HTML, CSS, and client-side JavaScript in the sidebar run in the user's browser
3. **Communication**: The `google.script.run` API bridges the gap between client and server

```
┌─────────────────┐         ┌─────────────────┐
│                 │         │                 │
│  Client-Side    │         │  Server-Side    │
│  (Sidebar HTML) │◄────────┤  (Script Files) │
│                 │         │                 │
└─────────────────┘         └─────────────────┘
       ▲                           ▲
       │                           │
       └───────────────────────────┘
           google.script.run API
```

## Creating Server-Side Functions

1. Create your function in a `.js` file:

```javascript
/**
 * Apply a specific style to selected elements
 * @param {number} styleNumber - The style number to apply
 * @return {boolean} - Success indicator
 */
function applyStyle(styleNumber) {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  
  // Function implementation...
  
  return true; // Always return a value to confirm execution
}

// Helper functions that call the main function
function applyStyle1() {
  return applyStyle(1);
}

function applyStyle2() {
  return applyStyle(2);
}
```

## Exposing Functions to the Sidebar

For a function to be callable from the sidebar:

1. It must be a **top-level function** (not inside another function)
2. It should **return a value** to confirm execution
3. It must be **public** (no private scope)

```javascript
// ✅ CORRECT: Properly exposed function
function myPublicFunction() {
  // Implementation
  return true;
}

// ❌ INCORRECT: Function inside another function
function outerFunction() {
  function innerFunction() { // Not callable from sidebar
    // Implementation
  }
}

// ❌ INCORRECT: No return value
function noReturnFunction() {
  // Implementation with no return
}
```

## Calling Server Functions from the Sidebar

Use the `google.script.run` API with success and failure handlers:

```javascript
// Basic pattern
google.script.run
  .withSuccessHandler(function(returnValue) {
    // Handle success
  })
  .withFailureHandler(function(error) {
    // Handle error
  })
  .myServerFunction(param1, param2);
```

### Best Practice Pattern with Loading State

```javascript
function callServerFunction(functionName, params) {
  // Get the button element
  const button = document.getElementById('my-button');
  
  // Show loading state
  button.classList.add('loading');
  
  // Call the server function
  google.script.run
    .withSuccessHandler(function(result) {
      // Hide loading state
      button.classList.remove('loading');
      
      // Show success message
      showStatusMessage('Operation completed successfully!', 'success');
    })
    .withFailureHandler(function(error) {
      // Hide loading state
      button.classList.remove('loading');
      
      // Show error message
      showStatusMessage('Error: ' + error.message, 'error');
    })
    [functionName](params); // Dynamic function call
}
```

## Adding UI Elements to the Sidebar

1. Add HTML elements to `sidebar.html`:

```html
<div class="form-group">
  <label>Style Options:</label>
  <div class="button-container">
    <div class="style-button" id="style-button-1">
      <div class="style-preview">
        <div class="preview-text">A</div>
      </div>
      <div class="style-label">Style 1</div>
    </div>
    <!-- More buttons... -->
  </div>
</div>
```

2. Add CSS for styling:

```html
<style>
  .style-button {
    cursor: pointer;
    text-align: center;
    width: 30%;
    padding: 5px;
    border-radius: 4px;
    transition: background-color 0.2s;
  }
  .style-button:hover {
    background-color: #f5f5f5;
  }
  .style-preview {
    width: 40px;
    height: 40px;
    margin: 0 auto 5px;
    border-radius: 4px;
    display: flex;
    justify-content: center;
    align-items: center;
  }
  /* More styles... */
</style>
```

3. Add JavaScript to connect UI elements to server functions:

```html
<script>
  // Set up event listeners when the page loads
  window.onload = function() {
    // Initialize the sidebar
    initializeSidebar();
    
    // Set up button event listeners
    document.getElementById('style-button-1').addEventListener('click', function() {
      applyStyle(1);
    });
    
    document.getElementById('style-button-2').addEventListener('click', function() {
      applyStyle(2);
    });
    
    // More event listeners...
  };
  
  // Function to call server-side functions
  function applyStyle(styleNumber) {
    const button = document.getElementById('style-button-' + styleNumber);
    button.classList.add('loading');
    
    google.script.run
      .withSuccessHandler(function(result) {
        button.classList.remove('loading');
        showStatusMessage('Style applied!', 'success');
      })
      .withFailureHandler(function(error) {
        button.classList.remove('loading');
        showStatusMessage('Error: ' + error.message, 'error');
      })
      ['applyStyle' + styleNumber]();
  }
  
  // Function to show status messages
  function showStatusMessage(message, type) {
    // Implementation...
  }
</script>
```

## Best Practices

1. **Always return values** from server-side functions
2. **Use descriptive function names** that clearly indicate their purpose
3. **Add loading indicators** to provide visual feedback during server calls
4. **Implement error handling** on both client and server sides
5. **Use success and error messages** to inform users about operation results
6. **Separate concerns** between UI, event handling, and server communication
7. **Document your functions** with JSDoc comments

## Troubleshooting

### Common Issues and Solutions

1. **Function not found error**
   - Ensure the function is a top-level function
   - Check for typos in the function name
   - Verify the function is in an active script file

2. **No response from server**
   - Add proper success and failure handlers
   - Check browser console for errors
   - Verify the function returns a value

3. **UI not updating after server call**
   - Make sure DOM updates happen in the success handler
   - Check if the function is executing asynchronously

## Complete Examples

### Example 1: Style Application Buttons

**Server-side function (util_default_style.js):**
```javascript
/**
 * Apply style 1: White fill, main color border and text
 * @return {boolean} Success indicator
 */
function applyStyle1() {
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const selectedElements = selection.getPageElementRange() ? 
    selection.getPageElementRange().getPageElements() : [];
  
  // Get the main color from script properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const mainColor = scriptProperties.getProperty('mainColor') || '#3D6869';
  
  // Apply the style to each selected element
  for (let i = 0; i < selectedElements.length; i++) {
    const element = selectedElements[i];
    
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      
      // Apply white fill
      shape.getFill().setSolidFill('#FFFFFF');
      
      // Apply border
      shape.getBorder().setWeight(1);
      shape.getBorder().getLineFill().setSolidFill(mainColor);
      
      // Apply text color if the shape has text
      if (shape.getText()) {
        shape.getText().getTextStyle().setForegroundColor(mainColor);
      }
    }
  }
  
  return true;
}
```

**Sidebar HTML (sidebar.html):**
```html
<!-- Style button in the sidebar -->
<div class="style-button" id="style-button-1" title="White fill, main color border and text">
  <div class="style-preview" id="style-preview-1">
    <div class="preview-text">A</div>
  </div>
  <div class="style-label">Style 1</div>
</div>

<!-- JavaScript to handle the button click -->
<script>
  document.getElementById('style-button-1').addEventListener('click', function() {
    // Show loading state
    this.classList.add('loading');
    
    // Call the server function
    google.script.run
      .withSuccessHandler(function(result) {
        // Hide loading state
        document.getElementById('style-button-1').classList.remove('loading');
        // Show success message
        showStatusMessage('Style applied successfully!', 'success');
      })
      .withFailureHandler(function(error) {
        // Hide loading state
        document.getElementById('style-button-1').classList.remove('loading');
        // Show error message
        showStatusMessage('Error: ' + error.message, 'error');
      })
      .applyStyle1();
  });
</script>
```

---

This guide was created for the Lizard Slides Module to help developers correctly integrate new functionality into the sidebar.

Last updated: May 2025
