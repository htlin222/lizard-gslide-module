# Google Slides API Styling Guide

This guide provides best practices and examples for styling elements in Google Slides using Google Apps Script.

## Table of Contents
- [Shape and Text Box Styling](#shape-and-text-box-styling)
- [Border Styling](#border-styling)
- [Fill Styling](#fill-styling)
- [Text Styling](#text-styling)
- [Common Errors and Solutions](#common-errors-and-solutions)
- [Complete Examples](#complete-examples)

## Shape and Text Box Styling

### Getting Elements from Selection

```javascript
// Get the active presentation and selection
const presentation = SlidesApp.getActivePresentation();
const selection = presentation.getSelection();
const selectedElements = selection.getPageElementRange() ? selection.getPageElementRange().getPageElements() : [];

// Process each selected element
for (let i = 0; i < selectedElements.length; i++) {
  const element = selectedElements[i];
  
  // Check if the element is a shape
  if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
    const shape = element.asShape();
    // Style the shape...
  }
  
  // Check if the element is a text box (which is also a shape in Google Slides)
  else if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE && 
           element.asShape().getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
    const textBox = element.asShape();
    // Style the text box...
  }
}
```

## Border Styling

### Setting Border Weight and Color

**IMPORTANT**: When styling borders, you must use separate statements for weight and color. Do not chain these methods.

```javascript
// CORRECT WAY
// First set the border weight
shape.getBorder().setWeight(2); // 2 points thickness

// Then set the border color using getLineFill()
shape.getBorder().getLineFill().setSolidFill('#FF0000'); // Red border
```

**INCORRECT WAY** (will cause errors):
```javascript
// This will cause an error
shape.getBorder().setWeight(2).setSolidFill('#FF0000');
```

### Border Dash Style

```javascript
// Set a dashed border
shape.getBorder().setDashStyle(SlidesApp.DashStyle.DASH);

// Available dash styles:
// - SlidesApp.DashStyle.SOLID (default)
// - SlidesApp.DashStyle.DASH
// - SlidesApp.DashStyle.DOT
// - SlidesApp.DashStyle.DASH_DOT
```

## Fill Styling

### Solid Fill

```javascript
// Set a solid fill color
shape.getFill().setSolidFill('#3D6869'); // Teal color

// With transparency (0-1, where 1 is fully opaque)
shape.getFill().setSolidFill('#3D6869', 0.8); // 20% transparent
```

### No Fill

```javascript
// Remove fill (transparent)
shape.getFill().setSolidFill('transparent');
// OR
shape.getFill().setTransparent();
```

### Gradient Fill

```javascript
// Create a linear gradient
const gradient = SlidesApp.GradientType.LINEAR;
const stops = [
  {color: '#FF0000', position: 0},    // Red at start
  {color: '#0000FF', position: 1}     // Blue at end
];
shape.getFill().setGradient(gradient, stops);
```

## Text Styling

### Setting Text Color

```javascript
// Set text color
shape.getText().getTextStyle().setForegroundColor('#000000'); // Black text
```

### Font Formatting

```javascript
// Get the text range
const textRange = shape.getText();

// Apply multiple styles
textRange.getTextStyle()
  .setForegroundColor('#000000')      // Text color
  .setFontFamily('Arial')             // Font family
  .setFontSize(12)                    // Font size in points
  .setBold(true)                      // Bold text
  .setItalic(false)                   // Not italic
  .setUnderline(false);               // Not underlined
```

### Paragraph Alignment

```javascript
// Set text alignment
textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

// Available alignments:
// - SlidesApp.ParagraphAlignment.START (left-aligned)
// - SlidesApp.ParagraphAlignment.CENTER
// - SlidesApp.ParagraphAlignment.END (right-aligned)
// - SlidesApp.ParagraphAlignment.JUSTIFIED
```

### Vertical Alignment

```javascript
// Set vertical alignment within the shape
shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

// Available alignments:
// - SlidesApp.ContentAlignment.TOP
// - SlidesApp.ContentAlignment.MIDDLE
// - SlidesApp.ContentAlignment.BOTTOM
```

## Common Errors and Solutions

### Error: "setSolidFill is not a function"

**Problem**: Attempting to call `setSolidFill()` directly on a border object.

**Solution**: Use `getLineFill()` before setting the solid fill:
```javascript
// WRONG
shape.getBorder().setSolidFill('#FF0000');

// CORRECT
shape.getBorder().getLineFill().setSolidFill('#FF0000');
```

### Error: "Cannot call method on null object"

**Problem**: Trying to style an element that doesn't have the expected property.

**Solution**: Always check if the property exists before styling:
```javascript
// Check if the shape has text before styling it
if (shape.getText()) {
  shape.getText().getTextStyle().setForegroundColor('#000000');
}
```

## Complete Examples

### Example 1: Styling a Shape with White Fill and Colored Border

```javascript
function styleShape(shape, mainColor) {
  // Set fill color to white
  shape.getFill().setSolidFill('#FFFFFF');
  
  // Set border weight and color
  shape.getBorder().setWeight(1);
  shape.getBorder().getLineFill().setSolidFill(mainColor);
  
  // Set text color if the shape has text
  if (shape.getText()) {
    shape.getText().getTextStyle().setForegroundColor(mainColor);
  }
}
```

### Example 2: Creating a Callout Box

```javascript
function createCallout(slide, x, y, width, height, mainColor) {
  // Create main box
  const mainBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, width, height);
  mainBox.getFill().setSolidFill('#FFFFFF');
  mainBox.getBorder().setWeight(1);
  mainBox.getBorder().getLineFill().setSolidFill(mainColor);
  
  // Create header
  const headerHeight = 20;
  const header = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    x,
    y - headerHeight,
    width,
    headerHeight
  );
  
  // Style header
  header.getFill().setSolidFill(mainColor);
  header.getBorder().setWeight(1);
  header.getBorder().getLineFill().setSolidFill(mainColor);
  
  // Add text to header
  header.getText().setText('TITLE');
  header.getText().getTextStyle()
    .setForegroundColor('#FFFFFF')
    .setBold(true);
  header.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Group the elements
  return slide.group([header, mainBox]);
}
```

---

This guide was created for the Lizard Slides Module to help developers correctly use the Google Slides API for styling elements.

Last updated: May 2025
