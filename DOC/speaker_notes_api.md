# Google Slides Speaker Notes API Guide

This document explains how to properly read and write speaker notes in Google Slides using Google Apps Script.

## The Critical Pattern

Working with speaker notes requires a specific API chain and validation steps that are **essential** for proper functionality.

## Working Example

```javascript
function getSpeakerNotesFromCurrentSlide() {
  var presentation = SlidesApp.getActivePresentation();
  var selection = presentation.getSelection();
  var currentPage = selection.getCurrentPage();

  if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
    SlidesApp.getUi().alert("Please select a slide first.");
    return;
  }

  // Convert currentPage to a Slide object
  var slide = currentPage.asSlide();
  var notesPage = slide.getNotesPage();
  var shape = notesPage.getSpeakerNotesShape();

  var notesText = shape ? shape.getText().asString() : "";
  SlidesApp.getUi().alert(
    "Speaker Notes:\n" + (notesText || "(No notes found)"),
  );
}
```

## API Chain Breakdown

### 1. Get the Selection Chain

```javascript
const presentation = SlidesApp.getActivePresentation();
const selection = presentation.getSelection(); // IMPORTANT: Get selection object
const currentPage = selection.getCurrentPage();
```

### 2. Validate Page Type (CRITICAL)

```javascript
if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
  // Not a slide - handle error
  return;
}
```

### 3. Convert to Slide Object (ESSENTIAL)

```javascript
const slide = currentPage.asSlide(); // MUST convert page to slide
```

### 4. Access Speaker Notes

```javascript
const notesPage = slide.getNotesPage();
const shape = notesPage.getSpeakerNotesShape();
const notesText = shape ? shape.getText().asString() : "";
```

## Common Mistakes to Avoid

### ❌ Wrong: Direct getCurrentPage() usage

```javascript
// This will NOT work reliably
const currentSlide = presentation.getSelection().getCurrentPage();
const notesPage = currentSlide.getNotesPage(); // ERROR!
```

### ❌ Wrong: Missing page type validation

```javascript
// This fails when master slides or layouts are selected
const currentPage = selection.getCurrentPage();
const slide = currentPage.asSlide(); // ERROR if not a slide!
```

### ❌ Wrong: Assuming shape exists

```javascript
// This crashes if shape is null
const shape = notesPage.getSpeakerNotesShape();
const text = shape.getText().asString(); // ERROR if shape is null!
```

### ✅ Correct: Full validation chain

```javascript
const presentation = SlidesApp.getActivePresentation();
const selection = presentation.getSelection();
const currentPage = selection.getCurrentPage();

if (!currentPage || currentPage.getPageType() !== SlidesApp.PageType.SLIDE) {
  throw new Error("Please select a slide first");
}

const slide = currentPage.asSlide();
const notesPage = slide.getNotesPage();
const shape = notesPage.getSpeakerNotesShape();

const notesText = shape ? shape.getText().asString() : "";
```

## Complete Implementation Examples

### Reading Speaker Notes

```javascript
function getCurrentSpeakerNotes() {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentPage = selection.getCurrentPage();

    if (
      !currentPage ||
      currentPage.getPageType() !== SlidesApp.PageType.SLIDE
    ) {
      return "";
    }

    const slide = currentPage.asSlide();
    const notesPage = slide.getNotesPage();
    const shape = notesPage.getSpeakerNotesShape();

    const notesText = shape ? shape.getText().asString() : "";
    return notesText.trim();
  } catch (e) {
    console.error(`Error getting speaker notes: ${e.message}`);
    return "";
  }
}
```

### Writing Speaker Notes (Replace)

```javascript
function replaceSpeakerNotes(newText) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentPage = selection.getCurrentPage();

    if (
      !currentPage ||
      currentPage.getPageType() !== SlidesApp.PageType.SLIDE
    ) {
      throw new Error("Please select a slide first");
    }

    const slide = currentPage.asSlide();
    const notesPage = slide.getNotesPage();
    const shape = notesPage.getSpeakerNotesShape();

    if (!shape) {
      throw new Error("Could not access speaker notes shape");
    }

    shape.getText().setText(newText || "");

    return {
      success: true,
      message: "Speaker notes replaced successfully",
    };
  } catch (e) {
    return {
      success: false,
      message: `Failed to replace speaker notes: ${e.message}`,
    };
  }
}
```

### Writing Speaker Notes (Append)

```javascript
function appendToSpeakerNotes(textToAppend) {
  try {
    const presentation = SlidesApp.getActivePresentation();
    const selection = presentation.getSelection();
    const currentPage = selection.getCurrentPage();

    if (
      !currentPage ||
      currentPage.getPageType() !== SlidesApp.PageType.SLIDE
    ) {
      throw new Error("Please select a slide first");
    }

    const slide = currentPage.asSlide();
    const notesPage = slide.getNotesPage();
    const shape = notesPage.getSpeakerNotesShape();

    if (!shape) {
      throw new Error("Could not access speaker notes shape");
    }

    const currentText = shape.getText().asString();
    const newText = currentText.trim()
      ? `${currentText.trim()}\n\n${textToAppend}`
      : textToAppend;

    shape.getText().setText(newText);

    return {
      success: true,
      message: "Speaker notes updated successfully",
    };
  } catch (e) {
    return {
      success: false,
      message: `Failed to update speaker notes: ${e.message}`,
    };
  }
}
```

## Google Slides Object Hierarchy

Understanding the object relationships:

```
Presentation
├── getSelection()
│   └── getCurrentPage() → Page
│       ├── getPageType() → PageType.SLIDE
│       └── asSlide() → Slide
└── getSlides() → Slide[]

Slide
└── getNotesPage() → NotesPage
    └── getSpeakerNotesShape() → Shape
        └── getText() → TextRange
            ├── asString() → string
            └── setText(text) → void
```

## Key Insights Learned

1. **Selection vs Direct Access**: Always use `presentation.getSelection()` rather than trying to access slides directly
2. **Page Type Matters**: Always validate that you're working with an actual slide, not a master or layout
3. **Object Conversion**: Must convert `Page` to `Slide` using `.asSlide()`
4. **Shape Validation**: Speaker notes shape can be null, always check before using
5. **Error Handling**: Wrap in try-catch for robust error handling

## Testing Your Implementation

Create a simple test function:

```javascript
function testSpeakerNotes() {
  // Test reading
  const currentNotes = getCurrentSpeakerNotes();
  console.log("Current notes:", currentNotes);

  // Test writing
  const result = replaceSpeakerNotes("Test notes from Apps Script");
  console.log("Write result:", result);

  // Verify
  const newNotes = getCurrentSpeakerNotes();
  console.log("New notes:", newNotes);
}
```

## Best Practices

1. **Always validate page type** before accessing slide-specific methods
2. **Use consistent variable naming** (`slide`, `shape`, `notesText`)
3. **Handle null shapes gracefully** with ternary operators
4. **Provide clear error messages** to help users understand what went wrong
5. **Return structured objects** for success/error status in write operations
6. **Trim text content** to remove unnecessary whitespace
7. **Use proper spacing** (`\n\n`) when appending to existing notes

This pattern has been tested and proven to work reliably with Google Slides speaker notes.
