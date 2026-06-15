---
name: insert
description: Insert and style elements (tables, text boxes, shapes) onto Google Slides at exact positions via the Apps Script advanced Slides service. Use when adding/positioning page elements programmatically, or when "object could not be found" / wrong-style / can't-control-paste-position issues appear.
---

# Inserting styled elements into Google Slides

Hard-won rules for creating page elements (tables, text boxes, shapes) on a slide
with exact position and full styling. Reference implementations:
`src/util/table_minter.js` (advanced-API table) and `src/util/style_table.js`
(border styling on an existing table).

## The #1 gotcha: SlidesApp objects are invisible to the advanced API

A page element created with **SlidesApp** (`slide.insertTable`, `insertTextBox`,
`insertShape`) is **NOT yet committed**, so a follow-up
`Slides.Presentations.batchUpdate` that references its `getObjectId()` fails with:

> Invalid requests[0]....: The object (SLIDES_API…_1) could not be found.

There is **no `SlidesApp.flush()`** (that's `SpreadsheetApp` only — calling it
throws "flush is not a function"). You cannot force a commit mid-execution.

**Fix:** create the element through the **advanced API** with a **self-assigned
objectId**, then every styling request targets an id that exists.

```js
const tableId = "tbl" + Utilities.getUuid().replace(/-/g, ""); // unique, [A-Za-z0-9_-], 5–50 chars
```

Reads are fine via SlidesApp (get the current slide / `pageObjectId`):
```js
let slide;
try { slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide(); }
catch (e) { slide = SlidesApp.getActivePresentation().getSlides()[0]; } // fallback
const pageId = slide.getObjectId();
```

A text box you create with SlidesApp AND style with SlidesApp (never referenced by
the advanced API) is fine — the rule only bites when you cross services on one id.

## Do it all in ONE batchUpdate (atomic + instant)

Within a single `batchUpdate`, requests run in order and **later requests can
reference objects created by earlier requests in the same batch**. So `createTable`
can be `requests[0]`, with text/borders/columns after it. One call = one visible
change (no styling "one step at a time") and it's faster.

```js
const requests = [ { createTable: {...} }, /* cell text, borders, widths */ ];
try { Slides.Presentations.batchUpdate({ requests }, presentation.getId()); }
catch (e) { return { success: false, error: e.message }; }
```

Use many small `batchUpdate` calls ONLY for debugging (to learn which request the
API rejects); collapse to one once it works.

## Prerequisite: enable the advanced service

`appsscript.json` must contain:
```json
"dependencies": { "enabledAdvancedServices": [
  { "userSymbol": "Slides", "serviceId": "slides", "version": "v1" }
]}
```
(`Slides` shows a TypeScript "Could not find name" lint warning — ignore; it's a
runtime global.)

## Units: everything is POINTS (pt)

- `getLeft/getTop/getWidth/getHeight` and the inspector (檢視物件屬性) report **pt**.
- API `transform` (`translateX/Y`) and `size` (`width/height.magnitude`) use `unit: "PT"`.
- Pasted HTML `px` widths are read by Slides at 96dpi → convert **`pt = px * 0.75`**
  to match a clipboard paste's physical size.
- Slide top-left is the origin; `translateX/Y` = Left/Top.

```js
elementProperties: {
  pageObjectId: pageId,
  size: { width: {magnitude: wPt, unit:"PT"}, height: {magnitude: hPt, unit:"PT"} },
  transform: { scaleX:1, scaleY:1, translateX: leftPt, translateY: topPt, unit:"PT" },
}
```

## Table recipes

**Cell text** — `insertText` requires NON-empty text (skip empty cells):
```js
{ insertText: { objectId, cellLocation:{rowIndex,columnIndex}, insertionIndex:0, text } }
```
**Text style** — color is nested `foregroundColor.opaqueColor.rgbColor` (0..1 floats):
```js
{ updateTextStyle: { objectId, cellLocation, textRange:{type:"ALL"},
  style:{ foregroundColor:{opaqueColor:{rgbColor:{red,green,blue}}}, bold, fontFamily,
          fontSize:{magnitude, unit:"PT"} },
  fields:"foregroundColor,bold,fontFamily,fontSize" } }
```
**Paragraph align** — `START` (left), `CENTER`, `END`:
```js
{ updateParagraphStyle: { objectId, cellLocation, textRange:{type:"ALL"},
  style:{ alignment:"START" }, fields:"alignment" } }
```
**Vertical align** — whole table in one request via `tableRange`:
```js
{ updateTableCellProperties: { objectId,
  tableRange:{ location:{rowIndex:0,columnIndex:0}, rowSpan:numRows, columnSpan:numCols },
  tableCellProperties:{ contentAlignment:"MIDDLE" }, fields:"contentAlignment" } }
```
**Column width** — per column; enforce a sane minimum (~40pt):
```js
{ updateTableColumnProperties: { objectId, columnIndices:[c],
  tableColumnProperties:{ columnWidth:{ magnitude:Math.max(40, wPt), unit:"PT" } },
  fields:"columnWidth" } }
```

## Borders: hide-all then add back

New tables have a full gray grid. To get a clean look (no verticals, header rule,
thin row separators):

1. **Hide everything** with a transparent fill — `solidFill` supports `alpha`,
   and `alpha: 0` makes a border invisible (don't paint white — it shows on
   non-white slides):
   ```js
   tableBorderFill:{ solidFill:{ color:{rgbColor:{red:1,green:1,blue:1}}, alpha:0 } }
   ```
2. Add `INNER_HORIZONTAL` (between rows) + `BOTTOM` (last-row underline), thin gray.
3. **Target one row** (e.g. accent strokes above/below the header) with `tableRange`:
   ```js
   { updateTableBorderProperties: { objectId,
     tableRange:{ location:{rowIndex:0,columnIndex:0}, rowSpan:1, columnSpan:numCols },
     borderPosition:"BOTTOM",  // and "TOP"
     tableBorderProperties:{ tableBorderFill:{solidFill:{color:{rgbColor:accent}}},
       weight:{magnitude:1.5, unit:"PT"}, dashStyle:"SOLID" },
     fields:"tableBorderFill,weight,dashStyle" } }
   ```
   `borderPosition`: `ALL | INNER | OUTER | INNER_HORIZONTAL | INNER_VERTICAL |
   LEFT | RIGHT | TOP | BOTTOM`. Omit `tableRange` to hit the whole table. Later
   requests override earlier ones, so order: hide-all → gray lines → accent header.

## Clipboard vs. API — when to use which

- **Clipboard (`text/html`)**: produces a pixel-perfect *native* table on paste,
  but you **cannot control paste position** — Slides decides placement. Good for
  "copy, I'll paste it myself."
- **Advanced API insert**: exact Left/Top/size and no manual paste; styling is
  rebuilt request-by-request (very close, not always byte-identical). Use when
  position matters.

## Checklist for a new "insert X" feature

- [ ] Create via advanced API with a `Utilities.getUuid()` objectId (not SlidesApp).
- [ ] One `batchUpdate`; createTable/createShape first, styling after.
- [ ] All geometry in PT; convert px→pt at ×0.75 if mirroring a paste.
- [ ] Skip `insertText` for empty strings.
- [ ] Colors as `{rgbColor:{red,green,blue}}` 0..1 (helper: hex→rgb /255).
- [ ] Hide unwanted borders with `alpha:0`, not white.
- [ ] Wrap `batchUpdate` in try/catch; return `{success, error}` to the client.
- [ ] Client `google.script.run.withSuccessHandler/withFailureHandler`; close the
      dialog only on success.
