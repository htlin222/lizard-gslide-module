// Section boxes module for Google Slides
/**
 * For each SECTION_HEADER slide, insert two text boxes showing
 * before/after section titles. "before" boxes bottom-align text,
 * "after" boxes top-align text.
 */
function processSectionBoxes(slides, requests) {
  const sections = getSectionHeaders(slides);
  if (!sections.length) return;

  sections.forEach((sec, idx) => {
    const slide    = slides[sec.index];
    const slideId  = sec.slideId;

    // Clean up any previous boxes
    deleteOldBoxes(slide, requests, 'before_');
    deleteOldBoxes(slide, requests, 'after_');
    deleteOldBoxes(slide, requests, 'label_');

    const beforeTitles = sections.slice(0, idx).map(s => s.title);
    const afterTitles  = sections.slice(idx + 1).map(s => s.title);

    // Before titles
    if (beforeTitles.length) {
      const beforeId = `before_${slideId}_${newGuid()}`;
      requests.push(
        createShapeRequest(beforeId, slideId, BOX_CONFIG.x, BOX_CONFIG.yBefore, BOX_CONFIG.width),
        {
          updateShapeProperties: {
            objectId: beforeId,
            shapeProperties: { contentAlignment: 'BOTTOM' },
            fields: 'contentAlignment'
          }
        },
        insertTextRequest(beforeId, beforeTitles),
        textStyleRequest(beforeId, BOX_CONFIG.fontSize, BOX_CONFIG.fontFamily, BOX_CONFIG.textColor, BOX_CONFIG.isBold),
        paragraphStyleRequest(beforeId)
      );
    }

    // After titles
    if (afterTitles.length) {
      const afterId = `after_${slideId}_${newGuid()}`;
      requests.push(
        createShapeRequest(afterId, slideId, BOX_CONFIG.x, BOX_CONFIG.yAfter, BOX_CONFIG.width),
        {
          updateShapeProperties: {
            objectId: afterId,
            shapeProperties: { contentAlignment: 'TOP' },
            fields: 'contentAlignment'
          }
        },
        insertTextRequest(afterId, afterTitles),
        textStyleRequest(afterId, BOX_CONFIG.fontSize, BOX_CONFIG.fontFamily, BOX_CONFIG.textColor, BOX_CONFIG.isBold),
        paragraphStyleRequest(afterId)
      );
    }

    // Section label
    const labelId = `label_${slideId}_${newGuid()}`;
    requests.push(
      {
        createShape: {
          objectId: labelId,
          shapeType: 'TEXT_BOX',
          elementProperties: {
            pageObjectId: slideId,
            size: { width: { magnitude: 80, unit: 'PT' }, height: { magnitude: 25, unit: 'PT' } },
            transform: { translateX: 50, translateY: 50, scaleX: 1, scaleY: 1, unit: 'PT' }
          }
        }
      },
      {
        updateShapeProperties: {
          objectId: labelId,
          shapeProperties: {
            contentAlignment: 'MIDDLE',
            shapeBackgroundFill: solidFill(main_color)
          },
          fields: 'contentAlignment,shapeBackgroundFill.solidFill.color'
        }
      },
      insertTextRequest(labelId, [`Section: ${idx + 1}`]),
      textStyleRequest(labelId, label_font_size, BOX_CONFIG.fontFamily, '#FFFFFF', true),
      paragraphStyleRequest(labelId)
    );
  });

  addOutlineInSecondPage(slides, requests);
}


// Add outline to the second slide titled "Outline"
function addOutlineInSecondPage(slides, requests) {
  const sections = getSectionHeaders(slides);
  if (!sections.length) return;
  const secondSlide = slides[1];
  if (!secondSlide) return;
  deleteOldBoxes(secondSlide, requests, 'outline_');
  let title = '';
  const placeholder = secondSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
  if (placeholder && placeholder.asShape) {
    title = placeholder.asShape().getText().asString().trim();
  } else {
    title = getFirstTextboxText(secondSlide);
  }
  if (title !== 'Outline') return;

  const outlineTitles = sections.map(s => s.title);
  if (!outlineTitles.length) return;

  const outlineId = `outline_${secondSlide.getObjectId()}_${newGuid()}`;
  requests.push(
    {
      createShape: {
        objectId: outlineId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: secondSlide.getObjectId(),
          size: {
            width:  { magnitude: 400, unit: 'PT' },
            height: { magnitude: 300, unit: 'PT' }
          },
          transform: {
            translateX: 280,
            translateY: 51,
            scaleX: 1,
            scaleY: 1,
            unit: 'PT'
          }
        }
      }
    },
    {
      updateShapeProperties: {
        objectId: outlineId,
        shapeProperties: {
          contentAlignment: 'MIDDLE'
        },
        fields: 'contentAlignment'
      }
    },
    insertTextRequest(outlineId, outlineTitles),
    textStyleRequest(outlineId, 28, main_font_family, main_color, false),
    {
      createParagraphBullets: {
        objectId: outlineId,
        textRange: { type: 'ALL' },
        bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE'
      }
    }
  );
}


// Config
const box_width = 600;
const box_x = (720 - box_width) / 2;

const BOX_CONFIG = {
  x: box_x,
  yBefore: 30,
  yAfter: 240,
  width: box_width,
  boxHeight: 150,
  fontSize: 20,
  fontFamily: main_font_family,
  textColor: '#aaaaaa',
  isBold: false
};


// Helpers
function getSectionHeaders(slides) {
  return slides.map((slide, i) => {
    if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') {
      const title = getFirstTextboxText(slide);
      return title ? { title, index: i, slideId: slide.getObjectId() } : null;
    }
    return null;
  }).filter(Boolean);
}

function getFirstTextboxText(slide) {
  for (const shape of slide.getShapes()) {
    if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      const txt = shape.getText().asString().trim();
      if (txt) return txt;
    }
  }
  return '';
}

function deleteOldBoxes(slide, requests, prefix) {
  slide.getShapes().forEach(shape => {
    if (shape.getObjectId().startsWith(prefix)) {
      requests.push({ deleteObject: { objectId: shape.getObjectId() } });
    }
  });
}

function createShapeRequest(objectId, pageObjectId, x, y, w) {
  return {
    createShape: {
      objectId,
      shapeType: 'TEXT_BOX',
      elementProperties: {
        pageObjectId,
        size: {
          width: { magnitude: w, unit: 'PT' },
          height: { magnitude: BOX_CONFIG.boxHeight, unit: 'PT' }
        },
        transform: {
          translateX: x,
          translateY: y,
          scaleX: 1,
          scaleY: 1,
          unit: 'PT'
        }
      }
    }
  };
}

function insertTextRequest(objectId, lines) {
  return {
    insertText: {
      objectId,
      text: lines.join('\n')
    }
  };
}

function textStyleRequest(objectId, fontSize, fontFamily, hexColor, isBold = false) {
  return {
    updateTextStyle: {
      objectId,
      textRange: { type: 'ALL' },
      style: {
        fontSize: { magnitude: fontSize, unit: 'PT' },
        fontFamily,
        foregroundColor: { opaqueColor: { rgbColor: hexToRgb(hexColor) } },
        bold: isBold
      },
      fields: 'fontSize,fontFamily,foregroundColor,bold'
    }
  };
}

function paragraphStyleRequest(objectId) {
  return {
    updateParagraphStyle: {
      objectId,
      textRange: { type: 'ALL' },
      style: { alignment: 'CENTER' },
      fields: 'alignment'
    }
  };
}

function newGuid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}