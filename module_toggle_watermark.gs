// Toggle watermark module for Google Slides
/**
 * 切換浮水印：若頁面已有 title='WATERMARK' 的浮水印就刪除，
 * 否則就新增半透明、45° 旋轉、置中的文字浮水印。
 */
function toggleWaterMark(slides, requests) {
  const aColour = '#efefef';
  const slideWidth  = 720;
  const slideHeight = 405;
  const wmWidth  = 500;
  const wmHeight = 100;
  const dx = (slideWidth )  / 2;
  const dy = (slideHeight ) / 2 - 2 * wmHeight;
  const cos45 = Math.SQRT1_2;
  const sin45 = Math.SQRT1_2;

  slides.forEach(slide => {
    const slideId = slide.getObjectId();
    // 檢查該 slide 是否已有浮水印
    const existing = slide.getPageElements().filter(el => {
      return el.getPageElementType() === SlidesApp.PageElementType.SHAPE
          && el.asShape().getTitle && el.asShape().getTitle() === 'WATERMARK';
    });

    if (existing.length) {
      // 已有：刪除所有舊的 WATERMARK
      existing.forEach(shape => {
        requests.push({ deleteObject: { objectId: shape.getObjectId() } });
      });
    } else {
      // 不存在：新增浮水印
      const wmId = newObjectId(slideId);
      // 1) 建立文字框並旋轉 45°
      requests.push({
        createShape: {
          objectId: wmId,
          shapeType: 'TEXT_BOX',
          elementProperties: {
            pageObjectId: slideId,
            size: {
              width:  { magnitude: wmWidth,  unit: 'PT' },
              height: { magnitude: wmHeight, unit: 'PT' }
            },
            transform: {
              scaleX:     cos45,
              shearX:    -sin45,
              shearY:     sin45,
              scaleY:     cos45,
              translateX: dx,
              translateY: dy,
              unit:      'PT'
            }
          }
        }
      });
      // 3) 插入文字
      requests.push({
        insertText: {
          objectId:       wmId,
          insertionIndex: 0,
          text:           water_mark_text,
        }
      });
      // 4) 文字樣式：字型大小、顏色
      requests.push({
        updateTextStyle: {
          objectId: wmId,
          textRange: { type: 'ALL' },
          style: {
            fontSize: { magnitude: 56, unit: 'PT' },
            foregroundColor: { opaqueColor: { rgbColor: hexToRgb(aColour) } }
          },
          fields: 'fontSize,foregroundColor'
        }
      });
      // 5) 段落置中
      requests.push({
        updateParagraphStyle: {
          objectId:  wmId,
          textRange: { type: 'ALL' },
          style:     { alignment: 'CENTER' },
          fields:    'alignment'
        }
      });
      // 6) 垂直置中
      requests.push({
        updateShapeProperties: {
          objectId:       wmId,
          shapeProperties:{ contentAlignment: 'MIDDLE' },
          fields:         'contentAlignment'
        }
      });
      // 7) 打上 title 標記方便後續刪除
      requests.push({
        updatePageElementAltText: {
          objectId: wmId,
          title:    'WATERMARK'
        }
      });
    }
  });
}