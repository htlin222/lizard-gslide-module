// Right footer module for Google Slides
/**
 * 生成全局唯一的 objectId
 */
function newObjectId(slideId) {
  const uuidPart  = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const timestamp = Date.now().toString(36);
  return `obj_${slideId}_${timestamp}_${uuidPart}`;
}

/**
 * 从第一张幻灯片抓主标题文字
 */
function getMainTitleFromFirstSlide(slide) {
  const elements = slide.getPageElements();
  for (let el of elements) {
    if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const txt = el.asShape().getText().asString().trim();
      if (txt) return txt;
    }
  }
  return '';
}

/**
 * 批次插入旋转 90° 的脚注文字框（title='MAIN_TITLE'）
 * 并将文字链接到第一张幻灯片（通过 pageObjectId）
 */
function updateTitleFootnotes(slides, requests) {
  if (slides.length < 2) return;

  // 取第一张幻灯片的 objectId，当做链接目标
  const firstSlideId = slides[0].getObjectId();
  const mainTitle    = getMainTitleFromFirstSlide(slides[0]);
  const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
  const slideHeight = SlidesApp.getActivePresentation().getPageHeight();
  if (!mainTitle) return;
  
  for (let i = 1; i < slides.length; i++) {
    const slide   = slides[i];
    const slideId = slide.getObjectId();

    // 1) 删除旧的 MAIN_TITLE
    slide.getPageElements().forEach(el => {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = el.asShape();
        if (shape.getTitle && shape.getTitle() === 'MAIN_TITLE') {
          requests.push({ deleteObject: { objectId: shape.getObjectId() } });
        }
      }
    });

    // 2) 创建新的文本框，并内建 90° 旋转矩阵、指定新位移
    //    90° 旋转矩阵 R = [[0, -1],[1, 0]] 对应 scaleX, shearX, shearY, scaleY
    const footnoteId = newObjectId(slideId);
    const box_width = 360;
    const box_y = (slideHeight - box_width) / 2
    requests.push({
      createShape: {
        objectId: footnoteId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: slideId,
          size: {
            width:  { magnitude: box_width,  unit: 'PT' },
            height: { magnitude: 30, unit: 'PT' }
          },
          transform: {
            scaleX:     0,
            shearX:    -1,
            shearY:     1,
            scaleY:     0,
            translateX: slideWidth,
            translateY: box_y,
            unit:      'PT'
          }
        }
      }
    });

    // 3) 插入文字
    requests.push({
      insertText: {
        objectId:       footnoteId,
        insertionIndex: 0,
        text:           mainTitle
      }
    });

    // 4) 设置链接到第一张幻灯片（通过 pageObjectId）
    requests.push({
      updateTextStyle: {
        objectId: footnoteId,
        textRange: { type: 'ALL' },
        style: {
          link: { pageObjectId: firstSlideId }
        },
        fields: 'link'
      }
    });

    // 5) 文字样式：颜色、大小、字体、下划线
    requests.push({
      updateTextStyle: {
        objectId: footnoteId,
        textRange: { type: 'ALL' },
        style: {
          foregroundColor: { opaqueColor: { rgbColor: hexToRgb('#888888') } },
          fontSize:   { magnitude: 10, unit: 'PT' },
          fontFamily: main_font_family,
          underline:  false
        },
        fields: 'foregroundColor,fontSize,fontFamily,underline'
      }
    });

    // 6) 段落居中
    requests.push({
      updateParagraphStyle: {
        objectId:  footnoteId,
        textRange: { type: 'ALL' },
        style:     { alignment: 'CENTER' },
        fields:    'alignment'
      }
    });

    // 7) 垂直居中内容
    requests.push({
      updateShapeProperties: {
        objectId:       footnoteId,
        shapeProperties:{ contentAlignment: 'MIDDLE' },
        fields:         'contentAlignment'
      }
    });

    // 8) 打上 title 标记，方便后续删除
    requests.push({
      updatePageElementAltText: {
        objectId: footnoteId,
        title:    'MAIN_TITLE'
      }
    });
  }
}