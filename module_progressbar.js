// Progress bar module for Google Slides
function updateProgressBars(slides, requests) {
  const totalSlides = slides.length;
  const maxWidth = slideWidth;
  const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
const slideHeight = SlidesApp.getActivePresentation().getPageHeight();
  const height = progressBarHeight;
  const yPosition = slideHeight - height;

  for (let i = 1; i < totalSlides; i++) {
    const slide = slides[i];
    const slideId = slide.getObjectId();

    // 刪除舊的 progress bar
    slide.getShapes().forEach(shape => {
      if (shape.getTitle && shape.getTitle() === 'PROGRESS') {
        requests.push({
          deleteObject: { objectId: shape.getObjectId() }
        });
      }
    });

    const progressRatio = i / (totalSlides - 1);
    const barWidth = maxWidth * progressRatio;
    const progressId = `progress_${slideId}_${newGuid()}`;

    requests.push(
      {
        createShape: {
          objectId: progressId,
          shapeType: 'RECTANGLE',
          elementProperties: {
            pageObjectId: slideId,
            size: {
              height: { magnitude: height, unit: 'PT' },
              width: { magnitude: barWidth, unit: 'PT' }
            },
            transform: {
              scaleX: 1,
              scaleY: 1,
              translateX: 0,
              translateY: yPosition,
              unit: 'PT'
            }
          }
        }
      },
      {
        updateShapeProperties: {
          objectId: progressId,
          shapeProperties: {
            shapeBackgroundFill: solidFill(main_color),
            outline: {
              weight: { magnitude: 0.1, unit: 'PT' }, // ✅ 最小有效值
              outlineFill: {
                solidFill: {
                  color: {
                    rgbColor: hexToRgb(main_color) // ✅ 與背景同色
                  }
                }
              }
            }
          },
          fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color'
        }
      },
      {
        updatePageElementAltText: {
          objectId: progressId,
          title: 'PROGRESS'
        }
      }
    );
  }
}

function newGuid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}