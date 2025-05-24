// Progress bar module for Google Slides
function updateProgressBars(slides, requests) {
  const totalSlides = slides.length;
  const slideWidth = SlidesApp.getActivePresentation().getPageWidth();
  const maxWidth = slideWidth;
  const slideHeight = SlidesApp.getActivePresentation().getPageHeight();
  const height = progressBarHeight;
  const yPosition = slideHeight - height;
  const grayBackgroundColor = '#E0E0E0'; // 灰色背景色

  for (let i = 1; i < totalSlides; i++) {
    const slide = slides[i];
    const slideId = slide.getObjectId();

    // 刪除舊的 progress bar
    slide.getShapes().forEach(shape => {
      if (shape.getTitle && shape.getTitle() === 'PROGRESS' || shape.getTitle() === 'PROGRESS_BG') {
        requests.push({
          deleteObject: { objectId: shape.getObjectId() }
        });
      }
    });

    const progressRatio = i / (totalSlides - 1);
    const barWidth = maxWidth * progressRatio;
    const progressId = `progress_${slideId}_${newGuid()}`;
    const backgroundId = `progress_bg_${slideId}_${newGuid()}`;

    // 先創建灰色背景條
    requests.push(
      {
        createShape: {
          objectId: backgroundId,
          shapeType: 'RECTANGLE',
          elementProperties: {
            pageObjectId: slideId,
            size: {
              height: { magnitude: height, unit: 'PT' },
              width: { magnitude: maxWidth, unit: 'PT' } // 完整寬度
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
          objectId: backgroundId,
          shapeProperties: {
            shapeBackgroundFill: {
              solidFill: {
                color: {
                  rgbColor: hexToRgb(grayBackgroundColor)
                }
              }
            },
            outline: {
              weight: { magnitude: 0.1, unit: 'PT' },
              outlineFill: {
                solidFill: {
                  color: {
                    rgbColor: hexToRgb(grayBackgroundColor)
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
          objectId: backgroundId,
          title: 'PROGRESS_BG'
        }
      }
    );

    // 再創建進度條
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