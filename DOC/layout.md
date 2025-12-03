⏺ 基於 Google Apps Script 官方文件，以下是 Google Slides 中所有可用的預定義版面配置（PredefinedLayout）：

  Google Slides PredefinedLayout 完整清單

  基本版面配置

- BLANK - 空白版面，沒有任何預設區塊
- TITLE_ONLY - 只有標題的版面配置
- CAPTION_ONLY - 只有底部說明文字的版面配置

  標題相關版面配置

- TITLE - 標題和副標題版面（首頁常用）
- TITLE_AND_BODY - 標題和內容版面（最常用）
- TITLE_AND_TWO_COLUMNS - 標題和兩欄內容版面
- ONE_COLUMN_TEXT - 單欄標題和內容版面

  章節分隔版面配置

- SECTION_HEADER - 章節標題版面
- SECTION_TITLE_AND_DESCRIPTION - 章節標題和描述版面（一側標題副標題，另一側描述）

  特殊內容版面配置

- MAIN_POINT - 重點強調版面
- BIG_NUMBER - 大數字標題版面

  系統值

- UNSUPPORTED - 不支援的版面配置

  使用範例

  // 建立不同版面配置的投影片
  const blankSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const titleSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);
  const contentSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
  const twoColumnSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_TWO_COLUMNS);
  const sectionSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.SECTION_HEADER);
  const mainPointSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.MAIN_POINT);

  重要注意事項

  1. 主題相依性：這些預定義版面配置可能不會存在於所有主題中，因為它們可能已被刪除或不屬於使用的主題
  2. 版面配置變化：每個版面配置上的預留位置可能已被修改
  3. 當前專案使用：目前的 md2slides 模組只使用了 SECTION_HEADER 和 TITLE_AND_BODY 兩種版面配置
