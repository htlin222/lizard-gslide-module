/**
 * Download Selected Image Utilities
 *
 * Google Apps Script runs server-side, so direct browser downloads aren't possible.
 * This module provides two approaches:
 * 1. Save to Google Drive (recommended)
 * 2. Open in new tab for manual download
 */

/**
 * Check if the current selection is an image
 * @returns {GoogleAppsScript.Slides.Image|null} The selected image or null
 */
function getSelectedImage() {
	const selection = SlidesApp.getActivePresentation().getSelection();
	const selectionType = selection.getSelectionType();

	if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
		return null;
	}

	const pageElements = selection.getPageElementRange().getPageElements();

	if (pageElements.length !== 1) {
		return null;
	}

	const element = pageElements[0];

	if (element.getPageElementType() !== SlidesApp.PageElementType.IMAGE) {
		return null;
	}

	return element.asImage();
}

/**
 * Save the selected image to Google Drive
 * Creates a PNG file in the user's Drive root folder
 */
function saveSelectedImageToDrive() {
	const image = getSelectedImage();

	if (!image) {
		SlidesApp.getUi().alert(
			"未選取圖片",
			"請先選取一張圖片再執行此功能。",
			SlidesApp.getUi().ButtonSet.OK,
		);
		return;
	}

	try {
		const contentUrl = image.getContentUrl();

		if (!contentUrl) {
			SlidesApp.getUi().alert(
				"無法取得圖片",
				"無法取得此圖片的內容網址。",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		// Fetch the image blob
		const response = UrlFetchApp.fetch(contentUrl);
		const blob = response.getBlob();

		// Generate filename with timestamp
		const timestamp = Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			"yyyyMMdd_HHmmss",
		);
		const filename = `slide_image_${timestamp}.png`;

		// Set blob name and type
		blob.setName(filename);

		// Create file in Drive
		const file = DriveApp.createFile(blob);

		// Show success message with file link
		const fileUrl = file.getUrl();
		const html = HtmlService.createHtmlOutput(
			`<p>圖片已儲存至 Google Drive！</p><p><a href="${fileUrl}" target="_blank">點此開啟檔案</a></p><p>檔案名稱: ${filename}</p>`,
		)
			.setWidth(300)
			.setHeight(150);

		SlidesApp.getUi().showModalDialog(html, "儲存成功");
	} catch (e) {
		SlidesApp.getUi().alert(
			"錯誤",
			`儲存圖片時發生錯誤: ${e.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Open the selected image in a new tab for manual download
 * User can right-click and "Save As" to download
 */
function openSelectedImageInNewTab() {
	const image = getSelectedImage();

	if (!image) {
		SlidesApp.getUi().alert(
			"未選取圖片",
			"請先選取一張圖片再執行此功能。",
			SlidesApp.getUi().ButtonSet.OK,
		);
		return;
	}

	try {
		let contentUrl = image.getContentUrl();

		if (!contentUrl) {
			// Try source URL as fallback
			contentUrl = image.getSourceUrl();
		}

		if (!contentUrl) {
			SlidesApp.getUi().alert(
				"無法取得圖片",
				"無法取得此圖片的網址。",
				SlidesApp.getUi().ButtonSet.OK,
			);
			return;
		}

		// Show dialog with link to open image
		const html = HtmlService.createHtmlOutput(
			`<p>點擊下方連結在新分頁開啟圖片，</p><p>然後右鍵選擇「另存圖片」即可下載。</p><p><a href="${contentUrl}" target="_blank" style="font-size: 16px;">開啟圖片</a></p>`,
		)
			.setWidth(350)
			.setHeight(150);

		SlidesApp.getUi().showModalDialog(html, "下載圖片");
	} catch (e) {
		SlidesApp.getUi().alert(
			"錯誤",
			`取得圖片時發生錯誤: ${e.message}`,
			SlidesApp.getUi().ButtonSet.OK,
		);
	}
}

/**
 * Save selected image to a specific Drive folder
 * @param {string} folderId - The Drive folder ID to save to
 */
function saveSelectedImageToFolder(folderId) {
	const image = getSelectedImage();

	if (!image) {
		return { success: false, error: "未選取圖片" };
	}

	try {
		const contentUrl = image.getContentUrl();
		const response = UrlFetchApp.fetch(contentUrl);
		const blob = response.getBlob();

		const timestamp = Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			"yyyyMMdd_HHmmss",
		);
		const filename = `slide_image_${timestamp}.png`;
		blob.setName(filename);

		const folder = DriveApp.getFolderById(folderId);
		const file = folder.createFile(blob);

		return {
			success: true,
			fileUrl: file.getUrl(),
			filename: filename,
		};
	} catch (e) {
		return { success: false, error: e.message };
	}
}

/**
 * Get image info for debugging
 */
function showSelectedImageInfo() {
	const image = getSelectedImage();

	if (!image) {
		SlidesApp.getUi().alert(
			"未選取圖片",
			"請先選取一張圖片再執行此功能。",
			SlidesApp.getUi().ButtonSet.OK,
		);
		return;
	}

	const info = {
		"Content URL": image.getContentUrl() || "(無)",
		"Source URL": image.getSourceUrl() || "(無)",
		Width: `${image.getWidth()} pt`,
		Height: `${image.getHeight()} pt`,
		Title: image.getTitle() || "(無)",
		Description: image.getDescription() || "(無)",
	};

	const infoText = Object.keys(info)
		.map((key) => `${key}: ${info[key]}`)
		.join("\n");

	SlidesApp.getUi().alert("圖片資訊", infoText, SlidesApp.getUi().ButtonSet.OK);
}
