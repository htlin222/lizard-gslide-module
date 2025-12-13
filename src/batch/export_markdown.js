// 📝 EXPORT MARKDOWN MODULE - Extract slide content as clean markdown
/**
 * Extracts meaningful content from slides and exports as Marp-like markdown.
 * Filters out auto-generated elements: page numbers, tabs, sections labels,
 * progress bars, index items, outlines, watermarks, and footnotes.
 */

/**
 * ID prefixes used by auto-generated elements that should be excluded
 */
const GENERATED_ELEMENT_PREFIXES = [
	"progress_bg_", // Progress bar background
	"progress_", // Progress bar
	"page_num_", // Page numbers
	"tab_bg_", // Tab navigation background
	"tab_", // Tab navigation items
	"tab_line_", // Tab navigation line
	"label_", // Section labels
	"sections_", // Section boxes
	"outline_", // Outline content
	"index_", // Index items
	"obj_", // Title footnotes (obj_slideId_timestamp_guid)
	"watermark_", // Watermarks
];

/**
 * Text patterns that indicate auto-generated content
 */
const GENERATED_TEXT_PATTERNS = [
	/^\d+\s*\/\s*\d+$/, // Page numbers like "1 / 10"
	/^Section:\s*\d+$/, // Section labels like "Section: 1"
	/^ⓒ\s*/, // Copyright watermarks
];

/**
 * Check if an element ID indicates it's auto-generated
 * @param {string} elementId - The object ID of the element
 * @returns {boolean} True if the element is auto-generated
 */
function isGeneratedElementById(elementId) {
	if (!elementId) return false;
	return GENERATED_ELEMENT_PREFIXES.some((prefix) =>
		elementId.startsWith(prefix),
	);
}

/**
 * Check if text content matches auto-generated patterns
 * @param {string} text - The text content to check
 * @returns {boolean} True if the text matches generated patterns
 */
function isGeneratedTextContent(text) {
	if (!text) return false;
	const trimmed = text.trim();
	return GENERATED_TEXT_PATTERNS.some((pattern) => pattern.test(trimmed));
}

/**
 * Extract text from a shape with bullet list detection
 * Uses SlidesApp's ListStyle API to detect bulleted paragraphs
 * @param {GoogleAppsScript.Slides.Shape} shape - The shape to extract text from
 * @returns {string} Formatted text with markdown bullet prefixes where applicable
 */
function extractTextWithBullets(shape) {
	const textRange = shape.getText();
	const paragraphs = textRange.getParagraphs();
	const lines = [];

	for (const paragraph of paragraphs) {
		const paragraphRange = paragraph.getRange();
		const paragraphText = paragraphRange.asString().replace(/\n$/, ""); // Remove trailing newline

		// Skip empty paragraphs
		if (!paragraphText.trim()) {
			continue;
		}

		// Check if this paragraph is in a list
		const listStyle = paragraphRange.getListStyle();
		const isInList = listStyle.isInList();

		if (isInList === true) {
			// Get nesting level for indentation (0-based)
			const nestingLevel = listStyle.getNestingLevel() || 0;
			const indent = "  ".repeat(nestingLevel); // 2 spaces per level
			lines.push(`${indent}- ${paragraphText.trim()}`);
		} else {
			lines.push(paragraphText.trim());
		}
	}

	return lines.join("\n");
}

/**
 * Extract clean text content from a slide, excluding generated elements
 * @param {GoogleAppsScript.Slides.Slide} slide - The slide to extract from
 * @returns {Object} Object with title and body text arrays
 */
function extractSlideContent(slide) {
	const layout = slide.getLayout();
	const layoutName = layout ? layout.getLayoutName() : "";
	const pageElements = slide.getPageElements();

	const result = {
		layoutName: layoutName,
		title: "",
		body: [],
		images: [],
		speakerNotes: "",
	};

	// Get speaker notes if available
	const notesPage = slide.getNotesPage();
	if (notesPage) {
		const notesShapes = notesPage.getShapes();
		for (const shape of notesShapes) {
			const placeholderType = shape.getPlaceholderType();
			if (placeholderType === SlidesApp.PlaceholderType.BODY) {
				const notesText = shape.getText().asString().trim();
				if (notesText) {
					result.speakerNotes = notesText;
				}
			}
		}
	}

	// Process page elements
	for (const element of pageElements) {
		const elementId = element.getObjectId();

		// Skip if element ID indicates it's auto-generated
		if (isGeneratedElementById(elementId)) {
			continue;
		}

		if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
			const shape = element.asShape();
			const rawText = shape.getText().asString().trim();

			// Skip empty text or auto-generated text patterns
			if (!rawText || isGeneratedTextContent(rawText)) {
				continue;
			}

			// Check if this is a placeholder (title, subtitle, body)
			const placeholderType = shape.getPlaceholderType();

			if (
				placeholderType === SlidesApp.PlaceholderType.TITLE ||
				placeholderType === SlidesApp.PlaceholderType.CENTERED_TITLE
			) {
				// Titles don't need bullet detection
				result.title = rawText;
			} else if (
				placeholderType === SlidesApp.PlaceholderType.SUBTITLE ||
				placeholderType === SlidesApp.PlaceholderType.BODY
			) {
				// Body content - extract with bullet detection
				const formattedText = extractTextWithBullets(shape);
				if (formattedText) {
					result.body.push(formattedText);
				}
			} else {
				// For non-placeholder shapes, check position to determine if it's likely content
				// Skip very small shapes or shapes at edges (likely UI elements)
				const transform = element.getTransform();
				const translateX = transform.getTranslateX();
				const translateY = transform.getTranslateY();

				// Skip elements at top edge (likely tabs) or bottom edge (likely page numbers)
				if (translateY < 20 || translateY > 360) {
					continue;
				}

				// Skip elements at right edge (likely footnotes)
				if (translateX > 680) {
					continue;
				}

				// Include other content shapes with bullet detection
				const formattedText = extractTextWithBullets(shape);
				if (formattedText) {
					result.body.push(formattedText);
				}
			}
		} else if (
			element.getPageElementType() === SlidesApp.PageElementType.TABLE
		) {
			// Extract table content
			const table = element.asTable();
			const tableText = extractTableContent(table);
			if (tableText) {
				result.body.push(tableText);
			}
		} else if (
			element.getPageElementType() === SlidesApp.PageElementType.IMAGE
		) {
			// Extract image
			const image = element.asImage();
			const imageUrl = image.getContentUrl();
			if (imageUrl) {
				const altText =
					image.getDescription() || image.getTitle() || "slide image";
				result.images.push({
					url: imageUrl,
					alt: altText,
				});
			}
		}
	}

	return result;
}

/**
 * Extract content from a table as markdown
 * @param {GoogleAppsScript.Slides.Table} table - The table to extract from
 * @returns {string} Markdown formatted table content
 */
function extractTableContent(table) {
	const numRows = table.getNumRows();
	const numCols = table.getNumColumns();

	if (numRows === 0 || numCols === 0) return "";

	const rows = [];

	for (let r = 0; r < numRows; r++) {
		const row = [];
		for (let c = 0; c < numCols; c++) {
			const cell = table.getCell(r, c);
			const cellText = cell.getText().asString().trim().replace(/\|/g, "\\|");
			row.push(cellText);
		}
		rows.push(`| ${row.join(" | ")} |`);

		// Add header separator after first row
		if (r === 0) {
			rows.push(`|${" --- |".repeat(numCols)}`);
		}
	}

	return rows.join("\n");
}

/**
 * Get or create a folder by name under a parent folder
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - Parent folder
 * @param {string} folderName - Name of folder to get or create
 * @returns {GoogleAppsScript.Drive.Folder} The folder
 */
function getOrCreateFolder(parentFolder, folderName) {
	const folders = parentFolder.getFoldersByName(folderName);
	if (folders.hasNext()) {
		return folders.next();
	}
	return parentFolder.createFolder(folderName);
}

/**
 * Clear all files in a folder (move to trash)
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder to clear
 */
function clearFolder(folder) {
	const files = folder.getFiles();
	while (files.hasNext()) {
		const file = files.next();
		file.setTrashed(true);
	}
}

/**
 * Convert a title string to a URL/filename-safe slug
 * @param {string} title - The title to convert
 * @returns {string} A slug-formatted string
 */
function slugify(title) {
	if (!title) return "";

	return (
		title
			.toLowerCase()
			.trim()
			// Replace Chinese/CJK characters with pinyin or keep as-is (just keep them for now)
			// Remove or replace problematic characters for filenames
			.replace(/[\/\\:*?"<>|]/g, "") // Remove: / \ : * ? " < > |
			.replace(/[&]/g, "-and-") // Replace & with -and-
			.replace(/[@]/g, "-at-") // Replace @ with -at-
			.replace(/[#]/g, "") // Remove #
			.replace(/[%]/g, "") // Remove %
			.replace(/\s+/g, "-") // Replace spaces with hyphens
			.replace(/[_]+/g, "-") // Replace underscores with hyphens
			.replace(/-+/g, "-") // Replace multiple hyphens with single
			.replace(/^-+|-+$/g, "")
	); // Remove leading/trailing hyphens
}

/**
 * Save image to Drive
 * @param {string} contentUrl - The image content URL from Slides
 * @param {GoogleAppsScript.Drive.Folder} folder - The folder to save to
 * @param {string} filename - The filename for the image
 * @returns {boolean} True if saved successfully, false otherwise
 */
function saveImageToDrive(contentUrl, folder, filename) {
	try {
		const response = UrlFetchApp.fetch(contentUrl, {
			muteHttpExceptions: true,
			followRedirects: true,
		});

		const responseCode = response.getResponseCode();
		if (responseCode !== 200) {
			Logger.log(`HTTP ${responseCode} for ${filename}`);
			return false;
		}

		const blob = response.getBlob().setName(filename);
		folder.createFile(blob);
		return true;
	} catch (e) {
		Logger.log(`Failed to save image ${filename}: ${e.message}`);
		return false;
	}
}

/**
 * Generate Quarto YAML front matter
 * @param {string} title - Presentation title
 * @returns {string} YAML front matter for Quarto
 */
function generateQuartoFrontMatter(title) {
	const today = Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		"yyyy-MM-dd",
	);

	return `---
title: "${title}"
author: "林協霆"
date: "${today}"
format: html
---`;
}

/**
 * Convert slides to Marp-like markdown format
 * @param {boolean} saveImagesToDrive - Whether to save images to Drive
 * @returns {Object} Object with markdown content and presentationFolder
 */
function exportSlidesToMarkdown(saveImagesToDrive) {
	const presentation = SlidesApp.getActivePresentation();
	const slides = presentation.getSlides();
	const presentationTitle = presentation.getName();
	const presentationId = presentation.getId();

	// Set up assets folder if saving images to Drive
	let assetsFolder = null;
	let presentationFolder = null;

	if (saveImagesToDrive) {
		// Get the presentation's parent folder
		const presentationFile = DriveApp.getFileById(presentationId);
		const parentFolders = presentationFile.getParents();
		const parentFolder = parentFolders.hasNext()
			? parentFolders.next()
			: DriveApp.getRootFolder();

		// Create presentation_name/assets/ folder structure
		presentationFolder = getOrCreateFolder(parentFolder, presentationTitle);
		assetsFolder = getOrCreateFolder(presentationFolder, "assets");

		// Clear existing files in assets folder
		clearFolder(assetsFolder);
	}

	const markdownParts = [];

	// Add YAML front matter (Marp style)
	markdownParts.push("---");
	markdownParts.push("marp: true");
	markdownParts.push(`title: ${presentationTitle}`);
	markdownParts.push("---");
	markdownParts.push("");

	for (let index = 0; index < slides.length; index++) {
		const slide = slides[index];
		const content = extractSlideContent(slide);

		// Add slide separator (except for first slide, and not needed if slide has title)
		if (index > 0) {
			markdownParts.push("");
			// Only add --- separator if slide has no title (headings act as separators)
			if (!content.title) {
				markdownParts.push("---");
				markdownParts.push("");
			}
		}

		// Determine heading level based on layout
		if (content.layoutName === "SECTION_HEADER") {
			// Section header slides get H1
			if (content.title) {
				markdownParts.push(`# ${content.title}`);
			}
			// Add body content if any
			for (const text of content.body) {
				markdownParts.push("");
				markdownParts.push(text);
			}
		} else if (content.layoutName === "TITLE") {
			// Title slide (first slide typically)
			if (content.title) {
				markdownParts.push(`# ${content.title}`);
			}
			for (const text of content.body) {
				markdownParts.push("");
				markdownParts.push(text);
			}
		} else {
			// Regular slides get H2
			if (content.title) {
				markdownParts.push(`## ${content.title}`);
			}

			// Add body content
			for (const text of content.body) {
				markdownParts.push("");
				markdownParts.push(text);
			}
		}

		// Add images at the end of slide content
		if (content.images.length > 0) {
			markdownParts.push("");
			for (let imgIndex = 0; imgIndex < content.images.length; imgIndex++) {
				const img = content.images[imgIndex];
				let imageUrl = img.url;

				// Save to Drive if enabled
				if (saveImagesToDrive && assetsFolder) {
					const slideNum = String(index + 1).padStart(2, "0");
					const titleSlug = slugify(content.title);
					const filename = titleSlug
						? `slide_${slideNum}_${titleSlug}_img_${imgIndex + 1}.png`
						: `slide_${slideNum}_img_${imgIndex + 1}.png`;
					const saved = saveImageToDrive(img.url, assetsFolder, filename);
					if (saved) {
						// Use relative path for markdown
						imageUrl = `./assets/${filename}`;
					}
				}

				markdownParts.push(`![${img.alt}](${imageUrl})`);
			}
		}

		// Add speaker notes as HTML comment if present
		if (content.speakerNotes) {
			markdownParts.push("");
			markdownParts.push("<!--");
			markdownParts.push("Speaker notes:");
			markdownParts.push(content.speakerNotes);
			markdownParts.push("-->");
		}
	}

	const marpMarkdown = markdownParts.join("\n");

	// Save .md and .qmd files if saving to Drive
	if (saveImagesToDrive && presentationFolder) {
		const fileSlug = slugify(presentationTitle) || "slides";

		// Generate Quarto markdown (replace Marp front matter with Quarto)
		const quartoFrontMatter = generateQuartoFrontMatter(presentationTitle);
		const contentWithoutFrontMatter = marpMarkdown.replace(
			/^---[\s\S]*?---\n*/,
			"",
		);
		const quartoMarkdown = `${quartoFrontMatter}\n\n${contentWithoutFrontMatter}`;

		// Remove old .md and .qmd files in presentation folder
		const existingFiles = presentationFolder.getFiles();
		while (existingFiles.hasNext()) {
			const file = existingFiles.next();
			const fileName = file.getName();
			if (fileName.endsWith(".md") || fileName.endsWith(".qmd")) {
				file.setTrashed(true);
			}
		}

		// Save .md file (Marp format)
		presentationFolder.createFile(`${fileSlug}.md`, marpMarkdown, "text/plain");

		// Save .qmd file (Quarto format)
		presentationFolder.createFile(
			`${fileSlug}.qmd`,
			quartoMarkdown,
			"text/plain",
		);

		// Return markdown and folder URL for Drive exports
		return {
			markdown: marpMarkdown,
			folderUrl: presentationFolder.getUrl(),
			folderName: presentationTitle,
		};
	}

	return marpMarkdown;
}

/**
 * Show dialog with exported markdown content (temporary image URLs)
 */
function showExportMarkdownDialog() {
	const markdown = exportSlidesToMarkdown(false);

	const htmlTemplate = HtmlService.createTemplateFromFile(
		"src/components/export-markdown-dialog.html",
	);
	htmlTemplate.markdownContent = markdown;

	const html = htmlTemplate
		.evaluate()
		.setWidth(800)
		.setHeight(600)
		.setTitle("Export to Markdown");

	SlidesApp.getUi().showModalDialog(html, "📝 Export to Markdown");
}

/**
 * Show dialog with exported markdown content (images saved to Drive)
 */
function showExportMarkdownWithImagesDialog() {
	const ui = SlidesApp.getUi();

	// Confirm with user since this will create files
	const response = ui.alert(
		"匯出 Markdown 與圖片",
		"這將會在簡報所在資料夾建立:\n\n" +
			"📁 [簡報名稱]/\n" +
			"    📄 [slug].md (Marp 格式)\n" +
			"    📄 [slug].qmd (Quarto 格式)\n" +
			"    📁 assets/\n" +
			"        🖼️ slide_01_title_img_1.png\n" +
			"        ...\n\n" +
			"⚠️ 舊的 .md/.qmd 及 assets 內檔案會被清除\n\n" +
			"要繼續嗎？",
		ui.ButtonSet.YES_NO,
	);

	if (response !== ui.Button.YES) {
		return;
	}

	try {
		const result = exportSlidesToMarkdown(true);

		const htmlTemplate = HtmlService.createTemplateFromFile(
			"src/components/export-markdown-dialog.html",
		);
		htmlTemplate.markdownContent = result.markdown;
		htmlTemplate.folderUrl = result.folderUrl;
		htmlTemplate.folderName = result.folderName;

		const html = htmlTemplate
			.evaluate()
			.setWidth(800)
			.setHeight(600)
			.setTitle("Export to Markdown (with Drive images)");

		ui.showModalDialog(html, "📝 匯出完成 (已存至 Drive)");
	} catch (e) {
		ui.alert("匯出失敗", `發生錯誤: ${e.message}`, ui.ButtonSet.OK);
	}
}
