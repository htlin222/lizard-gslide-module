/**
 * Content Manager Module
 *
 * Handles adding titles, body content, and speaker notes to slides
 */

/**
 * Adds content to all created slides (titles, body content, speaker notes)
 * @param {Array} createdSlides - Array of created slide objects
 * @return {boolean} Success status
 */
function addContentToSlides(createdSlides) {
	try {
		for (let i = 0; i < createdSlides.length; i++) {
			const slideObj = createdSlides[i];
			const slide = slideObj.slide;
			const info = slideObj.info;

			// Add title to all slides
			addTitleToSlide(slide, info.title, info.layout);

			// Add parent title (H2) if this is an H3 slide
			if (info.parentTitle && info.parentTitle.length > 0) {
				addParentTitleToSlide(slide, info.parentTitle);
			}

			// Add body content if it exists for TITLE_AND_BODY slides
			if (
				info.layout === "TITLE_AND_BODY" &&
				info.bodyItems &&
				info.bodyItems.length > 0
			) {
				addBodyContentToSlide(slide, info.bodyItems);
			}

			// Add code blocks if they exist
			if (info.codeBlocks && info.codeBlocks.length > 0) {
				addCodeBlocksToSlide(slide, info.codeBlocks);
			}

			// Add speaker notes if they exist
			if (info.speakerNotes && info.speakerNotes.length > 0) {
				addSpeakerNotesToSlide(slide, info.speakerNotes);
			}

			// Add footer items if they exist
			if (info.footerItems && info.footerItems.length > 0) {
				addFooterItemsToSlide(slide, info.footerItems);
			}
		}
		return true;
	} catch (error) {
		console.error(`Error adding content to slides: ${error.message}`);
		return false;
	}
}

/**
 * Adds title to a slide using multiple fallback approaches with font sizing
 * @param {Slide} slide - The slide to add title to
 * @param {string} title - The title text
 * @param {string} layout - The slide layout (SECTION_HEADER or TITLE_AND_BODY)
 * @return {boolean} Success status
 */
function addTitleToSlide(slide, title, layout) {
	const shapes = slide.getShapes();
	let titleAdded = false;
	let titleTextRange = null;

	// First pass: Look for TITLE placeholder
	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.TITLE) {
				titleTextRange = shape.getText();
				titleTextRange.setText(title);
				titleAdded = true;
				break;
			}
		} catch (e) {
			Logger.log(`Error checking placeholder type: ${e.message}`);
		}
	}

	// If title wasn't added, try another approach
	if (!titleAdded) {
		try {
			const titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
			if (titleShape) {
				titleTextRange = titleShape.getText();
				titleTextRange.setText(title);
				titleAdded = true;
			}
		} catch (e) {
			Logger.log(`Error getting title placeholder: ${e.message}`);
		}
	}

	// If title still wasn't added, use the first text box
	if (!titleAdded) {
		for (let j = 0; j < shapes.length; j++) {
			const shape = shapes[j];
			try {
				if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
					titleTextRange = shape.getText();
					titleTextRange.setText(title);
					titleAdded = true;
					break;
				}
			} catch (e) {
				Logger.log(`Error using text box for title: ${e.message}`);
			}
		}
	}

	// Apply title font sizing if title was added successfully
	if (titleAdded && titleTextRange) {
		try {
			// Use 36pt for H1 (SECTION_HEADER), calculated size for H2/H3
			const fontSize =
				layout === "SECTION_HEADER" ? 36 : getTitleFontSize(title);
			titleTextRange.getTextStyle().setFontSize(fontSize);
		} catch (e) {
			Logger.log(`Error applying title font size: ${e.message}`);
		}
	}

	return titleAdded;
}

/**
 * Adds body content to a slide using multiple fallback approaches
 * @param {Slide} slide - The slide to add body content to
 * @param {Array} bodyItems - Array of body text items
 * @return {boolean} Success status
 */
function addBodyContentToSlide(slide, bodyItems) {
	const shapes = slide.getShapes();
	let bodyContentAdded = false;

	// Method 1: Look for BODY placeholder by iterating through shapes
	bodyContentAdded = tryAddContentToBodyPlaceholder(shapes, bodyItems);

	// Method 2: Try getPlaceholder approach
	if (!bodyContentAdded) {
		bodyContentAdded = tryAddContentUsingGetPlaceholder(slide, bodyItems);
	}

	// Method 3: Find existing text box that's not the title
	if (!bodyContentAdded) {
		bodyContentAdded = tryAddContentToExistingTextBox(slide, shapes, bodyItems);
	}

	// Method 4: Create new text box if all else fails
	if (!bodyContentAdded) {
		bodyContentAdded = tryAddContentToNewTextBox(slide, bodyItems);
	}

	return bodyContentAdded;
}

/**
 * Tries to add content to BODY placeholder by iterating through shapes
 * @param {Array} shapes - Array of slide shapes
 * @param {Array} bodyItems - Body content items
 * @return {boolean} Success status
 */
function tryAddContentToBodyPlaceholder(shapes, bodyItems) {
	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
				addTextWithFontSizing(shape.getText(), bodyItems);
				return true;
			}
		} catch (e) {
			Logger.log(`Error checking for BODY placeholder: ${e.message}`);
		}
	}
	return false;
}

/**
 * Tries to add content using getPlaceholder method
 * @param {Slide} slide - The slide
 * @param {Array} bodyItems - Body content items
 * @return {boolean} Success status
 */
function tryAddContentUsingGetPlaceholder(slide, bodyItems) {
	try {
		const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
		if (bodyShape) {
			addTextWithFontSizing(bodyShape.getText(), bodyItems);
			return true;
		}
	} catch (e) {
		Logger.log(`Error getting body placeholder: ${e.message}`);
	}
	return false;
}

/**
 * Tries to add content to existing text box that's not the title
 * @param {Slide} slide - The slide
 * @param {Array} shapes - Array of slide shapes
 * @param {Array} bodyItems - Body content items
 * @return {boolean} Success status
 */
function tryAddContentToExistingTextBox(slide, shapes, bodyItems) {
	// Get title for comparison (assuming it's already set)
	const title = getTitleFromSlide(slide);

	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (
				shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX &&
				shape.getText().asString() !== title
			) {
				addTextWithFontSizing(shape.getText(), bodyItems);
				return true;
			}
		} catch (e) {
			Logger.log(`Error using text box for body: ${e.message}`);
		}
	}
	return false;
}

/**
 * Creates a new text box and adds content to it
 * @param {Slide} slide - The slide
 * @param {Array} bodyItems - Body content items
 * @return {boolean} Success status
 */
function tryAddContentToNewTextBox(slide, bodyItems) {
	try {
		const slideWidth = slide.getWidth();
		const slideHeight = slide.getHeight();

		const textBox = slide.insertTextBox(
			slideWidth * 0.1, // Left position
			slideHeight * 0.3, // Top position
			slideWidth * 0.8, // Width
			slideHeight * 0.6, // Height
		);

		addTextWithFontSizing(textBox.getText(), bodyItems);
		return true;
	} catch (e) {
		Logger.log(`Error creating new text box: ${e.message}`);
	}
	return false;
}

/**
 * Adds text content to a text range with automatic font sizing
 * @param {TextRange} textRange - The text range to add content to
 * @param {Array} bodyItems - Array of text items
 */
function addTextWithFontSizing(textRange, bodyItems) {
	textRange.clear();

	// Add each body item as a paragraph
	for (let k = 0; k < bodyItems.length; k++) {
		if (k === 0) {
			textRange.setText(bodyItems[k]);
		} else {
			textRange.appendParagraph(bodyItems[k]);
		}
	}

	// Calculate and apply optimal font size
	const allBodyText = bodyItems.join("\n");
	const fontSize = getFontSize(allBodyText);
	textRange.getTextStyle().setFontSize(fontSize);
}

/**
 * Gets the title text from a slide (for comparison purposes)
 * @param {Slide} slide - The slide to get title from
 * @return {string} The title text or empty string
 */
function getTitleFromSlide(slide) {
	try {
		const titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
		if (titleShape) {
			return titleShape.getText().asString();
		}
	} catch (e) {
		// Fallback: try to find title in shapes
		const shapes = slide.getShapes();
		for (let i = 0; i < shapes.length; i++) {
			try {
				if (
					shapes[i].getPlaceholderType() === SlidesApp.PlaceholderType.TITLE
				) {
					return shapes[i].getText().asString();
				}
			} catch (err) {
				// Skip this shape and continue to next one
			}
		}
	}
	return "";
}

/**
 * Adds code blocks to a slide as separate text shapes
 * @param {Slide} slide - The slide to add code blocks to
 * @param {Array} codeBlocks - Array of code block objects with language and content
 * @return {boolean} Success status
 */
function addCodeBlocksToSlide(slide, codeBlocks) {
	try {
		console.log(
			`addCodeBlocksToSlide called with ${codeBlocks.length} code blocks`,
		);

		// 計算佈局參數
		const maxWidth = 620; // 最大總寬度
		const gap = 10; // code blocks 之間的間距
		const numBlocks = codeBlocks.length;
		const totalGaps = (numBlocks - 1) * gap; // 總間距
		const blockWidth = (maxWidth - totalGaps) / numBlocks; // 每個 block 的寬度

		console.log(
			`Layout: ${numBlocks} blocks, total width: ${maxWidth}, gap: ${gap}, block width: ${blockWidth}`,
		);

		for (let i = 0; i < codeBlocks.length; i++) {
			const codeBlock = codeBlocks[i];
			console.log(
				`Processing code block ${i}: language=${codeBlock.language}, content="${codeBlock.content}"`,
			);

			// 計算行數和高度
			const lines = codeBlock.content.split("\n");
			const lineCount = lines.length;
			let height = 100; // 基礎高度（5行）
			if (lineCount > 5) {
				height += (lineCount - 5) * 30; // 每多一行加30
			}

			// 水平排列位置計算
			const x = 50 + i * (blockWidth + gap); // X位置：起始位置 + 索引 * (寬度 + 間距)
			const y = 250; // 所有 blocks 使用相同的 Y 位置（同一行）

			console.log(
				`Creating shape ${i}: x=${x}, y=${y}, width=${blockWidth}, height=${height}, lines=${lineCount}`,
			);

			// Create a rectangle shape for the code block
			const codeShape = slide.insertShape(
				SlidesApp.ShapeType.RECTANGLE,
				x,
				y,
				blockWidth,
				height,
			);

			console.log("Shape created successfully");

			// Set the code content first
			const textRange = codeShape.getText();
			textRange.setText(codeBlock.content);
			console.log("Text content set");

			// Apply basic formatting - 14號字，預設字體
			const textStyle = textRange.getTextStyle();
			textStyle.setFontSize(14); // 改為14號字
			textStyle.setForegroundColor("#000000");
			console.log("Text style applied - 14pt font");

			// 設定文字靠左對齊
			try {
				textRange
					.getParagraphStyle()
					.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
				console.log("Text alignment set to left");
			} catch (alignError) {
				console.log(`Text alignment error: ${alignError.message}`);
			}

			// 設定淺灰背景和白色邊框
			try {
				codeShape.getFill().setSolidFill("#f5f5f5"); // 淺灰色背景
				console.log("Light gray background applied");
			} catch (fillError) {
				console.log(`Fill error: ${fillError.message}`);
			}

			try {
				codeShape.getBorder().setWeight(1); // 細邊框
				codeShape.getBorder().getLineFill().setSolidFill("#FFFFFF"); // 白色邊框
				console.log("Gray background with white border applied");
			} catch (borderError) {
				console.log(`Border error: ${borderError.message}`);
			}

			// Set content alignment
			try {
				codeShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
				console.log("Content alignment set");
			} catch (alignError) {
				console.log(`Alignment error: ${alignError.message}`);
			}

			console.log(`Code block ${i} created successfully`);
		}

		console.log("All code blocks processed successfully");
		return true;
	} catch (e) {
		console.error(`Error adding code blocks to slide: ${e.message}`);
		console.error(`Error stack: ${e.stack}`);
		Logger.log(`Error adding code blocks to slide: ${e.message}`);
		return false;
	}
}

/**
 * Adds parent title (H2) to H3 slides in the style of insertStyledTitleBox
 * @param {Slide} slide - The slide to add parent title to
 * @param {string} parentTitle - The parent H2 title text
 * @return {boolean} Success status
 */
function addParentTitleToSlide(slide, parentTitle) {
	try {
		// Remove existing parent title if it exists
		const shapes = slide.getShapes();
		for (const shape of shapes) {
			if (
				shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX &&
				shape.getTitle() === "PREVIOUS_TITLE"
			) {
				shape.remove();
			}
		}

		// Insert styled text box at fixed position and size (same as insertStyledTitleBox)
		const parentTitleShape = slide.insertTextBox(parentTitle, 24, 18, 200, 25);
		parentTitleShape.setTitle("PREVIOUS_TITLE");

		// Apply font styling (same as insertStyledTitleBox)
		const textRange = parentTitleShape.getText();
		textRange.getTextStyle().setFontSize(12);
		textRange.getTextStyle().setForegroundColor("#888888");

		console.log(`Added parent title "${parentTitle}" to slide`);
		return true;
	} catch (e) {
		Logger.log(`Error adding parent title to slide: ${e.message}`);
		console.error(`Error adding parent title: ${e.message}`);
		return false;
	}
}

/**
 * Adds speaker notes to a slide
 * @param {Slide} slide - The slide to add notes to
 * @param {Array} speakerNotes - Array of speaker note strings
 * @return {boolean} Success status
 */
function addSpeakerNotesToSlide(slide, speakerNotes) {
	try {
		const speakerNotesText = speakerNotes.join("\n");
		slide
			.getNotesPage()
			.getSpeakerNotesShape()
			.getText()
			.setText(speakerNotesText);
		return true;
	} catch (e) {
		Logger.log(`Error adding speaker notes to slide: ${e.message}`);
		return false;
	}
}

/**
 * Adds footer items to a slide (in the bottom text area)
 * @param {Slide} slide - The slide to add footer items to
 * @param {Array} footerItems - Array of footer item strings
 * @return {boolean} Success status
 */
function addFooterItemsToSlide(slide, footerItems) {
	try {
		console.log(`Adding ${footerItems.length} footer items to slide`);

		// Try to find an existing footer placeholder first
		let footerAdded = false;

		// Method 1: Look for FOOTER placeholder
		footerAdded = tryAddFooterToPlaceholder(slide, footerItems);

		// Method 2: If no footer placeholder, try to find the bottom-most text box
		if (!footerAdded) {
			footerAdded = tryAddFooterToBottomTextBox(slide, footerItems);
		}

		// Method 3: If still no success, create a new text box at the bottom
		if (!footerAdded) {
			footerAdded = createFooterTextBox(slide, footerItems);
		}

		if (footerAdded) {
			console.log("Footer items added successfully");
		} else {
			console.log("Failed to add footer items");
		}

		return footerAdded;
	} catch (e) {
		Logger.log(`Error adding footer items to slide: ${e.message}`);
		console.error(`Error adding footer items: ${e.message}`);
		return false;
	}
}

/**
 * Tries to add footer items to a FOOTER placeholder
 * @param {Slide} slide - The slide
 * @param {Array} footerItems - Footer items to add
 * @return {boolean} Success status
 */
function tryAddFooterToPlaceholder(slide, footerItems) {
	try {
		const footerShape = slide.getPlaceholder(SlidesApp.PlaceholderType.FOOTER);
		if (footerShape && footerShape.asShape) {
			const textRange = footerShape.asShape().getText();
			const footerText = footerItems.join(" • ");
			textRange.setText(footerText);

			// Style the footer text
			textRange.getTextStyle().setFontSize(10);
			textRange.getTextStyle().setForegroundColor("#666666");

			console.log("Footer added to FOOTER placeholder");
			return true;
		}
	} catch (e) {
		console.log(`Error getting footer placeholder: ${e.message}`);
	}
	return false;
}

/**
 * Tries to add footer items to the bottom-most text box
 * @param {Slide} slide - The slide
 * @param {Array} footerItems - Footer items to add
 * @return {boolean} Success status
 */
function tryAddFooterToBottomTextBox(slide, footerItems) {
	try {
		const shapes = slide.getShapes();
		let bottomMostShape = null;
		let bottomMostY = 0;

		// Find the bottom-most text box (excluding title)
		const titleText = getTitleFromSlide(slide);

		for (let i = 0; i < shapes.length; i++) {
			const shape = shapes[i];
			try {
				if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
					const shapeText = shape.getText().asString().trim();
					// Skip if this is the title or already contains footer content
					if (shapeText === titleText || shapeText.includes("•")) {
						continue;
					}

					const top = shape.getTop();
					if (top > bottomMostY) {
						bottomMostY = top;
						bottomMostShape = shape;
					}
				}
			} catch (e) {
				// Skip this shape if we can't get its properties
			}
		}

		if (bottomMostShape) {
			const textRange = bottomMostShape.getText();
			const existingText = textRange.asString().trim();
			const footerText = footerItems.join(" • ");

			// Append footer items to existing content
			if (existingText) {
				textRange.setText(existingText + "\n\n" + footerText);
			} else {
				textRange.setText(footerText);
			}

			console.log("Footer added to bottom text box");
			return true;
		}
	} catch (e) {
		console.log(`Error adding footer to bottom text box: ${e.message}`);
	}
	return false;
}

/**
 * Creates a new text box at the bottom for footer items
 * @param {Slide} slide - The slide
 * @param {Array} footerItems - Footer items to add
 * @return {boolean} Success status
 */
function createFooterTextBox(slide, footerItems) {
	try {
		const slideWidth = slide.getWidth();
		const slideHeight = slide.getHeight();

		// Create footer text box at the bottom of the slide
		const footerBox = slide.insertTextBox(
			slideWidth * 0.05, // Left margin (5%)
			slideHeight * 0.9, // Near bottom (90% down)
			slideWidth * 0.9, // Width (90% of slide)
			slideHeight * 0.08, // Height (8% of slide)
		);

		// Set footer content
		const footerText = footerItems.join(" • ");
		const textRange = footerBox.getText();
		textRange.setText(footerText);

		// Style the footer text
		textRange.getTextStyle().setFontSize(10);
		textRange.getTextStyle().setForegroundColor("#666666");

		// Set paragraph alignment to center
		try {
			textRange
				.getParagraphStyle()
				.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
		} catch (alignError) {
			console.log(`Footer alignment error: ${alignError.message}`);
		}

		console.log("Footer text box created at bottom of slide");
		return true;
	} catch (e) {
		console.log(`Error creating footer text box: ${e.message}`);
		return false;
	}
}
