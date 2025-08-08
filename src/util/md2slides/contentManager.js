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
			addTitleToSlide(slide, info.title);

			// Add body content if it exists for TITLE_AND_BODY slides
			if (
				info.layout === "TITLE_AND_BODY" &&
				info.bodyItems &&
				info.bodyItems.length > 0
			) {
				addBodyContentToSlide(slide, info.bodyItems);
			}

			// Add speaker notes if they exist
			if (info.speakerNotes && info.speakerNotes.length > 0) {
				addSpeakerNotesToSlide(slide, info.speakerNotes);
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
 * @return {boolean} Success status
 */
function addTitleToSlide(slide, title) {
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
			const fontSize = getTitleFontSize(title);
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
				continue;
			}
		}
	}
	return "";
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
