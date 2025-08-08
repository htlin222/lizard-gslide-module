/**
 * List Formatting Module
 *
 * Handles list type detection and applying appropriate Google Slides list presets
 */

/**
 * Applies list formatting to all TITLE_AND_BODY slides
 * @param {Array} createdSlides - Array of created slide objects
 * @return {boolean} Success status
 */
function applyListFormattingToSlides(createdSlides) {
	try {
		for (let i = 0; i < createdSlides.length; i++) {
			const slideObj = createdSlides[i];

			if (
				slideObj.info.layout === "TITLE_AND_BODY" &&
				slideObj.info.bodyItems &&
				slideObj.info.bodyItems.length > 0
			) {
				applyListFormattingToSlide(slideObj);
			}
		}
		return true;
	} catch (error) {
		console.error(`Error applying list formatting to slides: ${error.message}`);
		return false;
	}
}

/**
 * Applies list formatting to a single slide
 * @param {Object} slideObj - Slide object with slide and info properties
 * @return {boolean} Success status
 */
function applyListFormattingToSlide(slideObj) {
	try {
		const shapes = slideObj.slide.getShapes();
		let bodyFormattingApplied = false;

		// Determine the appropriate list preset based on list type
		const listPreset = getListPresetFromType(slideObj.info.listType);

		// Method 1: Look for BODY placeholder by iterating through shapes
		bodyFormattingApplied = tryApplyListFormattingToBodyPlaceholder(
			shapes,
			listPreset,
		);

		// Method 2: Try getPlaceholder approach
		if (!bodyFormattingApplied) {
			bodyFormattingApplied = tryApplyListFormattingUsingGetPlaceholder(
				slideObj.slide,
				listPreset,
			);
		}

		// Method 3: Find text boxes that aren't the title
		if (!bodyFormattingApplied) {
			bodyFormattingApplied = tryApplyListFormattingToTextBoxes(
				shapes,
				slideObj.info.title,
				listPreset,
			);
		}

		// Method 4: Fallback - manually add list markers
		if (!bodyFormattingApplied) {
			bodyFormattingApplied = tryApplyManualListFormatting(
				shapes,
				slideObj.info.title,
				slideObj.info.listType,
			);
		}

		return bodyFormattingApplied;
	} catch (error) {
		Logger.log(`Error applying list formatting to slide: ${error.message}`);
		return false;
	}
}

/**
 * Gets the appropriate Google Slides list preset for a given list type
 * @param {string} listType - "numbered", "lettered", or "bullet"
 * @return {SlidesApp.ListPreset} The appropriate list preset
 */
function getListPresetFromType(listType) {
	switch (listType) {
		case "numbered":
			return SlidesApp.ListPreset.NUMBERED_LIST_ARABIC_1;
		case "lettered":
			return SlidesApp.ListPreset.LETTERED_LIST_UPPER_ALPHA_PERIOD;
		case "bullet":
		default:
			return SlidesApp.ListPreset.DISC_CIRCLE_SQUARE;
	}
}

/**
 * Tries to apply list formatting to BODY placeholder by iterating through shapes
 * @param {Array} shapes - Array of slide shapes
 * @param {SlidesApp.ListPreset} listPreset - List preset to apply
 * @return {boolean} Success status
 */
function tryApplyListFormattingToBodyPlaceholder(shapes, listPreset) {
	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
				shape.getText().getListStyle().applyListPreset(listPreset);
				return true;
			}
		} catch (e) {
			Logger.log(
				`Error checking placeholder type for list formatting: ${e.message}`,
			);
		}
	}
	return false;
}

/**
 * Tries to apply list formatting using getPlaceholder method
 * @param {Slide} slide - The slide
 * @param {SlidesApp.ListPreset} listPreset - List preset to apply
 * @return {boolean} Success status
 */
function tryApplyListFormattingUsingGetPlaceholder(slide, listPreset) {
	try {
		const bodyShape = slide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
		if (bodyShape) {
			bodyShape.getText().getListStyle().applyListPreset(listPreset);
			return true;
		}
	} catch (e) {
		Logger.log(
			`Error getting body placeholder for list formatting: ${e.message}`,
		);
	}
	return false;
}

/**
 * Tries to apply list formatting to text boxes that aren't the title
 * @param {Array} shapes - Array of slide shapes
 * @param {string} title - Slide title (to exclude from formatting)
 * @param {SlidesApp.ListPreset} listPreset - List preset to apply
 * @return {boolean} Success status
 */
function tryApplyListFormattingToTextBoxes(shapes, title, listPreset) {
	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
				const text = shape.getText().asString().trim();
				// Skip if this is the title text box
				if (text !== "" && text !== title) {
					shape.getText().getListStyle().applyListPreset(listPreset);
					return true;
				}
			}
		} catch (e) {
			Logger.log(`Error applying list formatting to text box: ${e.message}`);
		}
	}
	return false;
}

/**
 * Tries to apply manual list formatting by adding list markers
 * @param {Array} shapes - Array of slide shapes
 * @param {string} title - Slide title (to exclude from formatting)
 * @param {string} listType - Type of list ("numbered", "lettered", "bullet")
 * @return {boolean} Success status
 */
function tryApplyManualListFormatting(shapes, title, listType) {
	for (let j = 0; j < shapes.length; j++) {
		const shape = shapes[j];
		try {
			if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
				const textRange = shape.getText();
				const text = textRange.asString().trim();

				// Skip if this is the title text box
				if (text !== "" && text !== title) {
					// Clear the text box
					textRange.clear();

					// Add each line with appropriate list marker
					const lines = text.split("\n");
					for (let k = 0; k < lines.length; k++) {
						const line = lines[k].trim();
						if (line !== "") {
							const marker = getListMarkerForIndex(k, listType);
							if (k === 0) {
								textRange.setText(marker + " " + line);
							} else {
								textRange.appendParagraph(marker + " " + line);
							}
						}
					}

					return true;
				}
			}
		} catch (e) {
			Logger.log(`Error applying manual list formatting: ${e.message}`);
		}
	}
	return false;
}

/**
 * Gets the appropriate list marker for a given index and list type
 * @param {number} index - Zero-based index of the list item
 * @param {string} listType - Type of list ("numbered", "lettered", "bullet")
 * @return {string} The list marker
 */
function getListMarkerForIndex(index, listType) {
	switch (listType) {
		case "numbered":
			return (index + 1).toString() + ".";
		case "lettered":
			// Convert index to letter (0=A, 1=B, etc.)
			return String.fromCharCode(65 + (index % 26)) + ".";
		case "bullet":
		default:
			return "•";
	}
}

/**
 * Validates list type and returns a safe default if invalid
 * @param {string} listType - The list type to validate
 * @return {string} Valid list type ("numbered", "lettered", or "bullet")
 */
function validateListType(listType) {
	const validTypes = ["numbered", "lettered", "bullet"];
	return validTypes.includes(listType) ? listType : "bullet";
}
