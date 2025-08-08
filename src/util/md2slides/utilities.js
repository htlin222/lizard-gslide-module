/**
 * Utilities Module
 *
 * Helper functions shared across md2slides modules
 */

/**
 * Safely gets text content from a text range
 * @param {TextRange} textRange - The text range to get content from
 * @return {string} The text content or empty string if error
 */
function safeGetTextContent(textRange) {
	try {
		return textRange.asString();
	} catch (error) {
		Logger.log(`Error getting text content: ${error.message}`);
		return "";
	}
}

/**
 * Safely sets text content to a text range
 * @param {TextRange} textRange - The text range to set content to
 * @param {string} text - The text to set
 * @return {boolean} Success status
 */
function safeSetTextContent(textRange, text) {
	try {
		textRange.setText(text);
		return true;
	} catch (error) {
		Logger.log(`Error setting text content: ${error.message}`);
		return false;
	}
}

/**
 * Safely appends a paragraph to a text range
 * @param {TextRange} textRange - The text range to append to
 * @param {string} text - The text to append
 * @return {boolean} Success status
 */
function safeAppendParagraph(textRange, text) {
	try {
		textRange.appendParagraph(text);
		return true;
	} catch (error) {
		Logger.log(`Error appending paragraph: ${error.message}`);
		return false;
	}
}

/**
 * Safely gets a placeholder from a slide
 * @param {Slide} slide - The slide
 * @param {SlidesApp.PlaceholderType} placeholderType - The placeholder type
 * @return {Shape|null} The placeholder shape or null if not found
 */
function safeGetPlaceholder(slide, placeholderType) {
	try {
		return slide.getPlaceholder(placeholderType);
	} catch (error) {
		Logger.log(`Error getting placeholder: ${error.message}`);
		return null;
	}
}

/**
 * Safely checks if a shape has a specific placeholder type
 * @param {Shape} shape - The shape to check
 * @param {SlidesApp.PlaceholderType} placeholderType - The placeholder type to check for
 * @return {boolean} True if shape has the placeholder type
 */
function safeCheckPlaceholderType(shape, placeholderType) {
	try {
		return shape.getPlaceholderType() === placeholderType;
	} catch (error) {
		// Shape might not have a placeholder type
		return false;
	}
}

/**
 * Safely gets the shape type of a shape
 * @param {Shape} shape - The shape to check
 * @return {SlidesApp.ShapeType|null} The shape type or null if error
 */
function safeGetShapeType(shape) {
	try {
		return shape.getShapeType();
	} catch (error) {
		Logger.log(`Error getting shape type: ${error.message}`);
		return null;
	}
}

/**
 * Checks if a slide layout supports body content
 * @param {string} layout - The slide layout ("SECTION_HEADER" or "TITLE_AND_BODY")
 * @return {boolean} True if layout supports body content
 */
function layoutSupportsBodyContent(layout) {
	return layout === "TITLE_AND_BODY";
}

/**
 * Validates a slide structure object
 * @param {Object} slideInfo - The slide info object to validate
 * @return {boolean} True if valid
 */
function validateSlideStructure(slideInfo) {
	if (!slideInfo || typeof slideInfo !== "object") {
		return false;
	}

	// Check required properties
	if (!slideInfo.layout || !slideInfo.title) {
		return false;
	}

	// Check valid layout
	const validLayouts = ["SECTION_HEADER", "TITLE_AND_BODY"];
	if (!validLayouts.includes(slideInfo.layout)) {
		return false;
	}

	// Ensure arrays exist
	if (!Array.isArray(slideInfo.bodyItems)) {
		slideInfo.bodyItems = [];
	}
	if (!Array.isArray(slideInfo.speakerNotes)) {
		slideInfo.speakerNotes = [];
	}

	return true;
}

/**
 * Calculates slide dimensions for positioning elements
 * @param {Slide} slide - The slide to get dimensions for
 * @return {Object} Object with width, height, and common positioning values
 */
function getSlideLayoutDimensions(slide) {
	try {
		const width = slide.getWidth();
		const height = slide.getHeight();

		return {
			width,
			height,
			centerX: width / 2,
			centerY: height / 2,
			// Common positioning values
			marginLeft: width * 0.1,
			marginTop: height * 0.3,
			contentWidth: width * 0.8,
			contentHeight: height * 0.6,
		};
	} catch (error) {
		Logger.log(`Error getting slide dimensions: ${error.message}`);
		// Return default values
		return {
			width: 720,
			height: 405,
			centerX: 360,
			centerY: 202.5,
			marginLeft: 72,
			marginTop: 121.5,
			contentWidth: 576,
			contentHeight: 243,
		};
	}
}

/**
 * Cleans up markdown text by removing extra whitespace and formatting
 * @param {string} text - The text to clean
 * @return {string} Cleaned text
 */
function cleanMarkdownText(text) {
	if (typeof text !== "string") {
		return "";
	}

	return text
		.trim()
		.replace(/\r\n/g, "\n") // Normalize line endings
		.replace(/\r/g, "\n") // Convert remaining carriage returns
		.replace(/\n{3,}/g, "\n\n"); // Limit consecutive newlines to 2
}

/**
 * Splits text into chunks based on character limits
 * @param {string} text - The text to split
 * @param {number} maxChars - Maximum characters per chunk
 * @return {Array<string>} Array of text chunks
 */
function splitTextIntoChunks(text, maxChars = 1000) {
	if (!text || text.length <= maxChars) {
		return [text];
	}

	const chunks = [];
	const sentences = text.split(/(?<=[.!?])\s+/);
	let currentChunk = "";

	for (const sentence of sentences) {
		if ((currentChunk + sentence).length <= maxChars) {
			currentChunk += (currentChunk ? " " : "") + sentence;
		} else {
			if (currentChunk) {
				chunks.push(currentChunk);
			}
			currentChunk = sentence;
		}
	}

	if (currentChunk) {
		chunks.push(currentChunk);
	}

	return chunks;
}

/**
 * Escapes special characters in text for safe display
 * @param {string} text - The text to escape
 * @return {string} Escaped text
 */
function escapeTextForDisplay(text) {
	if (typeof text !== "string") {
		return "";
	}

	return text
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&#x27;");
}

/**
 * Generates a unique identifier for debugging purposes
 * @return {string} A unique identifier
 */
function generateUniqueId() {
	return `md2slides_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Logs debug information with consistent formatting
 * @param {string} module - The module name
 * @param {string} operation - The operation being performed
 * @param {string} message - The debug message
 * @param {*} data - Optional data to log
 */
function debugLog(module, operation, message, data = null) {
	const timestamp = new Date().toISOString();
	let logMessage = `[${timestamp}] [${module}] [${operation}] ${message}`;

	if (data !== null) {
		logMessage += ` | Data: ${JSON.stringify(data)}`;
	}

	Logger.log(logMessage);
}

/**
 * Creates a standardized error object for the md2slides system
 * @param {string} module - The module where the error occurred
 * @param {string} operation - The operation that failed
 * @param {string} message - The error message
 * @param {Error} originalError - The original error object (optional)
 * @return {Object} Standardized error object
 */
function createMd2SlidesError(
	module,
	operation,
	message,
	originalError = null,
) {
	return {
		module,
		operation,
		message,
		originalError: originalError ? originalError.message : null,
		timestamp: new Date().toISOString(),
		id: generateUniqueId(),
	};
}
