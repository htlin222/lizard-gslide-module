/**
 * Markdown Parser Module
 *
 * Handles parsing markdown text into structured slide data
 */

/**
 * Parse markdown text into a structured slide format
 * @param {string} markdownText - The markdown text to parse
 * @return {Array} Array of slide objects with layout, title, bodyItems, speakerNotes, and listType
 */
function parseMarkdownToStructure(markdownText) {
	try {
		const lines = markdownText.split("\n");
		const slideStructure = [];
		let currentSlide = null;

		for (let i = 0; i < lines.length; i++) {
			const line = lines[i].trim();

			// Skip empty lines
			if (line === "") continue;

			// Check for horizontal rule (---) as slide separator
			if (line === "---") {
				// Skip the separator line - don't add it to any slide content
				continue;
			}
			// Check for speaker notes (> content)
			if (line.startsWith("> ")) {
				if (currentSlide) {
					const speakerNote = line.substring(2).trim();
					if (!currentSlide.speakerNotes) {
						currentSlide.speakerNotes = [];
					}
					currentSlide.speakerNotes.push(speakerNote);
				}
			}
			// Check for H1 heading (# Heading)
			else if (line.startsWith("# ")) {
				// Extract title and remove page numbering pattern if present
				let title = line.substring(2).trim();
				// Remove patterns like "Page 1:" or "Page 10:" from the title
				title = title.replace(/^Page\s+\d+:\s*/i, "");

				// Create a new SECTION_HEADER slide
				currentSlide = {
					layout: "SECTION_HEADER",
					title: title,
					bodyItems: [],
					speakerNotes: [],
				};
				slideStructure.push(currentSlide);
			}
			// Check for H2 heading (## Heading)
			else if (line.startsWith("## ")) {
				// Extract title and remove page numbering pattern if present
				let title = line.substring(3).trim();
				// Remove patterns like "Page 1:" or "Page 10:" from the title
				title = title.replace(/^Page\s+\d+:\s*/i, "");

				// Create a new TITLE_AND_BODY slide
				currentSlide = {
					layout: "TITLE_AND_BODY",
					title: title,
					bodyItems: [],
					speakerNotes: [],
					listType: "bullet", // Default to bullet list
				};
				slideStructure.push(currentSlide);
			}
			// Add content to current slide if it's a TITLE_AND_BODY
			else if (currentSlide && currentSlide.layout === "TITLE_AND_BODY") {
				// Process list items and regular text
				let content = line;

				// Detect and set list type based on the first list item encountered
				if (currentSlide.bodyItems.length === 0) {
					if (/^\(\d+\)\s/.test(line)) {
						currentSlide.listType = "numbered_parens";
					} else if (/^\d+\.\s/.test(line)) {
						currentSlide.listType = "numbered";
					} else if (/^[A-Z]\.\s/.test(line)) {
						currentSlide.listType = "lettered";
					} else if (line.startsWith("- ") || line.startsWith("* ")) {
						currentSlide.listType = "bullet";
					}
				}

				// Remove list markers if present
				if (line.startsWith("- ")) {
					content = line.substring(2).trim();
				} else if (line.startsWith("* ")) {
					content = line.substring(2).trim();
				} else if (/^\(\d+\)\s/.test(line)) {
					content = line.substring(line.indexOf(")") + 1).trim();
				} else if (/^\d+\.\s/.test(line)) {
					content = line.substring(line.indexOf(".") + 1).trim();
				} else if (/^[A-Z]\.\s/.test(line)) {
					content = line.substring(2).trim();
				}

				currentSlide.bodyItems.push(content);
			}
		}

		return slideStructure;
	} catch (error) {
		console.error(`Error parsing markdown: ${error.message}`);
		return [];
	}
}

/**
 * Detects the list type from a markdown line
 * @param {string} line - The line to analyze
 * @return {string} - "numbered", "lettered", "numbered_parens", "bullet", or "none"
 */
function detectListType(line) {
	if (/^\(\d+\)\s/.test(line)) {
		return "numbered_parens";
	}
	if (/^\d+\.\s/.test(line)) {
		return "numbered";
	}
	if (/^[A-Z]\.\s/.test(line)) {
		return "lettered";
	}
	if (line.startsWith("- ") || line.startsWith("* ")) {
		return "bullet";
	}
	return "none";
}

/**
 * Removes list markers from a line
 * @param {string} line - The line to process
 * @return {string} - The line with list markers removed
 */
function removeListMarkers(line) {
	if (line.startsWith("- ")) {
		return line.substring(2).trim();
	}
	if (line.startsWith("* ")) {
		return line.substring(2).trim();
	}
	if (/^\(\d+\)\s/.test(line)) {
		return line.substring(line.indexOf(")") + 1).trim();
	}
	if (/^\d+\.\s/.test(line)) {
		return line.substring(line.indexOf(".") + 1).trim();
	}
	if (/^[A-Z]\.\s/.test(line)) {
		return line.substring(2).trim();
	}
	return line;
}
