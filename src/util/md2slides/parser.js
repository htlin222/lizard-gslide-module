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
		let inCodeBlock = false;
		let codeBlockContent = [];
		let codeBlockLanguage = "";
		let currentH2Title = ""; // Track the current H2 title for H3 slides

		for (let i = 0; i < lines.length; i++) {
			const line = lines[i];
			const trimmedLine = line.trim();

			// Check for code block markers
			if (trimmedLine.startsWith("```")) {
				if (!inCodeBlock) {
					// Starting a code block
					inCodeBlock = true;
					// Extract language if specified (e.g., ```python or ```r)
					codeBlockLanguage = trimmedLine.substring(3).trim();
					codeBlockContent = [];
				} else {
					// Ending a code block
					inCodeBlock = false;

					// Add the code block to the current slide if it exists
					if (currentSlide && codeBlockContent.length > 0) {
						if (!currentSlide.codeBlocks) {
							currentSlide.codeBlocks = [];
						}
						currentSlide.codeBlocks.push({
							language: codeBlockLanguage,
							content: codeBlockContent.join("\n"),
						});
					}

					// Reset code block variables
					codeBlockContent = [];
					codeBlockLanguage = "";
				}
				continue; // Skip the ``` line itself
			}

			// If we're inside a code block, collect the content
			if (inCodeBlock) {
				// Preserve original line without trimming for code formatting
				codeBlockContent.push(line);
				continue;
			}

			// Skip empty lines outside of code blocks
			if (trimmedLine === "") continue;

			// Check for horizontal rule (---) as slide separator
			if (trimmedLine === "---") {
				// Skip the separator line - don't add it to any slide content
				continue;
			}
			// Check for speaker notes (> content)
			if (trimmedLine.startsWith("> ")) {
				if (currentSlide) {
					const speakerNote = trimmedLine.substring(2).trim();
					if (!currentSlide.speakerNotes) {
						currentSlide.speakerNotes = [];
					}
					currentSlide.speakerNotes.push(speakerNote);
				}
			}
			// Check for footer items (@ content)
			else if (trimmedLine.startsWith("@ ")) {
				if (currentSlide) {
					const footerItem = trimmedLine.substring(2).trim();
					if (!currentSlide.footerItems) {
						currentSlide.footerItems = [];
					}
					currentSlide.footerItems.push(footerItem);
				}
			}
			// Check for H1 heading (# Heading)
			else if (trimmedLine.startsWith("# ")) {
				// Extract title and remove page numbering pattern if present
				let title = trimmedLine.substring(2).trim();
				// Remove patterns like "Page 1:" or "Page 10:" from the title
				title = title.replace(/^Page\s+\d+:\s*/i, "");

				// Create a new SECTION_HEADER slide
				currentSlide = {
					layout: "SECTION_HEADER",
					title: title,
					bodyItems: [],
					speakerNotes: [],
					codeBlocks: [],
					footerItems: [], // For @ prefixed lines
				};
				slideStructure.push(currentSlide);
			}
			// Check for H2 heading (## Heading)
			else if (trimmedLine.startsWith("## ")) {
				// Extract title and remove page numbering pattern if present
				let title = trimmedLine.substring(3).trim();
				// Remove patterns like "Page 1:" or "Page 10:" from the title
				title = title.replace(/^Page\s+\d+:\s*/i, "");

				// Track this H2 title for future H3 slides
				currentH2Title = title;

				// Create a new TITLE_AND_BODY slide
				currentSlide = {
					layout: "TITLE_AND_BODY",
					title: title,
					bodyItems: [],
					speakerNotes: [],
					listType: "bullet", // Default to bullet list
					codeBlocks: [],
					footerItems: [], // For @ prefixed lines
				};
				slideStructure.push(currentSlide);
			}
			// Check for H3 heading (### Heading)
			else if (trimmedLine.startsWith("### ")) {
				// Extract title and remove page numbering pattern if present
				let title = trimmedLine.substring(4).trim();
				// Remove patterns like "Page 1:" or "Page 10:" from the title
				title = title.replace(/^Page\s+\d+:\s*/i, "");

				// Create a new TITLE_AND_BODY slide with parent H2 title
				currentSlide = {
					layout: "TITLE_AND_BODY",
					title: title,
					bodyItems: [],
					speakerNotes: [],
					listType: "bullet", // Default to bullet list
					codeBlocks: [],
					footerItems: [], // For @ prefixed lines
					parentTitle: currentH2Title, // Store the parent H2 title
				};
				slideStructure.push(currentSlide);
			}
			// Add content to current slide if it's a TITLE_AND_BODY
			else if (currentSlide && currentSlide.layout === "TITLE_AND_BODY") {
				// Process list items and regular text
				let content = trimmedLine;

				// Detect and set list type based on the first list item encountered
				if (currentSlide.bodyItems.length === 0) {
					if (/^\(\d+\)\s/.test(trimmedLine)) {
						currentSlide.listType = "numbered_parens";
					} else if (/^\d+\.\s/.test(trimmedLine)) {
						currentSlide.listType = "numbered";
					} else if (/^[A-Z]\.\s/.test(trimmedLine)) {
						currentSlide.listType = "lettered";
					} else if (
						trimmedLine.startsWith("- ") ||
						trimmedLine.startsWith("* ")
					) {
						currentSlide.listType = "bullet";
					}
				}

				// Remove list markers if present
				if (trimmedLine.startsWith("- ")) {
					content = trimmedLine.substring(2).trim();
				} else if (trimmedLine.startsWith("* ")) {
					content = trimmedLine.substring(2).trim();
				} else if (/^\(\d+\)\s/.test(trimmedLine)) {
					content = trimmedLine.substring(trimmedLine.indexOf(")") + 1).trim();
				} else if (/^\d+\.\s/.test(trimmedLine)) {
					content = trimmedLine.substring(trimmedLine.indexOf(".") + 1).trim();
				} else if (/^[A-Z]\.\s/.test(trimmedLine)) {
					content = trimmedLine.substring(2).trim();
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
