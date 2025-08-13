/**
 * Shape utility functions for common shape operations
 * Handles styling, positioning, connection sites, and validation
 */

/**
 * Gets center coordinates of a shape
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get center of
 * @returns {Object} - Center coordinates {x, y}
 */
function getCenterOf(shape) {
	return {
		x: shape.getLeft() + shape.getWidth() / 2,
		y: shape.getTop() + shape.getHeight() / 2,
	};
}

/**
 * Gets preferred connection site mapping for different shape types
 * @param {GoogleAppsScript.Slides.ShapeType} shapeType - Type of shape
 * @param {number} connectionCount - Number of connection sites available
 * @returns {Object} - Mapping of sides to connection site indices
 */
function getPreferredConnectionMapping(shapeType, connectionCount) {
	// 8 connection points (common case): original LEFT:7, RIGHT:3 → swap to LEFT:3, RIGHT:7
	if (connectionCount >= 8) {
		return { LEFT: 3, RIGHT: 7, TOP: 1, BOTTOM: 5 };
	}

	// 4 connection points: assume [TOP, RIGHT, BOTTOM, LEFT]
	// Swap left-right → LEFT:1, RIGHT:3 (TOP/BOTTOM unchanged)
	if (connectionCount === 4) {
		return { LEFT: 1, RIGHT: 3, TOP: 0, BOTTOM: 2 };
	}

	// 2 connection points: swap left-right (TOP/BOTTOM first maintain)
	if (connectionCount === 2) {
		return { LEFT: 1, RIGHT: 0, TOP: 0, BOTTOM: 1 };
	}

	// 1 or other non-standard: can only use 0
	if (connectionCount === 1) {
		return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };
	}

	// fallback
	return { LEFT: 0, RIGHT: 0, TOP: 0, BOTTOM: 0 };
}

/**
 * Picks the best connection site for a shape on a given side
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get connection site from
 * @param {string} side - Side to connect to (LEFT, RIGHT, TOP, BOTTOM)
 * @returns {GoogleAppsScript.Slides.ConnectionSite|null} - Connection site or null if none available
 */
function pickConnectionSite(shape, side) {
	const sites = shape.getConnectionSites();
	if (!sites || sites.length === 0) return null;

	const mapping = getPreferredConnectionMapping(
		shape.getShapeType(),
		sites.length,
	);
	const index = mapping[side];

	if (index != null && index < sites.length) {
		return sites[index];
	}
	return sites[0];
}

/**
 * Copies style properties from source shape to target shape
 * @param {GoogleAppsScript.Slides.Shape} sourceShape - Shape to copy style from
 * @param {GoogleAppsScript.Slides.Shape} targetShape - Shape to apply style to
 * @param {boolean} copyTextContent - Whether to copy text content (default: false)
 */
function copyShapeStyle(sourceShape, targetShape, copyTextContent = false) {
	try {
		// Copy fill with alpha support
		try {
			const sourceFill = sourceShape.getFill();
			if (sourceFill && sourceFill.getType() === SlidesApp.FillType.SOLID) {
				const sourceSolidFill = sourceFill.getSolidFill();
				if (sourceSolidFill) {
					// Get the color and alpha separately and apply them
					const color = sourceSolidFill.getColor();
					const alpha = sourceSolidFill.getAlpha();
					targetShape.getFill().setSolidFill(color, alpha);
				}
			}
		} catch (e) {
			console.log("Warning: Could not copy fill: " + e.message);
		}

		// Copy border with complete styling - following the exact pattern from documentation
		try {
			const sourceBorder = sourceShape.getBorder();
			const targetBorder = targetShape.getBorder();

			if (sourceBorder && targetBorder) {
				// First set the border weight (must be done separately)
				try {
					const borderWeight = sourceBorder.getWeight();
					if (borderWeight) {
						targetBorder.setWeight(borderWeight);
					}
				} catch (e) {
					console.log("Warning: Could not copy border weight: " + e.message);
				}

				// Then set the border color using getLineFill() - separate statement as per docs
				try {
					const sourceLineFill = sourceBorder.getLineFill();
					if (
						sourceLineFill &&
						sourceLineFill.getType() === SlidesApp.FillType.SOLID
					) {
						const sourceSolidFill = sourceLineFill.getSolidFill();
						if (sourceSolidFill) {
							const borderColor = sourceSolidFill.getColor();
							targetBorder.getLineFill().setSolidFill(borderColor);
						}
					}
				} catch (e) {
					console.log("Warning: Could not copy border color: " + e.message);
				}

				// Finally set dash style
				try {
					const dashStyle = sourceBorder.getDashStyle();
					if (dashStyle) {
						targetBorder.setDashStyle(dashStyle);
					}
				} catch (e) {
					console.log(
						"Warning: Could not copy border dash style: " + e.message,
					);
				}
			}
		} catch (e) {
			console.log("Warning: Could not copy border styling: " + e.message);
		}

		// Copy text style (but not content unless specified)
		const sourceText = sourceShape.getText();
		const targetText = targetShape.getText();
		if (sourceText && targetText) {
			// Only copy text content if explicitly requested
			if (copyTextContent) {
				targetText.setText(sourceText.asString());
			}

			// Always copy text styling
			try {
				const sourceStyle = sourceText.getTextStyle();
				const targetStyle = targetText.getTextStyle();

				// Copy font family
				try {
					const fontFamily = sourceStyle.getFontFamily();
					if (fontFamily) {
						targetStyle.setFontFamily(fontFamily);
					}
				} catch (e) {
					console.log("Warning: Could not copy font family: " + e.message);
				}

				// Copy font size
				try {
					const fontSize = sourceStyle.getFontSize();
					if (fontSize) {
						targetStyle.setFontSize(fontSize);
					}
				} catch (e) {
					console.log("Warning: Could not copy font size: " + e.message);
				}

				// Copy font color (foreground color)
				try {
					const foregroundColor = sourceStyle.getForegroundColor();
					if (foregroundColor) {
						targetStyle.setForegroundColor(foregroundColor);
					}
				} catch (e) {
					console.log("Warning: Could not copy font color: " + e.message);
				}

				// Copy bold styling - using safer boolean check
				try {
					const isBold = sourceStyle.isBold();
					if (typeof isBold === "boolean") {
						targetStyle.setBold(isBold);
					}
				} catch (e) {
					console.log("Warning: Could not copy bold style: " + e.message);
				}

				// Copy italic styling - using safer boolean check
				try {
					const isItalic = sourceStyle.isItalic();
					if (typeof isItalic === "boolean") {
						targetStyle.setItalic(isItalic);
					}
				} catch (e) {
					console.log("Warning: Could not copy italic style: " + e.message);
				}

				// Copy underline styling - using safer boolean check
				try {
					const isUnderline = sourceStyle.isUnderline();
					if (typeof isUnderline === "boolean") {
						targetStyle.setUnderline(isUnderline);
					}
				} catch (e) {
					console.log("Warning: Could not copy underline style: " + e.message);
				}

				// Copy strikethrough styling if available
				try {
					if (typeof sourceStyle.isStrikethrough === "function") {
						const isStrikethrough = sourceStyle.isStrikethrough();
						if (typeof isStrikethrough === "boolean") {
							targetStyle.setStrikethrough(isStrikethrough);
						}
					}
				} catch (e) {
					console.log(
						"Warning: Could not copy strikethrough style: " + e.message,
					);
				}
			} catch (e) {
				console.log("Warning: Could not copy text style: " + e.message);
			}
		}
	} catch (e) {
		console.log("Warning: Could not copy shape style: " + e.message);
	}
}

/**
 * Debug function to test style copying on selected shapes
 * Select two shapes and run this to copy style from first to second
 */
function debugStyleCopy() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			SlidesApp.getUi().alert("Please select TWO shapes");
			return;
		}

		const elements = range.getPageElements();
		if (elements.length !== 2) {
			SlidesApp.getUi().alert("Please select exactly TWO shapes");
			return;
		}

		const sourceShape = elements[0].asShape();
		const targetShape = elements[1].asShape();

		console.log("=== DEBUG: Style Copy Test ===");
		console.log("Source shape ID: " + sourceShape.getObjectId());
		console.log("Target shape ID: " + targetShape.getObjectId());

		// Test the style copying
		copyShapeStyle(sourceShape, targetShape);

		console.log("Style copying completed - check console for any warnings");
		SlidesApp.getUi().alert(
			"Style copying test completed. Check console logs for details.",
		);
	} catch (e) {
		console.error("Debug style copy error: " + e.message);
		SlidesApp.getUi().alert("Error: " + e.message);
	}
}

/**
 * Validates that a selection contains valid shapes
 * @param {GoogleAppsScript.Slides.PageElementRange} range - Selection range
 * @param {number} expectedCount - Expected number of shapes
 * @returns {Object} - Validation result with shapes array or error message
 */
function validateShapeSelection(range, expectedCount) {
	if (!range) {
		const message =
			expectedCount === 1
				? "Please select a shape."
				: `Please select exactly ${expectedCount === 2 ? "TWO" : expectedCount} shapes.`;
		return { error: message };
	}

	const elements = range.getPageElements();
	if (elements.length !== expectedCount) {
		const message =
			expectedCount === 1
				? "Please select exactly ONE shape."
				: `Please select exactly ${expectedCount === 2 ? "TWO" : expectedCount} shapes.`;
		return { error: message };
	}

	// Validate all elements are shapes
	for (const element of elements) {
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			const message =
				expectedCount === 1
					? "Selected item must be a SHAPE."
					: "All selected items must be SHAPES.";
			return { error: message };
		}
	}

	// Convert to shapes
	const shapes = elements.map((element) => element.asShape());

	// For multiple shapes, verify they're on the same slide
	if (expectedCount > 1) {
		const firstSlideId = String(shapes[0].getParentPage().getObjectId());
		for (let i = 1; i < shapes.length; i++) {
			if (String(shapes[i].getParentPage().getObjectId()) !== firstSlideId) {
				return { error: "All shapes must be on the SAME slide." };
			}
		}
	}

	return expectedCount === 1 ? { shape: shapes[0] } : { shapes };
}

/**
 * Gets the current selection from the active presentation
 * @returns {GoogleAppsScript.Slides.PageElementRange|null} - Selection range or null
 */
function getCurrentSelection() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		return selection.getPageElementRange();
	} catch (e) {
		console.log(`Warning: Could not get selection: ${e.message}`);
		return null;
	}
}

/**
 * Displays an alert to the user
 * @param {string} message - Message to display
 * @param {string} title - Optional title for the alert
 */
function showAlert(message, title = "Alert") {
	try {
		SlidesApp.getUi().alert(title, message);
	} catch (e) {
		console.error(`Failed to show alert: ${e.message}`);
	}
}

/**
 * Gets shape properties in a standardized format
 * @param {GoogleAppsScript.Slides.Shape} shape - Shape to get properties from
 * @returns {Object} - Shape properties {left, top, width, height, center}
 */
function getShapeProperties(shape) {
	const left = shape.getLeft();
	const top = shape.getTop();
	const width = shape.getWidth();
	const height = shape.getHeight();

	return {
		left,
		top,
		width,
		height,
		center: {
			x: left + width / 2,
			y: top + height / 2,
		},
	};
}
