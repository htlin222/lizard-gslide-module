/**
 * Main flowchart interface functions
 * Provides the main entry points for flowchart operations
 */

/**
 * Shows the flowchart sidebar for interactive flowchart operations
 */
function showFlowchartSidebar() {
	try {
		const html = HtmlService.createHtmlOutputFromFile(
			"src/components/flowchartSidebar.html",
		)
			.setWidth(300)
			.setTitle("Flowchart Tools");

		SlidesApp.getUi().showSidebar(html);
	} catch (e) {
		console.error(`Error showing flowchart sidebar: ${e.message}`);
		SlidesApp.getUi().alert(
			"Error",
			`Could not open the flowchart sidebar: ${e.message}`,
		);
	}
}

/**
 * Connects two selected shapes with a smart line
 * This is the main function called from the sidebar
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 */
function connectSelectedShapesSmart(lineType = "STRAIGHT") {
	// Default to horizontal connection for backwards compatibility
	return connectSelectedShapesHorizontal(lineType);
}

/**
 * Connects two selected shapes vertically (top/down)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function connectSelectedShapesVertical(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return connectSelectedShapes("vertical", lineType, startArrow, endArrow);
}

/**
 * Connects two selected shapes horizontally (left/right)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function connectSelectedShapesHorizontal(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return connectSelectedShapes("horizontal", lineType, startArrow, endArrow);
}

/**
 * Base function to create child shapes in any direction
 * @param {string} direction - Direction to create child ("TOP", "RIGHT", "BOTTOM", "LEFT")
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildInDirection(
	direction,
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildrenInDirection(
		direction,
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes above the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildTop(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"TOP",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes to the right of the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildRight(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"RIGHT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes below the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildBottom(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"BOTTOM",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Creates child shapes to the left of the selected shape
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 */
function createChildLeft(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	return createChildInDirection(
		"LEFT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
	);
}

/**
 * Updates connection between two existing graph shapes
 * Also handles establishing the parent-child relationship in their IDs
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style
 * @param {string} endArrow - End arrow style
 */
function connectExistingGraphShapes(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert("Please select exactly TWO graph shapes.");
	}

	const els = range.getPageElements();
	if (els.length !== 2) {
		return SlidesApp.getUi().alert("Please select exactly TWO graph shapes.");
	}

	if (
		els[0].getPageElementType() !== SlidesApp.PageElementType.SHAPE ||
		els[1].getPageElementType() !== SlidesApp.PageElementType.SHAPE
	) {
		return SlidesApp.getUi().alert("Both selected items must be SHAPES.");
	}

	const shapeA = els[0].asShape();
	const shapeB = els[1].asShape();

	// Get their graph IDs
	const idA = getShapeGraphId(shapeA);
	const idB = getShapeGraphId(shapeB);

	if (!idA || !idB) {
		return SlidesApp.getUi().alert(
			"Both shapes must have graph IDs. Use 'Initialize Root' first.",
		);
	}

	// Parse the IDs to understand the hierarchy
	const parsedA = parseGraphId(idA);
	const parsedB = parseGraphId(idB);

	if (!parsedA || !parsedB) {
		return SlidesApp.getUi().alert(
			"Invalid graph ID format on one or both shapes.",
		);
	}

	// Determine parent-child relationship based on hierarchy level
	// Lower levels (A < B < C) are parents of higher levels
	const levelA = parsedA.current.match(/^([A-Z]+)/)?.[1] || "A";
	const levelB = parsedB.current.match(/^([A-Z]+)/)?.[1] || "A";

	let parentShape;
	let childShape;
	let parentId;
	let childId;

	if (levelA <= levelB) {
		// A is parent, B is child
		parentShape = shapeA;
		childShape = shapeB;
		parentId = parsedA;
		childId = parsedB;
	} else {
		// B is parent, A is child
		parentShape = shapeB;
		childShape = shapeA;
		parentId = parsedB;
		childId = parsedA;
	}

	// Update parent to include this child
	if (!parentId.children.includes(childId.current)) {
		const updatedChildren = [...parentId.children, childId.current];
		const newParentId = generateGraphId(
			parentId.parent,
			parentId.layout,
			parentId.current,
			updatedChildren,
		);
		setShapeGraphId(parentShape, newParentId);
	}

	// Update child to reflect correct parent using parent's layout
	const newChildId = generateGraphId(
		parentId.current,
		parentId.layout,
		childId.current,
		childId.children,
	);
	setShapeGraphId(childShape, newChildId);

	// Create the visual connection using the existing connection logic
	if (parentShape === shapeA) {
		return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
	}
	return connectSelectedShapesHorizontal(lineType, startArrow, endArrow);
}

/**
 * Initializes a selected shape as a root graph node
 * Useful for starting a new flowchart hierarchy
 */
function initializeRootGraphShape() {
	const pres = SlidesApp.getActivePresentation();
	const selection = pres.getSelection();
	const range = selection.getPageElementRange();

	if (!range) {
		return SlidesApp.getUi().alert(
			"Please select a shape to initialize as root graph node.",
		);
	}

	const els = range.getPageElements();
	if (els.length !== 1) {
		return SlidesApp.getUi().alert("Please select exactly ONE shape.");
	}

	const element = els[0];
	if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
		return SlidesApp.getUi().alert("Selected item must be a SHAPE.");
	}

	const shape = element.asShape();
	// Initialize as root without layout annotation
	const rootId = generateGraphId("", "", "A1", []);
	setShapeGraphId(shape, rootId);

	SlidesApp.getUi().alert(
		"Root graph shape initialized",
		`Shape is now: ${rootId}`,
	);
}

/**
 * Shows the Graph ID of the currently selected shape
 * @returns {string} Graph ID information about the selected shape
 */
function showSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to show its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get Graph ID from title (alt text)
		const graphId = getShapeGraphId(shape);

		if (graphId) {
			// Parse the graph ID to show more details
			const parsed = parseGraphId(graphId);
			if (parsed) {
				const details = [];
				details.push(`📊 Graph ID:\n${graphId}`);
				details.push(`├─ Parent: ${parsed.parent || "(root)"}`);
				details.push(`├─ Layout: ${parsed.layout || "(none)"}`);
				details.push(`├─ Current: ${parsed.current}`);
				details.push(
					`└─ Children: ${parsed.children.length > 0 ? parsed.children.join(", ") : "(none)"}`,
				);
				return details.join("\n");
			}
			return `📊 Graph ID:\n${graphId}`;
		} else {
			return "No Graph ID found. This shape may not be part of a flowchart.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

/**
 * Clears the Graph ID from the currently selected shape
 * @returns {string} Confirmation message
 */
function clearSelectedShapeGraphId() {
	try {
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		const range = selection.getPageElementRange();

		if (!range) {
			return "Please select a shape to clear its Graph ID.";
		}

		const els = range.getPageElements();
		if (els.length !== 1) {
			return "Please select exactly ONE shape.";
		}

		const element = els[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
			return "Selected item must be a SHAPE.";
		}

		const shape = element.asShape();

		// Get current Graph ID for confirmation
		const currentGraphId = getShapeGraphId(shape);

		if (currentGraphId) {
			// Clear the title (alt text)
			shape.setTitle("");
			return `✅ Graph ID cleared successfully!\nPrevious ID was: ${currentGraphId}`;
		} else {
			return "No Graph ID to clear. This shape doesn't have a Graph ID.";
		}
	} catch (e) {
		const errorMsg = `Error: ${e.message}`;
		console.error(errorMsg);
		return errorMsg;
	}
}

// Keep the old function name for backward compatibility but redirect to new function
function debugShowTitlePlaceholders() {
	return showSelectedShapeGraphId();
}

/**
 * Identifies which shapes are connected to the selected line
 * This helps debug connection issues by showing what shapes would be connected
 * @returns {string} - Information about connected shapes
 */
function identifyConnectedShapes() {
	try {
		console.log("identifyConnectedShapes function called");
		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();
		console.log("Got selection:", selection ? "exists" : "null");

		if (!selection) {
			return "Please select a line to identify its connected shapes.";
		}

		const selectionType = selection.getSelectionType();

		let elements = [];

		if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT_RANGE) {
			const range = selection.getPageElementRange();
			if (!range) {
				return "No elements selected.";
			}
			elements = range.getPageElements();
		} else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
			const range = selection.getPageElementRange();
			if (range) {
				elements = range.getPageElements();
			} else {
				return "Unable to access selected element.";
			}
		} else {
			return "Please select a line element.";
		}

		if (elements.length !== 1) {
			return "Please select exactly ONE line.";
		}

		const element = elements[0];
		if (element.getPageElementType() !== SlidesApp.PageElementType.LINE) {
			return "Selected element must be a LINE.";
		}

		const line = element.asLine();
		const slide = line.getParentPage();

		// Get line properties
		const lineLeft = line.getLeft();
		const lineTop = line.getTop();
		const lineWidth = line.getWidth();
		const lineHeight = line.getHeight();

		// Get all shapes on slide
		const allShapes = slide.getShapes();

		const results = [];
		results.push(`📊 Line Analysis:`);
		results.push(`Position: (${Math.round(lineLeft)}, ${Math.round(lineTop)})`);
		results.push(`Size: ${Math.round(lineWidth)} × ${Math.round(lineHeight)}`);
		results.push(`Total shapes on slide: ${allShapes.length}`);
		results.push(``);

		// Show first few shapes as candidates
		results.push(`🎯 Shape Candidates:`);
		for (let i = 0; i < Math.min(4, allShapes.length); i++) {
			const shape = allShapes[i];
			const shapeLeft = Math.round(shape.getLeft());
			const shapeTop = Math.round(shape.getTop());
			const shapeWidth = Math.round(shape.getWidth());
			const shapeHeight = Math.round(shape.getHeight());

			// Calculate distance from line
			const lineCenter = {
				x: lineLeft + lineWidth / 2,
				y: lineTop + lineHeight / 2,
			};
			const shapeCenter = {
				x: shapeLeft + shapeWidth / 2,
				y: shapeTop + shapeHeight / 2,
			};
			const distance = Math.round(
				Math.sqrt(
					(lineCenter.x - shapeCenter.x) ** 2 +
						(lineCenter.y - shapeCenter.y) ** 2,
				),
			);

			results.push(
				`Shape ${i + 1}: (${shapeLeft}, ${shapeTop}) ${shapeWidth}×${shapeHeight} - Distance: ${distance}px`,
			);
		}

		results.push(``);
		results.push(`💡 Current logic will connect:`);
		results.push(
			`• Shape 1: (${Math.round(allShapes[0].getLeft())}, ${Math.round(allShapes[0].getTop())})`,
		);
		results.push(
			`• Shape 2: (${Math.round(allShapes[1].getLeft())}, ${Math.round(allShapes[1].getTop())})`,
		);

		return results.join("\n");
	} catch (e) {
		return `Error: ${e.message}`;
	}
}

/**
 * Updates the selected lines with new line type and arrow styles
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @returns {string} - Result message
 */
function updateSelectedLines(
	lineType = "STRAIGHT",
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
) {
	try {
		console.log(
			`Starting line update: type=${lineType}, startArrow=${startArrow}, endArrow=${endArrow}`,
		);

		const pres = SlidesApp.getActivePresentation();
		const selection = pres.getSelection();

		if (!selection) {
			console.log("No selection found");
			return "Please select one or more lines to update.";
		}

		const selectionType = selection.getSelectionType();
		console.log(`Selection type: ${selectionType}`);

		let elements = [];

		if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT_RANGE) {
			// Multiple elements selected
			const range = selection.getPageElementRange();
			if (!range) {
				console.log("No page element range found");
				return "No elements selected.";
			}
			elements = range.getPageElements();
		} else if (selectionType === SlidesApp.SelectionType.PAGE_ELEMENT) {
			// Single element selected
			const range = selection.getPageElementRange();
			if (range) {
				elements = range.getPageElements();
			} else {
				console.log("Could not get page element range for single selection");
				return "Unable to access selected element. Please try selecting the line again.";
			}
		} else {
			console.log(`Unsupported selection type: ${selectionType}`);
			return "Please select one or more line elements to update.";
		}

		console.log(`Found ${elements.length} selected elements`);

		let updatedCount = 0;
		let skippedCount = 0;
		let errorCount = 0;

		for (let i = 0; i < elements.length; i++) {
			const element = elements[i];
			const elementType = element.getPageElementType();
			console.log(`Element ${i}: type = ${elementType}`);

			if (elementType === SlidesApp.PageElementType.LINE) {
				try {
					const line = element.asLine();
					console.log(`Processing line ${i}`);

					// Get connection info before deletion
					const slide = line.getParentPage();
					const startConnection = line.getStart();
					const endConnection = line.getEnd();

					console.log(
						`Line connections - Start: ${startConnection ? "exists" : "null"}, End: ${endConnection ? "exists" : "null"}`,
					);

					if (!startConnection || !endConnection) {
						console.log(`Line ${i} missing connection sites`);
						skippedCount++;
						continue;
					}

					// Simple approach: just get the first two shapes on the slide
					// Since the user selected a line that connects shapes, we assume it connects the main shapes
					const allShapes = slide.getShapes();
					console.log(`Found ${allShapes.length} shapes on slide`);

					if (allShapes.length < 2) {
						console.log(`Line ${i} - not enough shapes on slide to connect`);
						skippedCount++;
						continue;
					}

					// Take the first two shapes - this is a simple heuristic
					// In practice, the user is updating a line they can see connecting two specific shapes
					const startShape = allShapes[0];
					const endShape = allShapes[1];

					console.log(
						`Using shapes: shape1 and shape2 as connection candidates`,
					);
					console.log(
						`Shape1 at (${startShape.getLeft()}, ${startShape.getTop()}), Shape2 at (${endShape.getLeft()}, ${endShape.getTop()})`,
					);

					// Store line style properties before deletion (with error handling)
					let lineWeight = 2; // default weight
					let lineColor = null;

					try {
						const lineStyle = line.getLineStyle();
						if (lineStyle && lineStyle.getWeight) {
							lineWeight = lineStyle.getWeight();
						}

						if (lineStyle && lineStyle.getSolidFill) {
							const solidFill = lineStyle.getSolidFill();
							if (solidFill) {
								const color = solidFill.getColor();
								if (color && color.getColorType() === SlidesApp.ColorType.RGB) {
									const rgb = color.asRgbColor();
									lineColor = {
										red: rgb.getRed(),
										green: rgb.getGreen(),
										blue: rgb.getBlue(),
									};
								}
							}
						}
					} catch (styleError) {
						console.log(
							`Could not read line style: ${styleError.message}, using defaults`,
						);
					}

					console.log(
						`Line ${i} style - Weight: ${lineWeight}, Color: ${lineColor ? "exists" : "default"}`,
					);

					// Remove the old line first
					line.remove();
					console.log(`Line ${i} removed`);

					// Create new line using the createConnection function which handles this properly
					console.log(`Creating new line with type: ${lineType}`);

					// Determine orientation based on shape positions (like connectionUtils does)
					const startCenter = {
						x: startShape.getLeft() + startShape.getWidth() / 2,
						y: startShape.getTop() + startShape.getHeight() / 2,
					};
					const endCenter = {
						x: endShape.getLeft() + endShape.getWidth() / 2,
						y: endShape.getTop() + endShape.getHeight() / 2,
					};

					// Determine if it's more horizontal or vertical
					const dx = Math.abs(endCenter.x - startCenter.x);
					const dy = Math.abs(endCenter.y - startCenter.y);
					const orientation = dx > dy ? "horizontal" : "vertical";

					console.log(`Using orientation: ${orientation} (dx=${dx}, dy=${dy})`);

					// Use the createConnection function from connectionUtils
					console.log(
						`Calling createConnection with startShape, endShape, orientation=${orientation}, lineType=${lineType}`,
					);

					const newLine = createConnection(
						startShape,
						endShape,
						orientation,
						lineType,
						startArrow,
						endArrow,
					);

					console.log(
						`createConnection returned: ${newLine ? "LINE OBJECT" : "NULL"}`,
					);

					if (!newLine) {
						console.log(
							`❌ Failed to create connection for line ${i} - createConnection returned null`,
						);
						errorCount++;
						continue;
					}

					console.log(
						`✅ New line created for element ${i} with arrows already applied`,
					);

					// Apply line style properties (with error handling)
					try {
						const newLineStyle = newLine.getLineStyle();
						if (newLineStyle && newLineStyle.setWeight) {
							newLineStyle.setWeight(lineWeight);
							console.log(`Set line weight: ${lineWeight}`);
						}

						if (lineColor && newLineStyle && newLineStyle.setSolidFill) {
							newLineStyle.setSolidFill(
								lineColor.red,
								lineColor.green,
								lineColor.blue,
							);
							console.log(
								`Set line color: RGB(${lineColor.red}, ${lineColor.green}, ${lineColor.blue})`,
							);
						}
					} catch (styleError) {
						console.log(
							`Could not apply line style: ${styleError.message}, but line was created`,
						);
					}

					updatedCount++;
					console.log(`Successfully updated line ${i}`);
				} catch (lineError) {
					console.error(`Error updating line ${i}: ${lineError.message}`);
					console.error(`Stack trace: ${lineError.stack}`);
					errorCount++;
				}
			} else {
				console.log(`Element ${i} is not a line, skipping`);
				skippedCount++;
			}
		}

		console.log(
			`Update complete - Updated: ${updatedCount}, Skipped: ${skippedCount}, Errors: ${errorCount}`,
		);

		if (updatedCount === 0) {
			return errorCount > 0
				? `❌ Failed to update lines (${errorCount} errors). Check console for details.`
				: "No lines were selected. Please select one or more lines.";
		}

		let message = `✅ Updated ${updatedCount} line${updatedCount > 1 ? "s" : ""}`;
		if (skippedCount > 0) {
			message += ` (skipped ${skippedCount} non-line element${skippedCount > 1 ? "s" : ""})`;
		}
		if (errorCount > 0) {
			message += ` (${errorCount} error${errorCount > 1 ? "s" : ""})`;
		}
		return message;
	} catch (e) {
		console.error(`Error updating lines: ${e.message}`);
		console.error(`Stack trace: ${e.stack}`);
		return `Error: ${e.message}`;
	}
}

/**
 * Creates child shapes above the selected shape with custom text
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
 */
function createChildTopWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"TOP",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes to the right of the selected shape with custom text
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
 */
function createChildRightWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"RIGHT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes below the selected shape with custom text
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
 */
function createChildBottomWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"BOTTOM",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}

/**
 * Creates child shapes to the left of the selected shape with custom text
 * @param {number} gap - Gap between shapes in points (default 20)
 * @param {string} lineType - Type of line to use (STRAIGHT, BENT, or CURVED)
 * @param {number} count - Number of children to create (default 1)
 * @param {string} startArrow - Start arrow style (NONE, FILL_ARROW, etc.)
 * @param {string} endArrow - End arrow style (NONE, FILL_ARROW, etc.)
 * @param {Array} texts - Array of text strings for each child shape
 */
function createChildLeftWithText(
	gap = 20,
	lineType = "STRAIGHT",
	count = 1,
	startArrow = "NONE",
	endArrow = "FILL_ARROW",
	texts = [],
) {
	return createChildrenInDirectionWithText(
		"LEFT",
		gap,
		lineType,
		count,
		startArrow,
		endArrow,
		texts,
	);
}
