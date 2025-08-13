/**
 * Flowchart module index - Global function availability for Google Apps Script
 *
 * This file ensures all modular flowchart functions are available globally
 * since Google Apps Script doesn't support ES6 modules or imports.
 *
 * Include this file in appsscript.json to make all flowchart utilities available.
 */

// Note: In Google Apps Script, all .js files are automatically included
// and all functions become globally available. This file serves as documentation
// of the modular system structure.

/**
 * Module Structure:
 *
 * src/util/flowchart/
 * ├── connectionUtils.js      # Shape connection logic and validation
 * ├── childCreationUtils.js   # Child shape creation and positioning
 * ├── graphIdUtils.js         # Hierarchical naming and ID management
 * ├── shapeUtils.js          # Common shape operations and styling
 * ├── layoutDetection.js     # Layout pattern detection (LR/TD)
 * ├── siblingCreationUtils.js # Sibling shape creation with layout detection
 * ├── backgroundUtils.js     # Background rectangle creation and styling
 * ├── stageBarUtils.js       # Stage bar (HOME_PLATE) creation for process flows
 * ├── debugUtils.js          # Debug and inspection utilities
 * ├── lineUpdateUtils.js     # Line update and style management
 * ├── smartSelectionUtils.js # Smart selection based on graph ID relationships
 * ├── main.js                # Main interface functions and sidebar
 * └── index.js              # This documentation file
 *
 * Key Functions Available:
 *
 * From connectionUtils.js:
 * - validateConnectionElements(elements)
 * - determineConnectionSides(shapeA, shapeB, orientation)
 * - createConnection(shapeA, shapeB, orientation, lineType, startArrow, endArrow)
 * - connectSelectedShapes(orientation, lineType, startArrow, endArrow)
 *
 * From childCreationUtils.js:
 * - validateParentElement(range)
 * - calculateChildPositions(parent, direction, gap, count)
 * - createSingleChild(parentShape, position, direction, lineType, startArrow, endArrow)
 * - createChildrenInDirection(direction, gap, lineType, count, startArrow, endArrow)
 *
 * From graphIdUtils.js:
 * - parseGraphId(graphId)
 * - generateGraphId(parent, layout, current, children)
 * - getNextLevel(parentLevel)
 * - generateSiblingIds(baseLevel, count)
 * - setShapeGraphId(shape, graphId)
 * - getShapeGraphId(shape)
 * - updateParentWithChildren(parentShape, newChildIds)
 * - initializeAsRootGraphShape(shape)
 * - determineParentChildRelationship(idA, idB)
 * - updateGraphShapeRelationship(shapeA, shapeB)
 *
 * From shapeUtils.js:
 * - getCenterOf(shape)
 * - getPreferredConnectionMapping(shapeType, connectionCount)
 * - pickConnectionSite(shape, side)
 * - copyShapeStyle(sourceShape, targetShape)
 * - validateShapeSelection(range, expectedCount)
 * - getCurrentSelection()
 * - showAlert(message, title)
 * - getShapeProperties(shape)
 *
 * From layoutDetection.js:
 * - detectLayoutFromConnections(parentShape, siblingShapes)
 * - detectLayoutFromPositions(siblingShapes, tolerance)
 * - detectLayout(parentShape, siblingShapes)
 *
 * From siblingCreationUtils.js:
 * - createSiblingShape(horizontalGap, verticalGap, lineType, startArrow, endArrow)
 *
 * From backgroundUtils.js:
 * - addBackgroundToSelectedElements(padding, bgColor, opacity)
 * - calculateShapesBoundingBox(shapes)
 * - createBackgroundRectangle(slide, left, top, width, height, bgColor, opacity)
 * - createCustomBackground(shapes, style)
 *
 * From stageBarUtils.js:
 * - addStageBar(baseY, offsetX, extraWidth, height, fillColor, opacity, strokeColor)
 * - addDefaultStageBar()
 * - addThemedStageBar(theme)
 * - addMultipleStageBar(yPositions, fillColor)
 *
 * From debugUtils.js:
 * - showSelectedShapeGraphId()
 * - formatGraphIdInfo(graphId)
 * - clearSelectedShapeGraphId()
 * - identifyConnectedShapes()
 * - initializeRootGraphShape()
 * - analyzeCurrentSlide()
 * - debugShowTitlePlaceholders() [alias for showSelectedShapeGraphId]
 *
 * From lineUpdateUtils.js:
 * - updateSelectedLines(lineType, startArrow, endArrow)
 * - getSelectedElements(selection)
 * - processLineUpdates(elements, lineType, startArrow, endArrow)
 * - updateSingleLine(line, lineType, startArrow, endArrow, index)
 * - extractLineStyle(line, index)
 * - recreateLine(startShape, endShape, lineType, startArrow, endArrow, lineStyle, index)
 * - calculateLineOrientation(startShape, endShape)
 * - applyLineStyle(line, lineStyle)
 * - formatUpdateResults(results)
 *
 * From smartSelectionUtils.js:
 * - validateSelectedShape()
 * - selectAllSiblings()
 * - selectAllLevel()
 * - selectAllParents()
 * - selectFamily()
 *
 * From main.js:
 * - showFlowchartSidebar()
 * - connectSelectedShapesSmart(lineType)
 * - connectSelectedShapesVertical(lineType, startArrow, endArrow)
 * - connectSelectedShapesHorizontal(lineType, startArrow, endArrow)
 * - createChildInDirection(direction, gap, lineType, count, startArrow, endArrow)
 * - createChildTop(gap, lineType, count, startArrow, endArrow)
 * - createChildRight(gap, lineType, count, startArrow, endArrow)
 * - createChildBottom(gap, lineType, count, startArrow, endArrow)
 * - createChildLeft(gap, lineType, count, startArrow, endArrow)
 * - createChildTopWithText(gap, lineType, count, startArrow, endArrow, texts)
 * - createChildRightWithText(gap, lineType, count, startArrow, endArrow, texts)
 * - createChildBottomWithText(gap, lineType, count, startArrow, endArrow, texts)
 * - createChildLeftWithText(gap, lineType, count, startArrow, endArrow, texts)
 * - connectExistingGraphShapes(lineType, startArrow, endArrow)
 */

// This file serves as documentation only.
// In Google Apps Script, all functions in .js files are automatically available globally.
