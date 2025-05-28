/**
 * Utility to center elements by averaging the padding around them
 * This helps to position elements evenly between their neighbors
 */

/**
 * Centers the selected element(s) by averaging the padding from its nearest neighbors
 * in all four directions (top, bottom, left, right)
 * If no neighbors are found in a particular direction, uses the presentation edge
 * 
 * @returns {boolean} True if the operation was successful, false otherwise
 */
function averagePadding() {
  // Get the active presentation and selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const selectionType = selection.getSelectionType();
  
  // Step 1: Validate that elements are selected
  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    SlidesApp.getUi().alert('Please select one or more elements');
    return false;
  }
  
  const pageElements = selection.getPageElementRange().getPageElements();
  if (pageElements.length === 0) {
    SlidesApp.getUi().alert('Please select at least one element');
    return false;
  }
  
  const currentPage = selection.getCurrentPage();
  const allPageElements = currentPage.getPageElements();
  
  // Get presentation dimensions
  const slideWidth = presentation.getPageWidth();
  const slideHeight = presentation.getPageHeight();
  
  // Step 2: Calculate the bounding box of all selected elements
  let minLeft = Infinity;
  let minTop = Infinity;
  let maxRight = 0;
  let maxBottom = 0;
  
  // Find the extreme coordinates of the selected elements
  pageElements.forEach(element => {
    const left = element.getLeft();
    const top = element.getTop();
    const right = left + element.getWidth();
    const bottom = top + element.getHeight();
    
    minLeft = Math.min(minLeft, left);
    minTop = Math.min(minTop, top);
    maxRight = Math.max(maxRight, right);
    maxBottom = Math.max(maxBottom, bottom);
  });
  
  // Calculate the dimensions of the bounding box
  const selectedLeft = minLeft;
  const selectedTop = minTop;
  const selectedWidth = maxRight - minLeft;
  const selectedHeight = maxBottom - minTop;
  const selectedRight = maxRight;
  const selectedBottom = maxBottom;
  
  // Step 3: Find nearest neighbors in all four directions
  // Create a set of selected element IDs for quick lookup
  const selectedElementIds = new Set();
  pageElements.forEach(element => {
    selectedElementIds.add(element.getObjectId());
  });
  
  const neighbors = {
    top: findNearestTopNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allPageElements, selectedElementIds),
    bottom: findNearestBottomNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allPageElements, selectedElementIds),
    left: findNearestLeftNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allPageElements, selectedElementIds),
    right: findNearestRightNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allPageElements, selectedElementIds)
  };
  
  // Step 4: Calculate new position based on average padding
  let newX = selectedLeft;
  let newY = selectedTop;
  let horizontalCentered = false;
  let verticalCentered = false;
  
  // Calculate horizontal centering
  // If both left and right neighbors exist, center between them
  if (neighbors.left && neighbors.right) {
    const leftEdge = neighbors.left.right;
    const rightEdge = neighbors.right.left;
    const availableWidth = rightEdge - leftEdge;
    newX = leftEdge + (availableWidth - selectedWidth) / 2;
    horizontalCentered = true;
  } 
  // If only left neighbor exists, use slide right edge
  else if (neighbors.left) {
    const leftEdge = neighbors.left.right;
    const rightEdge = slideWidth;
    const availableWidth = rightEdge - leftEdge;
    newX = leftEdge + (availableWidth - selectedWidth) / 2;
    horizontalCentered = true;
  } 
  // If only right neighbor exists, use slide left edge
  else if (neighbors.right) {
    const leftEdge = 0;
    const rightEdge = neighbors.right.left;
    const availableWidth = rightEdge - leftEdge;
    newX = leftEdge + (availableWidth - selectedWidth) / 2;
    horizontalCentered = true;
  }
  // If no horizontal neighbors, center in slide
  else {
    newX = (slideWidth - selectedWidth) / 2;
    horizontalCentered = true;
  }
  
  // Calculate vertical centering
  // If both top and bottom neighbors exist, center between them
  if (neighbors.top && neighbors.bottom) {
    const topEdge = neighbors.top.bottom;
    const bottomEdge = neighbors.bottom.top;
    const availableHeight = bottomEdge - topEdge;
    newY = topEdge + (availableHeight - selectedHeight) / 2;
    verticalCentered = true;
  }
  // If only top neighbor exists, use slide bottom edge
  else if (neighbors.top) {
    const topEdge = neighbors.top.bottom;
    const bottomEdge = slideHeight;
    const availableHeight = bottomEdge - topEdge;
    newY = topEdge + (availableHeight - selectedHeight) / 2;
    verticalCentered = true;
  }
  // If only bottom neighbor exists, use slide top edge
  else if (neighbors.bottom) {
    const topEdge = 0;
    const bottomEdge = neighbors.bottom.top;
    const availableHeight = bottomEdge - topEdge;
    newY = topEdge + (availableHeight - selectedHeight) / 2;
    verticalCentered = true;
  }
  // If no vertical neighbors, center in slide
  else {
    newY = (slideHeight - selectedHeight) / 2;
    verticalCentered = true;
  }
  
  // Step 5: Calculate the amount to move each element
  const deltaX = newX - selectedLeft;
  const deltaY = newY - selectedTop;
  
  // Apply the movement to all selected elements
  pageElements.forEach(element => {
    const currentLeft = element.getLeft();
    const currentTop = element.getTop();
    
    element.setLeft(currentLeft + deltaX);
    element.setTop(currentTop + deltaY);
  });
  
  // Log the operation for debugging
  console.log(`Centered ${pageElements.length} elements. Moved by: X=${deltaX}, Y=${deltaY}`);
  
  return true;
}

/**
 * Finds the nearest element above the selected elements
 * 
 * @param {number} selectedLeft - Left coordinate of the selection bounding box
 * @param {number} selectedTop - Top coordinate of the selection bounding box
 * @param {number} selectedRight - Right coordinate of the selection bounding box
 * @param {number} selectedBottom - Bottom coordinate of the selection bounding box
 * @param {PageElement[]} allElements - All elements on the current slide
 * @param {Set<string>} selectedElementIds - Set of IDs of the selected elements to exclude
 * @returns {Object|null} The nearest top neighbor's information or null if none found
 */
function findNearestTopNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allElements, selectedElementIds) {
  let nearestDistance = Infinity;
  let nearestElement = null;
  
  for (const element of allElements) {
    // Skip if it's one of the selected elements
    if (selectedElementIds.has(element.getObjectId())) continue;
    
    const elementBottom = element.getTop() + element.getHeight();
    const elementLeft = element.getLeft();
    const elementRight = elementLeft + element.getWidth();
    
    // Check if the element is above the selected elements
    if (elementBottom < selectedTop) {
      // Check if there's horizontal overlap
      const hasHorizontalOverlap = 
        (elementLeft <= selectedRight && elementRight >= selectedLeft);
      
      if (hasHorizontalOverlap) {
        const distance = selectedTop - elementBottom;
        if (distance < nearestDistance) {
          nearestDistance = distance;
          nearestElement = {
            element: element,
            bottom: elementBottom,
            distance: distance
          };
        }
      }
    }
  }
  
  return nearestElement;
}

/**
 * Finds the nearest element below the selected elements
 * 
 * @param {number} selectedLeft - Left coordinate of the selection bounding box
 * @param {number} selectedTop - Top coordinate of the selection bounding box
 * @param {number} selectedRight - Right coordinate of the selection bounding box
 * @param {number} selectedBottom - Bottom coordinate of the selection bounding box
 * @param {PageElement[]} allElements - All elements on the current slide
 * @param {Set<string>} selectedElementIds - Set of IDs of the selected elements to exclude
 * @returns {Object|null} The nearest bottom neighbor's information or null if none found
 */
function findNearestBottomNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allElements, selectedElementIds) {
  let nearestDistance = Infinity;
  let nearestElement = null;
  
  for (const element of allElements) {
    // Skip if it's one of the selected elements
    if (selectedElementIds.has(element.getObjectId())) continue;
    
    const elementTop = element.getTop();
    const elementLeft = element.getLeft();
    const elementRight = elementLeft + element.getWidth();
    
    // Check if the element is below the selected elements
    if (elementTop > selectedBottom) {
      // Check if there's horizontal overlap
      const hasHorizontalOverlap = 
        (elementLeft <= selectedRight && elementRight >= selectedLeft);
      
      if (hasHorizontalOverlap) {
        const distance = elementTop - selectedBottom;
        if (distance < nearestDistance) {
          nearestDistance = distance;
          nearestElement = {
            element: element,
            top: elementTop,
            distance: distance
          };
        }
      }
    }
  }
  
  return nearestElement;
}

/**
 * Finds the nearest element to the left of the selected elements
 * 
 * @param {number} selectedLeft - Left coordinate of the selection bounding box
 * @param {number} selectedTop - Top coordinate of the selection bounding box
 * @param {number} selectedRight - Right coordinate of the selection bounding box
 * @param {number} selectedBottom - Bottom coordinate of the selection bounding box
 * @param {PageElement[]} allElements - All elements on the current slide
 * @param {Set<string>} selectedElementIds - Set of IDs of the selected elements to exclude
 * @returns {Object|null} The nearest left neighbor's information or null if none found
 */
function findNearestLeftNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allElements, selectedElementIds) {
  let nearestDistance = Infinity;
  let nearestElement = null;
  
  for (const element of allElements) {
    // Skip if it's one of the selected elements
    if (selectedElementIds.has(element.getObjectId())) continue;
    
    const elementRight = element.getLeft() + element.getWidth();
    const elementTop = element.getTop();
    const elementBottom = elementTop + element.getHeight();
    
    // Check if the element is to the left of the selected elements
    if (elementRight < selectedLeft) {
      // Check if there's vertical overlap
      const hasVerticalOverlap = 
        (elementTop <= selectedBottom && elementBottom >= selectedTop);
      
      if (hasVerticalOverlap) {
        const distance = selectedLeft - elementRight;
        if (distance < nearestDistance) {
          nearestDistance = distance;
          nearestElement = {
            element: element,
            right: elementRight,
            distance: distance
          };
        }
      }
    }
  }
  
  return nearestElement;
}

/**
 * Finds the nearest element to the right of the selected elements
 * 
 * @param {number} selectedLeft - Left coordinate of the selection bounding box
 * @param {number} selectedTop - Top coordinate of the selection bounding box
 * @param {number} selectedRight - Right coordinate of the selection bounding box
 * @param {number} selectedBottom - Bottom coordinate of the selection bounding box
 * @param {PageElement[]} allElements - All elements on the current slide
 * @param {Set<string>} selectedElementIds - Set of IDs of the selected elements to exclude
 * @returns {Object|null} The nearest right neighbor's information or null if none found
 */
function findNearestRightNeighbor(selectedLeft, selectedTop, selectedRight, selectedBottom, allElements, selectedElementIds) {
  let nearestDistance = Infinity;
  let nearestElement = null;
  
  for (const element of allElements) {
    // Skip if it's one of the selected elements
    if (selectedElementIds.has(element.getObjectId())) continue;
    
    const elementLeft = element.getLeft();
    const elementTop = element.getTop();
    const elementBottom = elementTop + element.getHeight();
    
    // Check if the element is to the right of the selected elements
    if (elementLeft > selectedRight) {
      // Check if there's vertical overlap
      const hasVerticalOverlap = 
        (elementTop <= selectedBottom && elementBottom >= selectedTop);
      
      if (hasVerticalOverlap) {
        const distance = elementLeft - selectedRight;
        if (distance < nearestDistance) {
          nearestDistance = distance;
          nearestElement = {
            element: element,
            left: elementLeft,
            distance: distance
          };
        }
      }
    }
  }
  
  return nearestElement;
}

/**
 * Visualizes the padding around the selected element (for debugging)
 * Shows the distances to the nearest neighbors in all four directions
 */
function visualizePadding() {
  // Get the active presentation and selection
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const selectionType = selection.getSelectionType();
  
  // Validate that only one item or group is selected
  if (selectionType !== SlidesApp.SelectionType.PAGE_ELEMENT) {
    SlidesApp.getUi().alert('Please select a single element or group');
    return;
  }
  
  const pageElements = selection.getPageElementRange().getPageElements();
  if (pageElements.length !== 1) {
    SlidesApp.getUi().alert('Please select exactly one element or group');
    return;
  }
  
  // Get the selected element
  const selectedElement = pageElements[0];
  const currentPage = selection.getCurrentPage();
  const allPageElements = currentPage.getPageElements();
  
  // Find nearest neighbors in all four directions
  const neighbors = {
    top: findNearestTopNeighbor(selectedElement, allPageElements),
    bottom: findNearestBottomNeighbor(selectedElement, allPageElements),
    left: findNearestLeftNeighbor(selectedElement, allPageElements),
    right: findNearestRightNeighbor(selectedElement, allPageElements)
  };
  
  // Create a message with the padding information
  let message = 'Padding around the selected element:\n\n';
  
  if (neighbors.top) {
    message += `Top: ${Math.round(neighbors.top.distance)} points\n`;
  } else {
    message += 'Top: No neighbor found\n';
  }
  
  if (neighbors.bottom) {
    message += `Bottom: ${Math.round(neighbors.bottom.distance)} points\n`;
  } else {
    message += 'Bottom: No neighbor found\n';
  }
  
  if (neighbors.left) {
    message += `Left: ${Math.round(neighbors.left.distance)} points\n`;
  } else {
    message += 'Left: No neighbor found\n';
  }
  
  if (neighbors.right) {
    message += `Right: ${Math.round(neighbors.right.distance)} points\n`;
  } else {
    message += 'Right: No neighbor found\n';
  }
  
  SlidesApp.getUi().alert(message);
}
