# Child Layout System Documentation

## Overview

The enhanced flowchart system now supports **child-specific layout annotations** in Graph IDs, allowing each child to have its own directional layout while maintaining parent-child relationships.

## Layout Types

| Layout | Description | Parent Position | Child Position |
| ------ | ----------- | --------------- | -------------- |
| **LR** | Left-Right  | Left            | Right          |
| **TD** | Top-Down    | Top             | Bottom         |
| **RL** | Right-Left  | Right           | Left           |
| **DT** | Down-Top    | Bottom          | Top            |

## Graph ID Format

### Enhanced Format

```
graph[parent](layout)[current][children]
```

### Child Layout Syntax

```
graph[A1](LR)[B1][C1:RL,C2:TD,C3]
```

- `C1:RL` - Child C1 with Right-Left layout
- `C2:TD` - Child C2 with Top-Down layout
- `C3` - Child C3 inherits parent's layout (LR)

## Key Features

### 1. Child-Specific Layouts

Each child can have its own layout annotation:

```javascript
// Children with different layouts
const children = [
  { id: "B1", layout: "LR" }, // Positioned to the right
  { id: "B2", layout: "TD" }, // Positioned below
  { id: "B3", layout: "RL" }, // Positioned to the left
  { id: "B4", layout: null }, // Inherits parent layout
];
```

### 2. Smart Parent Detection

When creating children, the system automatically finds the correct parent:

```javascript
// If you select a child shape and create another child,
// the system finds the original parent instead of nesting incorrectly
findAppropriateParent(selectedShape, slide);
```

### 3. Layout-Specific Positioning

New children are grouped by layout type for proper sibling positioning:

- **LR/RL children**: Stacked vertically
- **TD/DT children**: Arranged horizontally

### 4. Backward Compatibility

- Legacy Graph IDs still work (default to LR layout)
- Existing functions unchanged
- Graceful migration of old format

## API Functions

### Core Functions

#### `parseChildWithLayout(childStr)`

```javascript
parseChildWithLayout("B1:RL");
// Returns: { id: "B1", layout: "RL" }

parseChildWithLayout("B1");
// Returns: { id: "B1", layout: null }
```

#### `getLayoutFromDirection(direction)`

```javascript
getLayoutFromDirection("TOP"); // Returns: "DT"
getLayoutFromDirection("RIGHT"); // Returns: "LR"
getLayoutFromDirection("BOTTOM"); // Returns: "TD"
getLayoutFromDirection("LEFT"); // Returns: "RL"
```

#### `findAppropriateParent(selectedShape, slide)`

Finds the correct parent shape when a child shape is selected, preventing incorrect nesting.

### Enhanced Functions

#### `parseGraphId(graphId)`

Now returns enhanced object with child layout information:

```javascript
const parsed = parseGraphId("graph[A1](LR)[B1][C1:RL,C2:TD]");
// Returns:
// {
//   parent: "A1",
//   layout: "LR",
//   current: "B1",
//   children: [
//     { id: "C1", layout: "RL" },
//     { id: "C2", layout: "TD" }
//   ],
//   childrenIds: ["C1", "C2"]  // Backward compatibility
// }
```

#### `generateGraphId(parent, layout, current, children)`

Now accepts child objects with layout information:

```javascript
const children = [
  { id: "C1", layout: "RL" },
  { id: "C2", layout: "TD" },
  { id: "C3", layout: null },
];
generateGraphId("A1", "LR", "B1", children);
// Returns: "graph[A1](LR)[B1][C1:RL,C2:TD,C3]"
```

## Usage Examples

### Creating Children with Different Layouts

1. **Create parent shape and initialize as root**:

   ```javascript
   initializeRootGraphShape(); // Creates graph[][A1][]
   ```

2. **Create children in different directions**:

   ```javascript
   // Creates child to the right with LR layout
   createChildRight(20, "STRAIGHT", 1);

   // Creates child below with TD layout
   createChildBottom(20, "STRAIGHT", 1);

   // Creates child to the left with RL layout
   createChildLeft(20, "STRAIGHT", 1);
   ```

3. **Result**: Parent will have Graph ID:
   ```
   graph[][A1][B1:LR,B2:TD,B3:RL]
   ```

### Debugging and Inspection

```javascript
// Show Graph ID information
showSelectedShapeGraphId();

// Analyze entire slide
analyzeCurrentSlide();
```

## Testing

### Unit Tests

Run comprehensive tests with:

```bash
node test/flowchart/childLayoutSystem.test.js
```

### Manual Verification

Use the verification script in Google Apps Script console:

```javascript
// Copy contents of test/verify-child-layouts.js
runVerification();
```

## Implementation Details

### Performance Optimizations

- **Pre-calculated boundaries** for layout inference
- **Efficient parent lookup** with early termination
- **Layout-specific grouping** to avoid redundant calculations

### Error Handling

- Graceful fallbacks for invalid Graph IDs
- Backward compatibility for legacy formats
- Safe parsing with null checks

### Memory Management

- Minimal object creation during parsing
- Reuse of calculated positions
- Efficient array operations

## Migration Guide

### From Legacy Format

Old Graph IDs are automatically detected and work without changes:

```javascript
// Legacy: graph[A1][B1][C1,C2]
// New:    graph[A1](LR)[B1][C1,C2]
```

### Best Practices

1. Use layout annotations for complex flowcharts
2. Leverage smart parent detection for better UX
3. Group related children by layout type
4. Test with the verification script

## Troubleshooting

### Common Issues

- **Wrong positioning**: Check if layout matches intended direction
- **Incorrect nesting**: Verify `findAppropriateParent` is working
- **Performance slow**: Use layout-specific positioning for large hierarchies

### Debug Tools

- `showSelectedShapeGraphId()` - Inspect Graph ID details
- `analyzeCurrentSlide()` - Overview of all flowchart shapes
- `runVerification()` - Test core functionality
