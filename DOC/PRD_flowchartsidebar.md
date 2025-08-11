PRD: Refactor Flowchart Sidebar HTML

1. Background

The current flowchartSidebar.html is a single large file containing:
 • HTML layout (UI controls and sections)
 • Large amounts of CSS styles
 • Inline JavaScript with both UI logic and business logic

This monolithic structure is hard to maintain, difficult to scale, and error-prone when adding features or debugging.

The goal is to modularize the sidebar UI into smaller, logically separated files inside src/components/flowchart/, and update src/util/flowchart/main.js to assemble and render the sidebar using Option A’s server-assembly pattern.

⸻

2. Objectives
 • Split large HTML file into partials by function/UI section.
 • Move inline JS into modular script files grouped by responsibility.
 • Move CSS into a dedicated styles file.
 • Ensure all partials can be reassembled into a complete sidebar HTML at runtime.
 • Preserve current sidebar functionality (no feature regressions).
 • Improve maintainability and readability.

⸻

3. Scope

In Scope
 • Refactoring sidebar HTML, CSS, and inline JS into smaller files.
 • Updating main.js to:
 • Load partials
 • Inject styles
 • Inject JS in a single <script> block
 • Return assembled HTML to be used in showFlowchartSidebar() (in Code.gs)
 • Creating new directory structure under src/components/flowchart/.

Out of Scope
 • Changes to backend Google Apps Script services (e.g., connectSelectedShapesVertical implementation).
 • Visual redesign or feature changes (layout, styling, logic remain the same).
 • Migration of Code.gs helpers (getPartial, includeRaw) — handled separately.

⸻

4. Deliverables

4.1 Directory Structure

src/
  components/
    flowchart/
      Partials_Head.html
      Styles.html
      Section_LineSettings.html
      Section_ConnectShapes.html
      Section_CreateChildren.html
      Section_AddBG.html
      Section_StageBar.html
      Section_GraphId.html
      Scripts_UI.html
      Scripts_Actions.html
      Scripts_State.html
      Scripts_Init.html

⸻

4.2 File Breakdown

HTML Partials

 1. Partials_Head.html
 • <base>, meta tags, viewport
 2. Styles.html
 • All CSS (remove <style> wrapper)
 3. Section_LineSettings.html
 • <details> block for line settings
 4. Section_ConnectShapes.html
 • Quadrant & directional connect buttons
 5. Section_CreateChildren.html
 • Count/Text tabs, child/sibling creation controls
 6. Section_AddBG.html
 • Background color/opacity controls
 7. Section_StageBar.html
 • Stage bar insertion controls
 8. Section_GraphId.html
 • Graph ID display and refresh/clear buttons

JS Modules

 1. Scripts_UI.html
 • Toast notifications, button progress animations, tab switching, accordion mutex
 2. Scripts_Actions.html
 • All event handlers for connect/create/background/stage/graph actions (calls google.script.run)
 3. Scripts_State.html
 • LocalStorage handling for settings persistence and live UI updates (opacity displays, line count, etc.)
 4. Scripts_Init.html
 • Binds event listeners via event delegation
 • Initializes UI state on load

⸻

4.3 main.js Changes

File: /Users/htlin/lizard-gslide-module/src/util/flowchart/main.js

Responsibilities after refactor:

 1. Import helper functions getPartial and includeRaw from a utility module or define them locally if needed.
 2. Export a function buildFlowchartSidebarHtml() that:
 • Loads partials from src/components/flowchart/
 • Assembles into full HTML document string:
 • <head>: Partials_Head + <style> with Styles.html
 • <body>: Each Section_*.html partial in desired order
 • <script>: Concatenated contents of Scripts_UI, Scripts_Actions, Scripts_State, Scripts_Init
 3. Return HTML string to be consumed by showFlowchartSidebar() in Code.gs.

Pseudocode:

export function buildFlowchartSidebarHtml() {
  const head    = getPartial('src/components/flowchart/Partials_Head');
  const styles  = includeRaw('src/components/flowchart/Styles');
  const sects   = [
    'Section_LineSettings',
    'Section_ConnectShapes',
    'Section_CreateChildren',
    'Section_AddBG',
    'Section_StageBar',
    'Section_GraphId'
  ].map(name => getPartial(`src/components/flowchart/${name}`));

  const scripts = [
    'Scripts_UI',
    'Scripts_Actions',
    'Scripts_State',
    'Scripts_Init'
  ].map(name => includeRaw(`src/components/flowchart/${name}`));

  return `
<!doctype html>
    <html>
      <head>
        ${head}
        <style>${styles}</style>
      </head>
      <body>
        ${sects.join('\n')}
        <script>${scripts.join('\n')}</script>
      </body>
    </html>
  `;
}

⸻

5. Functional Requirements
1. Assembly
 • Must generate valid HTML for the sidebar at runtime from modular files.
2. UI Behavior
 • All original buttons, tabs, accordions, and input elements must function identically.
3. State Persistence
 • LocalStorage persistence for gaps, arrow types, child text, background/stage settings.
4. Event Binding
 • Must use delegated event listeners in Scripts_Init.html.
5. Google Apps Script Integration
 • All server calls (google.script.run) must remain intact and functional.

⸻

6. Acceptance Criteria
 • Sidebar loads without errors.
 • UI visually matches pre-refactor version.
 • All features (connect, create children, background, stage bar, graph ID) work as before.
 • No inline onclick or inline <script> in HTML partials.
 • HTML passes basic validation (no unclosed tags, duplicate IDs are unchanged from before).
 • JS is modular, readable, and logically separated.

⸻

7. Risks & Mitigation
 • Risk: Breakage of google.script.run calls if IDs or data-action attributes mismatch.
Mitigation: Keep element IDs and attributes unchanged during split; use data-action mapping in Scripts_Init.html.
 • Risk: CSS scoping issues.
Mitigation: Keep all CSS in a single Styles.html to maintain global styling.
