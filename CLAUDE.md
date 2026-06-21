# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project that enhances Google Slides with automated formatting, styling, and content management features. The module provides a custom menu with various tools to improve slide design consistency and streamline presentation creation.

## Development Commands

### Initial Setup

```bash
# Install clasp globally if not already installed
npm install -g @google/clasp

# Login to Google account
clasp login

# Initialize new project (use the provided script)
./init.sh

# Or manually create project
clasp create --type slides --title "Your Presentation Title"
cp appsscript.example.json appsscript.json
clasp push
clasp open-container
```

### Development Workflow

```bash
# Push changes to Google Apps Script
clasp push

# Pull changes from Google Apps Script
clasp pull

# Open the Google Slides presentation
clasp open-container

# Open the Apps Script editor
clasp open
```

### clasp / git worktree Gotchas

- **`clasp push` scans the whole tree.** `.clasp.json` here uses `rootDir: ""`
  and `skipSubdirectories: false`, so clasp walks every subdirectory — including
  any git worktrees created under `.claude/worktrees/`.
- **Duplicate-file collision.** A worktree (e.g. `.claude/worktrees/t/`) holds a
  full copy of the source tree. Because all `.js`/`.html` files flatten to the
  same Apps Script names, clasp fails with
  `A file with this name already exists in the current project: appsscript`
  (two `appsscript.json`, etc.). Plain `clasp push` may also just print
  `Skipping push.`
- **`.claspignore`'s `.*` does NOT recurse.** The catch-all `.*` only ignores
  dotfiles/dirs at the root; it does not exclude a dotdir's *contents*. To keep
  clasp out of session files/worktrees, `.claspignore` has an explicit
  `.claude/**` entry — keep it.
- **Verify before pushing:** `clasp status` should show source files at repo-root
  paths (e.g. `src/util/grid_minter.js`), with **zero** `worktrees` entries. If
  you see `.claude/worktrees/...` lines, the ignore rule is missing or broke.
- **Don't commit the worktree.** `git status` lists `.claude/worktrees/t/` as
  untracked — never `git add` it; commit only the real source changes.
- **`clasp push --force`** skips the manifest-overwrite confirmation that the
  non-interactive shell otherwise auto-declines.
- **Never `git reset --hard` with uncommitted work present** — it silently
  discards working-tree edits (this is how an earlier CLAUDE.md edit was lost).
  To test whether a commit conflicts, use `git cherry-pick -n` then
  `git cherry-pick --abort` (not `reset --hard`), or test on a throwaway branch.
- **A global `~/.gitignore_global` can silently block `git add`** even after you
  remove the pattern from the repo `.gitignore` (e.g. it ignores `dist/`). If a
  `git add` reports "paths are ignored" but the repo `.gitignore` looks clean, run
  `git check-ignore -v <path>` — it names the actual rule + file. Fix with a
  repo-level negation (`!dist/`, `!dist/**`) and/or `git add -f` (once tracked,
  ignores no longer apply; CI runners have no global ignore so they're unaffected).
- **`appsscript.json` is gitignored — `appsscript.example.json` is the committed
  source of truth.** Both CI workflows (`clasp-push.yml`, `build-bundle.yml`)
  `cp appsscript.example.json appsscript.json` before pushing/building, since the
  real manifest is absent on the runner. So edit BOTH manifests together (scopes,
  advanced services) or CI silently ships the example's version.
- **An explicit `oauthScopes` list DISABLES scope auto-detection.** With NO
  `oauthScopes` field, Apps Script auto-detects scopes from the code at auth time
  (so `Ui.showModalDialog`/`showSidebar` silently pull in `script.container.ui`).
  The moment you list `oauthScopes` explicitly, ONLY those scopes are granted —
  auto-detection is off. A missing scope then surfaces at RUNTIME as e.g.
  `指定的權限不足，無法呼叫 Ui.showModalDialog. 必要權限: .../auth/script.container.ui`.
  This is how clones broke while the main deck (which had no explicit list, so
  auto-detect covered it) worked. Any deck opening a dialog/sidebar needs
  `https://www.googleapis.com/auth/script.container.ui` in the list. Keep the
  explicit list COMPLETE in both manifests.
- **Changing scopes forces every user to re-consent.** After a scope edit ships
  (clones self-update from `dist/bundle.json`), the next dialog/sidebar call
  triggers Google's re-authorization prompt. This is unavoidable per Google's
  OAuth rules — not a bug; it only happens once per account per scope change.

### Self-update for cloned decks

Clones pull updates themselves — no central push needed. `src/util/self_update.js`
(menu **⚙ 設定與批次 → 🔄 更新腳本**) fetches `dist/bundle.json` from GitHub raw and
overwrites the clone's own content via the Apps Script API (`projects.updateContent`,
authed with `ScriptApp.getOAuthToken()`). `dist/` is built by
`scripts/build-bundle.mjs` + `.github/workflows/build-bundle.yml` on every push to
`main` (and **committed back with `[skip ci]`** to avoid a workflow loop). `dist/` is
TRACKED in git (the pull reads it from raw). One-time per account: enable the Apps
Script API at `script.google.com/home/usersettings`. `cloned.txt` +
`batch_push_to_cloned_project.sh` remain only as a recovery fallback (force `clasp
push` to a clone whose self-update broke).

### Testing

- No automated testing framework is configured
- Testing is done manually in Google Slides after pushing code
- After making changes, run `clasp push` and test in the Google Slides presentation

## Architecture

### Core Files Structure

- **src/config.js** - Main configuration and menu creation logic
  - Contains `onOpen()` function that creates custom menus
  - Defines global configuration variables (colors, fonts, etc.)
  - Manages configuration persistence via PropertiesService

- **src/util/** - Utility functions for slide manipulation
  - Individual utility files with functions for specific tasks
  - No ES6 imports - functions are globally available in Google Apps Script

- **src/batch/** - Batch processing modules
  - Functions that process multiple slides at once
  - Uses Google Slides API batch update requests for efficiency

- **src/components/** - HTML components for sidebar interface
  - Modular HTML files included via `<?!= include() ?>` syntax
  - Contains configuration forms and style buttons
  - **flowchartSidebar.html** - Interactive flowchart creation interface

- **src/util/flowchart/** - Flowchart and hierarchical shape management
  - **main.js** - Main API functions for flowchart operations
  - **graphIdUtils.js** - Graph ID parsing, generation, and management
  - **childCreationUtils.js** - Child shape creation with positioning and styling
  - **siblingCreationUtils.js** - Sibling shape creation with layout consistency
  - **index.js** - Function exports and documentation

### Key Architecture Patterns

1. **Global Functions**: Google Apps Script doesn't support ES6 modules, so all functions are global
2. **Batch API Updates**: Uses `runRequestProcessors()` pattern to collect multiple API requests and send them as a batch
3. **HTML Service**: Uses server-side HTML templates with `<?!= include() ?>` for modular components
4. **Configuration Management**: Uses PropertiesService for persistent configuration storage
5. **Graph ID System**: Uses shape title (alt text) to store hierarchical graph IDs for flowchart management
6. **Flowchart Architecture**: Supports both LR (Left-Right) and TD (Top-Down) layout patterns

### Menu System

Three main menu categories are created in `src/config.js`:

- **🗃 批次處理 (Batch Processing)** - Functions that process multiple slides
- **🎨 加入元素 (Add Elements)** - Single slide beautification tools
- **🖖 跨頁功能 (Cross-page Functions)** - Functions that work across multiple slides

### Configuration System

Configuration is managed through:

- Global variables in `src/config.js` (defaults)
- PropertiesService for user-specific persistent settings
- Sidebar interface for real-time configuration updates

### Important Functions

- `onOpen()` - Automatically creates menus when presentation opens
- `runRequestProcessors(...)` - Batches multiple API requests for efficiency
- `createCustomMenu()` - Creates the custom menu structure
- `applyThemeToCurrentPresentation()` - Applies theme from template presentation

### Flowchart System Functions

- `createChildTop/Right/Bottom/Left()` - Creates child shapes in specified direction
- `createChildTopWithText()` - Creates child shapes with custom text content
- `createSiblingShape()` - Creates sibling shapes with proper positioning
- `showSelectedShapeGraphId()` - Displays Graph ID information for debugging
- `parseGraphId()` - Parses Graph ID format: `graph[parent](layout)[current][children]`
- `generateGraphId()` - Generates hierarchical Graph IDs with layout support
- `getShapeGraphId()` / `setShapeGraphId()` - Graph ID management via shape titles

## Development Notes

- This is a Google Apps Script project, not a Node.js project
- No package.json or npm dependencies
- Uses Google Slides API v1 (enabled in appsscript.json)
- HTML templates use server-side includes, not client-side frameworks
- All code runs in Google's V8 runtime (specified in appsscript.json)

## Configuration Variables

Key variables in `src/config.js`:

```javascript
var main_color = "#3D6869"; // Main theme color
var main_font_family = "Source Sans Pro"; // Font family
var water_mark_text = "ⓒ Hsieh-Ting Lin"; // Watermark text
var label_font_size = 14; // Font size for labels
const sourcePresentationId = "1qAZzq-..."; // Template presentation ID
```

## Google Apps Script Specifics

- Files must be .js or .gs extensions (both work the same)
- Uses HtmlService for UI components
- PropertiesService for data persistence
- SlidesApp and Slides API for slide manipulation
- No require() or import statements - everything is global

## Flowchart Features

### Graph ID System

- Stores hierarchical information in shape titles (alt text) instead of visible text
- Format: `graph[parent](layout)[current][children]`
- Examples:
  - Root: `graph[](TD)[A1][]`
  - Child: `graph[A1](TD)[B1][]`
  - Parent with children: `graph[A1](LR)[B1][C1,C2]`

### Supported Layouts

- **LR (Left-Right)**: Parent connects to children horizontally
- **TD (Top-Down)**: Parent connects to children vertically
- Layout consistency maintained across sibling relationships

### Child Creation Modes

1. **By Count**: Specify number of children to create (empty text)
2. **By Text**: Multi-line text input, each line becomes text for one child shape

### UI Components

- Collapsible Line Settings (default folded)
- Tab-based child creation interface
- Real-time Graph ID inspector
- Persistent settings via localStorage
