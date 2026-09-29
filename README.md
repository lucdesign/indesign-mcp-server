# InDesign MCP Server

A comprehensive **Model Context Protocol (MCP) server** for **Adobe InDesign automation** with **94 tools**. This server enables AI assistants like Claude to directly control Adobe InDesign, automating complex publishing workflows, document creation, and professional layout tasks.

## 🚀 Features

### 📄 **Document Management**
- Create, open, save, and close documents with advanced options
- Support for all standard formats (A4, A5, Letter, Legal, Custom)
- Facing pages, bleeds, slugs, and margin configuration
- Multi-page document handling

### ✍️ **Text & Typography**
- Advanced text frame creation and editing
- Character and paragraph style management
- Find/replace with GREP support
- Professional typography controls (leading, tracking, optical margin alignment)

### 🎨 **Graphics & Layout**
- Image placement with multiple fit options
- Shape creation (rectangles, ellipses with styling)
- Layer management and organization
- Color swatch creation and management (CMYK, RGB, Spot colors)

### 📊 **Tables & Data**
- Table creation and data population
- Support for headers and footers
- CSV data integration capabilities

### 📤 **Export & Production**
- Professional PDF export with presets
- Multi-format image export (PNG, JPEG, TIFF)
- EPUB export for digital publishing
- Package for print production
- Preflight checking

### 🔧 **Automation & Scripting**
- Data merge operations
- Custom ExtendScript execution
- Batch processing capabilities
- Professional publishing workflows

## 📋 Prerequisites

- **Adobe InDesign 2026** (or compatible; set `INDESIGN_APP_NAME` for another version)
- **macOS** (required for AppleScript integration)
- **Node.js 18+**
- **MCP-compatible client** (like Claude Desktop)

## 🛠️ Installation

### 1. Clone the Repository
```bash
git clone https://github.com/lucdesign/indesign-mcp-server.git
cd indesign-mcp-server
```

### 2. Install Dependencies
```bash
npm install
```

### 3. Configure MCP Client

Add the server to your MCP client configuration (e.g., Claude Desktop):

**claude_desktop_config.json:**
```json
{
  "mcpServers": {
    "indesign": {
      "command": "node",
      "args": ["/path/to/indesign-mcp-server/index.js"],
      "env": {}
    }
  }
}
```

### 4. Start Adobe InDesign

Ensure Adobe InDesign is running before using the MCP server.

## 🎯 Quick Start

### Create a Professional Document
```javascript
// Create an A4 document with facing pages
create_document({
  preset: "A4",
  orientation: "Portrait", 
  pages: 4,
  facingPages: true,
  marginTop: 20,
  marginBottom: 20,
  marginLeft: 20,
  marginRight: 20
})
```

### Add Formatted Text
```javascript
// Create a text frame with professional typography
create_text_frame({
  content: "Professional Publishing with InDesign MCP",
  x: 20,
  y: 20,
  width: 170,
  height: 50,
  paragraphStyle: "Title"   // or fontSize / fontFamily / fontStyle / alignment as local overrides
})
```

### Create and Apply Styles
```javascript
// Character style (fonts must be installed, swatches must exist)
create_character_style({
  name: "Emphasis",
  fontFamily: "Helvetica Neue",
  fontStyle: "Italic",
  textColor: "Black"
})

// Apply it to every occurrence of a text inside a frame
apply_character_style_to_range({
  frameId: 250,            // returned by create_text_frame
  styleName: "Emphasis",
  matchText: "important"
})
```

### Export Professional PDF
```javascript
export_pdf({
  filePath: "/path/to/output.pdf",
  preset: "HighQualityPrint",
  includeBleed: true,
  confirmDestructive: true   // export overwrites files, so it must be confirmed
})
```

## Concepts

### Coordinate system and units
- All positions and sizes are in **millimetres**; font sizes, leading and stroke weights are in **points**; angles in **degrees**.
- **Origin: the top-left corner of the page's trim area.** `x` grows to the right, `y` grows downwards. Every page of a spread has its own origin, so `x: 5` is always 5 mm from that page's left edge (also for the right-hand page of a facing-pages spread).
- **Negative values reach into the bleed** (for a full-bleed background of a document with 3 mm bleed use `x: -3, y: -3, width: page + 6, height: page + 6`).
- Rotation: positive angles turn **counter-clockwise**, as in the InDesign UI. For rotated objects, `x/y/w/h` reported by `list_page_items` describe the bounding box.
- The server switches the active document to millimetres and per-page ruler origin while a tool runs and restores your settings afterwards.

### Addressing objects
Every `create_*` tool and `place_image` returns the object's **id**. Use it with any tool that takes an object (`id`, or `frameId` for text tools). ids are stable: they do not change when objects are added, deleted or reordered. `pageIndex` + `index` (or `frameIndex`) are accepted as a fallback but **shift when the stacking order changes**: index 0 is the *frontmost* object, not the oldest. `list_text_frames` and `list_page_items` show both id and index.

### Text
- In `content`, `\n` makes a **paragraph break** by default. Pass `lineBreak: "forced"` to make every `\n` a forced line break (Shift+Enter), or use `<br>` inside the text for individual forced breaks.
- Local formatting is applied **only for parameters you pass** (plus the documented defaults when no `paragraphStyle` is given), so paragraph styles are never overridden by hidden defaults.
- `create_text_frame`, `edit_text_frame`, `apply_paragraph_style`, `insert_markdown_text` and `get_text_content` report `overflows=true` when the text does not fit its frame.
- `insert_markdown_text` maps Markdown to existing styles: `#`..`######` to paragraph styles `Heading 1`..`Heading 6` (also tried: `Header N`, `HN`, `Ueberschrift N`), `**bold**` to character style `Bold` (or `Strong`, `Fett`), `*italic*` to `Italic` (or `Emphasis`, `Kursiv`), `***both***` to `Bold Italic`. Normal paragraphs get `bodyStyle` if given. Override any name with `styleMap`. If a needed style is missing the tool fails with a list of the missing styles and changes nothing. Each non-blank line is one paragraph; lists, links and tables are not interpreted.

### Colour
New swatches default to **CMYK** (0-100). `hex` or RGB values are converted to CMYK in print documents **with InDesign's own colour management** (sRGB to the CMYK profile of the document, e.g. PSO Coated v3), so hues stay where they should; use `keepRgb: true` to keep RGB, or `conversion: "simple"` for the profile-free formula, which shifts hues (greens toward blue). `convert_rgb_to_cmyk` shows the converted values and tells you when a colour is outside the print gamut and how it will really look. SVG artwork is converted through the same profile at output. Presets: `rich_black`, `overprint_black`. Overprint is a property of objects (`set_object_overprint`), not of swatches.

### Export
- `export_pdf` uses a temporary copy of the chosen PDF preset so that `includeBleed` / `includeSlug` really take effect (InDesign presets otherwise override them). Preset aliases: `Print` / `HighQualityPrint` = `[High Quality Print]`, `Web` / `SmallestFileSize` = `[Smallest File Size]`, `PressQuality`, `PDFX1a`, `PDFX3`, `PDFX4`; any other value is looked up as the exact preset name (your own presets work too). The output folder is created if missing.
- **Page numbers are positions.** In `pageRange` (`export_images`, `export_pdf`) and `pageIndex`, page 1 is the first page of the document, whatever the page *names* are. Custom numbering (starting at 5), sections that restart numbering and repeated page names cannot change the meaning: positions are sent to InDesign as absolute page numbers (`+2`, `+1-+3`). Notations: `all`, `2`, `2-3`, `1,3-5`.
- `export_images` writes `<document>_page<N>.png|jpg` per page (N = the position) at the requested dpi and lists every file it wrote and the pages it exported.
- Both tools overwrite files and therefore require `confirmDestructive: true`.

### Saving
- `save_document` without `filePath` is a plain save of the active document (or of the one named by `name`). A document that was never saved gets a clear message asking for `filePath`.
- With `filePath` it is Save As: the same document is renamed and stays open. If the path is the document's own file, it is just a plain save (no confirmation, no failure). Overwriting a *different* existing file needs `confirmDestructive: true`; saving onto the file of another open document is refused. New files need no confirmation.
- "Has a file" is decided from the document's file name, not from `doc.saved`; `list_open_documents`, `get_document_info`, `close_document` and `package_document` use the same test, so they also work for documents that were reopened or have unsaved changes.

## 🛠️ Tool reference

<!-- TOOLS:START -->
_94 tools. Generated from the server's tool schemas by `npm run docs`._

### Documents

#### `get_document_info`

Get detailed information about the active document: size in mm, pages, facing pages, bleed, slug, margins, colour intent/profiles and swatches

_No parameters._

#### `create_document`

Create a new document (units: mm). Presets: A3, A4, A5, Letter, Legal, Custom. For Custom, width/height are in mm; an explicit orientation swaps them if needed (Landscape => width >= height, Portrait => height >= width); without orientation they are used as given.

| Parameter | Type | Description |
|---|---|---|
| `preset` | `A3` \| `A4` \| `A5` \| `Letter` \| `Legal` \| `Custom` | Page size preset Default: `"A4"`. |
| `width` | number | Page width in mm (Custom only) |
| `height` | number | Page height in mm (Custom only) |
| `orientation` | `Portrait` \| `Landscape` | Presets default to Portrait. For Custom, only applied when given. |
| `pages` | number | Number of pages Default: `1`. |
| `facingPages` | boolean | Enable facing pages Default: `false`. |
| `bleed` | number | Bleed on all sides in mm Default: `0`. |
| `slug` | number | Slug on all sides in mm Default: `0`. |
| `marginTop` | number | Top margin in mm Default: `20`. |
| `marginBottom` | number | Bottom margin in mm Default: `20`. |
| `marginLeft` | number | Left (inside) margin in mm Default: `20`. |
| `marginRight` | number | Right (outside) margin in mm Default: `20`. |

#### `open_document`

Open an existing InDesign document

| Parameter | Type | Description |
|---|---|---|
| `filePath` *(required)* | string | Path to the InDesign document (.indd) |

#### `save_document`

Save a document (the active one, or the one named by "name"). Without filePath: a plain save (a document that was never saved needs a filePath). With filePath: Save As, the same document is renamed and stays open; the document's own file (same path) is just a plain save; overwriting a DIFFERENT existing file needs confirmDestructive. Saving to the file of another open document is refused.

| Parameter | Type | Description |
|---|---|---|
| `filePath` | string | Optional: path of the .indd file to save to (Save As) |
| `name` | string | Name of the document as shown by list_open_documents (default: the active document) |
| `confirmDestructive` | boolean | Required only to overwrite an existing file that is not the document's own file Default: `false`. |

#### `close_document`

Close a document (the active one, or the one named by "name"). Documents with unsaved changes require save=true or confirmDestructive=true.

| Parameter | Type | Description |
|---|---|---|
| `name` | string | Document name as shown by list_open_documents (default: active document) |
| `save` | boolean | Save before closing (only for documents that already have a file) Default: `false`. |
| `confirmDestructive` | boolean | Required to discard unsaved changes Default: `false`. |

#### `list_open_documents`

List all open documents with name, page count, saved/modified state and which one is active

_No parameters._

#### `activate_document`

Make an open document the active one (by name as shown by list_open_documents)

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Document name |

#### `view_document`

Get visual representation and detailed info about the current document

_No parameters._

#### `zoom_to_page`

Zoom and fit page in view

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` | number | Page index to zoom to |
| `fitOption` | `FIT_PAGE` \| `FIT_SPREAD` \| `ACTUAL_SIZE` |  Default: `"FIT_PAGE"`. |

### Pages

#### `add_page`

Add a new page to the document

| Parameter | Type | Description |
|---|---|---|
| `position` | `before` \| `after` \| `end` |  Default: `"end"`. |
| `pageIndex` | number | Reference page index (for before/after) |
| `masterPage` | string | Master page to apply |

#### `delete_page`

Delete a page from the document

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` *(required)* | number | Page index to delete |
| `confirmDestructive` | boolean | REQUIRED: Confirm page deletion Default: `false`. |

#### `duplicate_page`

Duplicate a page with all its content (position: before/after the source page, or end of the document)

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` *(required)* | number | Page index to duplicate |
| `position` | `before` \| `after` \| `end` |  Default: `"after"`. |

#### `navigate_to_page`

Navigate to a specific page

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` *(required)* | number | Page index to navigate to |

### Seeing and measuring

#### `render_preview`

Render one page (or one object) to PNG and return it directly as an image, so you can see the result without exporting to a folder. Options: pageIndex, objectId (renders just that object), resolution in dpi (default 72), includeBleed. The longest side is limited to 4000 px.

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `objectId` | number | Render only this object instead of a page |
| `resolution` | number | Resolution in dpi (10-600) Default: `72`. |
| `includeBleed` | boolean | Include the document bleed (page previews) Default: `false`. |

#### `measure_text`

Measure a text frame without exporting: line count, the text of every line, the width of every line in mm, the frame width and the overflow state.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost) |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |

#### `list_page_items`

List all objects on a page (text frames, rectangles, ovals, polygons, lines, groups, image frames) with id, type, geometry in mm, rotation, fill and overflow state. Objects inside groups are listed with their group id.

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `type` | string | Only this type: TextFrame, Rectangle, Oval, Polygon, GraphicLine, Group |

#### `get_object_info`

Full details of an object: type, page, geometry (mm), rotation, fill (swatch, tint, values), stroke, opacity and blend mode, effects, applied object style, layer, name/label, group membership, and the anchor points of paths.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |

### Text frames and content

#### `create_text_frame`

Create a text frame (mm, origin top-left of the page trim). Only parameters you pass are applied as local formatting; defaults (Helvetica Neue 12 pt, Black, left) are used only when no paragraphStyle is given. Returns the frame id and whether the text overflows. In content, "\n" makes a paragraph break unless lineBreak is "forced"; "<br>" always makes a forced line break (Shift+Enter).

| Parameter | Type | Description |
|---|---|---|
| `content` *(required)* | string | Text content. "\n" = paragraph break (default), "<br>" = forced line break |
| `lineBreak` | `paragraph` \| `forced` | How "\n" in content is interpreted Default: `"paragraph"`. |
| `x` | number | X position in mm Default: `10`. |
| `y` | number | Y position in mm Default: `10`. |
| `width` | number | Width in mm Default: `100`. |
| `height` | number | Height in mm Default: `50`. |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fontSize` | number | Font size in points (local override) |
| `fontFamily` | string | Font family name (local override) |
| `fontStyle` | string | Font style (Regular, Bold, Italic, ...) used with fontFamily Default: `"Regular"`. |
| `textColor` | string | Swatch name (local override) |
| `alignment` | `LEFT_ALIGN` \| `CENTER_ALIGN` \| `RIGHT_ALIGN` \| `JUSTIFY` | Paragraph alignment (local override) |
| `paragraphStyle` | string | Paragraph style name to apply (error if missing) |
| `characterStyle` | string | Character style name to apply to all text |

#### `edit_text_frame`

Edit a text frame. Setting content replaces the whole story. Returns whether the text overflows. Use frameId (from create_text_frame / list_text_frames) or pageIndex + frameIndex.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId. |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `content` | string | New text (replaces the whole story). "\n" = paragraph break, "<br>" = forced line break |
| `lineBreak` | `paragraph` \| `forced` | How "\n" in content is interpreted Default: `"paragraph"`. |
| `fontSize` | number | Font size in points |
| `fontFamily` | string | Font family name |
| `fontStyle` | string | Font style used with fontFamily Default: `"Regular"`. |
| `textColor` | string | Swatch name |
| `alignment` | `LEFT_ALIGN` \| `CENTER_ALIGN` \| `RIGHT_ALIGN` \| `JUSTIFY` |  |

#### `list_text_frames`

List the text frames on a page: stable id, index, geometry (mm), overflow state and a text preview. Ordering: index 0 is the frontmost frame (stacking order), NOT creation order - indices shift when frames are added or reordered, so use the id.

| Parameter | Type | Description |
|---|---|---|
| `pageIndex` | number | Page index (0-based) Default: `0`. |

#### `get_text_content`

Get the text of a text frame (selection, frameId, or frameIndex). Also reports paragraph count and overflow state.

| Parameter | Type | Description |
|---|---|---|
| `normalizeSpaces` | boolean | Collapse line breaks and repeated spaces Default: `true`. |
| `frameId` | number | Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId. |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `maxLength` | number | Truncate output to this many characters (0 = no limit) Default: `0`. |

#### `insert_markdown_text`

Insert Markdown into a text frame, mapped to EXISTING styles. Each non-blank line is one paragraph. Supported: "# " to "###### " headers -> paragraph styles "Heading 1".."Heading 6" (aliases tried: "Header N", "HN", "Ueberschrift N"); **bold** -> character style "Bold" (alias "Strong"); *italic* -> "Italic" (alias "Emphasis"); ***both*** -> both styles. Plain paragraphs keep the frame formatting unless bodyStyle is given. Override any style name with styleMap. Fails with a clear error (and changes nothing) if a required style is missing.

| Parameter | Type | Description |
|---|---|---|
| `markdownText` *(required)* | string | Markdown text to insert |
| `frameId` | number | Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId. |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `useSelectedFrame` | boolean | Use the currently selected text frame Default: `false`. |
| `replaceContent` | boolean | Replace the whole story (true) or append (false) Default: `true`. |
| `bodyStyle` | string | Paragraph style for normal paragraphs (optional) |
| `styleMap` | object | Override style names: { h1..h6, body, bold, italic } (paragraph styles for h1-h6/body, character styles for bold/italic) |

#### `find_replace_text`

Find and replace text. scope: document, story (the story of the current selection) or selection (needs a text selection)

| Parameter | Type | Description |
|---|---|---|
| `findText` *(required)* | string | Text to find |
| `replaceText` *(required)* | string | Replacement text |
| `caseSensitive` | boolean | Case sensitive search Default: `false`. |
| `wholeWord` | boolean | Whole word only Default: `false`. |
| `useGrep` | boolean | Use GREP (regular expressions) Default: `false`. |
| `scope` | `document` \| `story` \| `selection` |  Default: `"document"`. |

#### `get_selected_objects`

Get information about currently selected objects in InDesign. ESSENTIAL for working with user-selected text frames.

_No parameters._

#### `analyze_embedded_objects`

Analyze embedded objects (MathML formulas, graphics, etc.) in selected text frame or specified frame

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Text frame index (optional if frame is selected) |
| `pageIndex` | number | Page index Default: `0`. |
| `maxObjects` | number | Maximum number of objects to analyze Default: `5`. |

### Text formatting and typography

#### `apply_character_style_to_range`

Apply a character style to a character range of a text frame (e.g. bold words inside a paragraph). Give startIndex/endIndex (inclusive, characters within the story) or matchText to style every occurrence of a text.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable id of the text frame. Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost) |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `styleName` *(required)* | string | Character style name |
| `startIndex` | number | First character (0-based) |
| `endIndex` | number | Last character (inclusive) |
| `matchText` | string | Style every occurrence of this text instead of a range |

#### `clear_overrides`

Remove local formatting overrides from a whole text frame or a character range so the paragraph/character styles show through

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable id of the text frame. Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost) |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `startIndex` | number | First character (0-based); omit for the whole frame |
| `endIndex` | number | Last character (inclusive) |

#### `fix_typography_in_selection`

Fix typography in selected text or story. Corrects dates (DD.MM.YYYY with thin spaces), quotes, dashes, and other typographic elements.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Text frame index to fix (use get_selected_objects to work with selection) |
| `pageIndex` | number | Page index Default: `0`. |
| `useSelectedFrame` | boolean | Use currently selected text frame Default: `false`. |
| `fixDates` | boolean | Fix date spacing (DD. MM. YYYY) Default: `true`. |
| `fixQuotes` | boolean | Fix quotes to typographic quotes Default: `true`. |
| `fixDashes` | boolean | Fix hyphens to em/en dashes Default: `true`. |
| `fixSpaces` | boolean | Fix multiple spaces and trailing spaces Default: `true`. |

#### `find_typography_issues`

Analyze text for common typography issues (wrong spaces in dates, straight quotes, double spaces, etc.)

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Text frame index to analyze |
| `pageIndex` | number | Page index Default: `0`. |
| `useSelectedFrame` | boolean | Analyze currently selected text frame Default: `false`. |

#### `clean_imported_text`

Clean imported text from common typography sins: double paragraph breaks, line breaks instead of paragraphs, trailing spaces, hyphens instead of dashes, manual formatting, bullet lists, hardcoded chapter numbers, etc.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Text frame index to clean |
| `pageIndex` | number | Page index Default: `0`. |
| `useSelectedFrame` | boolean | Clean currently selected text frame Default: `false`. |
| `fixParagraphs` | boolean | Fix double paragraph breaks and line breaks Default: `true`. |
| `fixDashes` | boolean | Fix hyphens to proper n-dashes for ranges/thoughts Default: `true`. |
| `fixLists` | boolean | Remove manual bullet lists and dashes Default: `true`. |
| `fixFormatting` | boolean | Remove manual bold/italic (prepare for character styles) Default: `true`. |
| `fixChapterNumbers` | boolean | Remove hardcoded chapter numbers Default: `true`. |
| `fixSpaces` | boolean | Remove trailing spaces and multiple spaces Default: `true`. |

#### `analyze_text_problems`

Analyze imported text for common problems before cleaning. Shows what issues exist.

| Parameter | Type | Description |
|---|---|---|
| `frameId` | number | Stable object id of the text frame (preferred over frameIndex) |
| `frameIndex` | number | Text frame index to analyze |
| `pageIndex` | number | Page index Default: `0`. |
| `useSelectedFrame` | boolean | Analyze currently selected text frame Default: `false`. |

#### `list_grep_searches`

List all saved GREP searches in the document (like DATUM search for dates)

_No parameters._

#### `list_fonts`

List installed font families with their styles (use them for fontFamily / fontStyle). Filter with search (substring of the family name) or family (exact name); limit caps the number of families shown.

| Parameter | Type | Description |
|---|---|---|
| `search` | string | Case-insensitive substring of the family name |
| `family` | string | Exact family name: list only its styles |
| `limit` | number | Maximum number of families (1-1000) Default: `60`. |

### Styles

#### `create_paragraph_style`

Create a paragraph style. Returns the font that was actually applied.

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Style name |
| `baseStyle` | string | Base style to inherit from |
| `fontFamily` | string | Font family (error if not installed) |
| `fontStyle` | string | Font style: Regular, Medium, Bold, Italic, ... |
| `fontSize` | number | Font size in points |
| `leading` | number | Leading in points |
| `tracking` | number | Tracking in 1/1000 em |
| `kerning` | `Metrics` \| `Optical` | Kerning method |
| `spaceBefore` | number | Space before in mm |
| `spaceBeforePt` | number | Space before in points (alternative to spaceBefore) |
| `spaceAfter` | number | Space after in mm |
| `spaceAfterPt` | number | Space after in points (alternative to spaceAfter) |
| `firstLineIndent` | number | First-line indent in mm |
| `leftIndent` | number | Left indent in mm |
| `rightIndent` | number | Right indent in mm |
| `alignment` | `LEFT_ALIGN` \| `CENTER_ALIGN` \| `RIGHT_ALIGN` \| `JUSTIFY` \| `LEFT_JUSTIFIED` \| `CENTER_JUSTIFIED` \| `RIGHT_JUSTIFIED` \| `FULLY_JUSTIFIED` |  |
| `hyphenation` | boolean | Enable hyphenation |
| `keepWithNext` | number | Keep with next N lines (0 = off) |
| `keepLinesTogether` | boolean | Keep lines together |
| `keepFirstLines` | number | Keep first N lines together |
| `keepLastLines` | number | Keep last N lines together |
| `keepWithPrevious` | boolean | Keep with previous paragraph |
| `composer` | `paragraph` \| `single-line` | Composer: "paragraph" = Adobe Paragraph Composer (default look), "single-line" = Adobe Single-line Composer |
| `capitalization` | `normal` \| `small_caps` \| `all_caps` \| `cap_to_small_cap` \| `lower_case` | Capitalisation |
| `ligatures` | boolean | Standard ligatures on/off |
| `openType` | object | OpenType features: booleans discretionaryLigatures, contextualAlternates, swash, titling, fractions, ordinals, slashedZero, historical; figureStyle (default, tabular_lining, proportional_oldstyle, proportional_lining, tabular_oldstyle); stylisticSets (bit mask number) |
| `hyphenateWordsLongerThan` | number | Hyphenation: minimum word length |
| `hyphenateAfterFirst` | number | Hyphenation: minimum letters before the hyphen |
| `hyphenateBeforeLast` | number | Hyphenation: minimum letters after the hyphen |
| `hyphenateLadderLimit` | number | Hyphenation: maximum consecutive hyphenated lines |
| `hyphenateCapitalizedWords` | boolean | Hyphenation: hyphenate capitalised words |
| `hyphenateLastWord` | boolean | Hyphenation: hyphenate the last word of a paragraph |
| `textColor` | string | Swatch name |

#### `modify_paragraph_style`

Modify a paragraph style (same properties as create_paragraph_style). Returns the font that was actually applied.

| Parameter | Type | Description |
|---|---|---|
| `styleName` *(required)* | string | Paragraph style name to modify |
| `baseStyle` | string | New base style |
| `fontFamily` | string | Font family (error if not installed) |
| `fontStyle` | string | Font style: Regular, Medium, Bold, Italic, ... |
| `fontSize` | number | Font size in points |
| `leading` | number | Leading in points |
| `tracking` | number | Tracking in 1/1000 em |
| `kerning` | `Metrics` \| `Optical` | Kerning method |
| `spaceBefore` | number | Space before in mm |
| `spaceBeforePt` | number | Space before in points |
| `spaceAfter` | number | Space after in mm |
| `spaceAfterPt` | number | Space after in points |
| `firstLineIndent` | number | First-line indent in mm |
| `leftIndent` | number | Left indent in mm |
| `rightIndent` | number | Right indent in mm |
| `alignment` | `LEFT_ALIGN` \| `CENTER_ALIGN` \| `RIGHT_ALIGN` \| `JUSTIFY` \| `LEFT_JUSTIFIED` \| `CENTER_JUSTIFIED` \| `RIGHT_JUSTIFIED` \| `FULLY_JUSTIFIED` |  |
| `hyphenation` | boolean | Enable hyphenation |
| `keepWithNext` | number | Keep with next N lines (0 = off) |
| `keepLinesTogether` | boolean | Keep lines together |
| `keepFirstLines` | number | Keep first N lines together |
| `keepLastLines` | number | Keep last N lines together |
| `keepWithPrevious` | boolean | Keep with previous paragraph |
| `composer` | `paragraph` \| `single-line` | Composer: "paragraph" = Adobe Paragraph Composer (default look), "single-line" = Adobe Single-line Composer |
| `capitalization` | `normal` \| `small_caps` \| `all_caps` \| `cap_to_small_cap` \| `lower_case` | Capitalisation |
| `ligatures` | boolean | Standard ligatures on/off |
| `openType` | object | OpenType features: booleans discretionaryLigatures, contextualAlternates, swash, titling, fractions, ordinals, slashedZero, historical; figureStyle (default, tabular_lining, proportional_oldstyle, proportional_lining, tabular_oldstyle); stylisticSets (bit mask number) |
| `hyphenateWordsLongerThan` | number | Hyphenation: minimum word length |
| `hyphenateAfterFirst` | number | Hyphenation: minimum letters before the hyphen |
| `hyphenateBeforeLast` | number | Hyphenation: minimum letters after the hyphen |
| `hyphenateLadderLimit` | number | Hyphenation: maximum consecutive hyphenated lines |
| `hyphenateCapitalizedWords` | boolean | Hyphenation: hyphenate capitalised words |
| `hyphenateLastWord` | boolean | Hyphenation: hyphenate the last word of a paragraph |
| `textColor` | string | Swatch name |

#### `apply_paragraph_style`

Apply a paragraph style to a text frame (or a character range of it)

| Parameter | Type | Description |
|---|---|---|
| `styleName` *(required)* | string | Paragraph style name |
| `frameId` | number | Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex. |
| `frameIndex` | number | Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId. |
| `pageIndex` | number | Page index (0-based), used with frameIndex Default: `0`. |
| `startIndex` | number | Start character index (optional) |
| `endIndex` | number | End character index (optional) |

#### `create_character_style`

Create a new character style

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Style name |
| `fontFamily` | string | Font family |
| `fontStyle` | string | Font style (Regular, Bold, Italic) |
| `fontSize` | number | Font size in points |
| `textColor` | string | Text color |
| `tracking` | number | Character tracking |
| `baseStyle` | string | Base style to inherit from |

#### `modify_character_style`

Modify properties of an existing character style

| Parameter | Type | Description |
|---|---|---|
| `styleName` *(required)* | string | Character style name to modify |
| `fontFamily` | string | Font family |
| `fontStyle` | string | Font style (Regular, Bold, Italic) |
| `fontSize` | number | Font size in points |
| `textColor` | string | Text color (swatch name) |
| `tracking` | number | Character tracking |

#### `create_object_style`

Create a new object style

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Style name |
| `fillColor` | string | Fill color (swatch name) |
| `strokeColor` | string | Stroke color (swatch name) |
| `strokeWidth` | number | Stroke width in points |
| `transparency` | number | Transparency percentage (0-100) |
| `baseStyle` | string | Base style to inherit from |

#### `modify_object_style`

Modify properties of an existing object style

| Parameter | Type | Description |
|---|---|---|
| `styleName` *(required)* | string | Object style name to modify |
| `fillColor` | string | Fill color (swatch name) |
| `strokeColor` | string | Stroke color (swatch name) |
| `strokeWidth` | number | Stroke width in points |
| `transparency` | number | Transparency percentage (0-100) |

#### `apply_object_style`

Apply an object style to an object (objectId, or pageIndex + objectIndex) or, if none is given, to the current selection

| Parameter | Type | Description |
|---|---|---|
| `styleName` *(required)* | string | Object style name |
| `objectId` | number | Stable object id (preferred) |
| `objectIndex` | number | Index in list_page_items for that page (optional if objects are selected) |
| `pageIndex` | number | Page index Default: `0`. |

#### `list_styles`

List all available styles in the document

| Parameter | Type | Description |
|---|---|---|
| `styleType` | `paragraph` \| `character` \| `object` \| `all` |  Default: `"all"`. |

### Vector shapes and paths

#### `create_path`

Create an editable open or closed Bezier path from anchor points. Each point: x, y (mm) and optionally leftDirection / rightDirection (handle positions, same coordinate system) and pointType. Returns the object id and the resulting points. Example (S-curve band): points [{x:0,y:8,rightDirection:{x:40,y:0}}, {x:120,y:8,leftDirection:{x:80,y:14},rightDirection:{x:160,y:2}}, ...], closed true.

| Parameter | Type | Description |
|---|---|---|
| `points` *(required)* | object[] | At least 2 anchor points |
| `closed` | boolean | Close the path Default: `false`. |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fill` | string | Fill swatch name, or "none" |
| `stroke` | string | Stroke swatch name, or "none" |
| `strokeWeight` | number | Stroke weight in points |
| `opacity` | number | Object opacity in percent (0-100) |
| `name` | string | Object name (shown in the Layers panel) |
| `label` | string | Script label (shown by list_page_items) |

#### `create_path_from_svg`

Create native, editable path(s) from SVG path data (commands M m L l H h V v C c S s Q q T t A a Z z; quadratics and arcs become cubic Beziers). The data is scaled from viewBox (default: the path bounds) into the target rectangle x, y, width, height (mm); one dimension only keeps the aspect ratio. Several sub-paths become one compound path. Example: d "M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z", viewBox "0 0 216 50", x -3, y 55, width 216, height 48.

| Parameter | Type | Description |
|---|---|---|
| `d` *(required)* | string | SVG path data |
| `viewBox` | string | SVG viewBox "minX minY width height" the data was drawn in (default: bounds of the path) |
| `x` | number | Target left edge in mm Default: `0`. |
| `y` | number | Target top edge in mm Default: `0`. |
| `width` | number | Target width in mm |
| `height` | number | Target height in mm |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fill` | string | Fill swatch name, or "none" |
| `stroke` | string | Stroke swatch name, or "none" |
| `strokeWeight` | number | Stroke weight in points |
| `opacity` | number | Object opacity in percent (0-100) |
| `name` | string | Object name (shown in the Layers panel) |
| `label` | string | Script label (shown by list_page_items) |

#### `edit_path_points`

Read or edit the anchor points and handles of an existing path (any rectangle, oval, polygon or line). action: read \| add \| move \| delete \| set_type \| reverse \| set_closed. Coordinates are mm (same system as create_path). Moving an anchor moves its handles with it unless moveHandles is false.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `action` | `read` \| `add` \| `move` \| `delete` \| `set_type` \| `reverse` \| `set_closed` |  Default: `"read"`. |
| `pathIndex` | number | Sub-path of a compound path (0-based) Default: `0`. |
| `pointIndex` | number | Point to move / delete / retype; for add: insert before this index (default: append) |
| `point` | object | add: the new point { x, y, leftDirection?, rightDirection?, pointType? } |
| `x` | number | move: new anchor x in mm |
| `y` | number | move: new anchor y in mm |
| `leftDirection` | object | move: new absolute position of the left handle |
| `rightDirection` | object | move: new absolute position of the right handle |
| `moveHandles` | boolean | move: translate the handles together with the anchor Default: `true`. |
| `pointType` | `corner` \| `smooth` \| `symmetrical` \| `plain` \| `auto` | set_type: new point type |
| `closed` | boolean | set_closed: true closes, false opens the path |

#### `create_line`

Create a straight or curved line between two points (mm). Curve: give controlPoint1 and/or controlPoint2 (Bezier handles). Options: weight (pt), stroke swatch, dash (solid, a stroke style name, or [dash, gap, ...] in pt), endCap, arrowheads (none, simple, simple-wide, triangle, triangle-wide, barbed, curved, circle, circle-solid, square, square-solid, bar).

| Parameter | Type | Description |
|---|---|---|
| `x1` *(required)* | number | Start x in mm |
| `y1` *(required)* | number | Start y in mm |
| `x2` *(required)* | number | End x in mm |
| `y2` *(required)* | number | End y in mm |
| `controlPoint1` | object | Curve: handle leaving the start point |
| `controlPoint2` | object | Curve: handle arriving at the end point |
| `weight` | number | Stroke weight in points |
| `stroke` | string | Stroke swatch name |
| `dash` | number[] | Custom dash/gap lengths in pt, e.g. [6, 3]. Use dashStyle for a named style or solid. |
| `dashStyle` | string | "solid" or the name of a stroke style (e.g. Dotted, Dashed) |
| `endCap` | `butt` \| `round` \| `projecting` |  |
| `arrowStart` | `none` \| `simple` \| `simple-wide` \| `triangle` \| `triangle-wide` \| `barbed` \| `curved` \| `circle` \| `circle-solid` \| `square` \| `square-solid` \| `bar` | Arrowhead at the start |
| `arrowEnd` | `none` \| `simple` \| `simple-wide` \| `triangle` \| `triangle-wide` \| `barbed` \| `curved` \| `circle` \| `circle-solid` \| `square` \| `square-solid` \| `bar` | Arrowhead at the end |
| `opacity` | number | Opacity in percent (0-100) |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `name` | string | Object name (shown in the Layers panel) |
| `label` | string | Script label (shown by list_page_items) |

#### `create_polygon`

Create a regular polygon or a star as an editable path. x, y is the centre and radius the outer radius (mm). innerRadiusRatio (0-1) makes a star; cornerRadius (mm) rounds the corners; rotation in degrees (0 = first corner points up).

| Parameter | Type | Description |
|---|---|---|
| `x` *(required)* | number | Centre x in mm |
| `y` *(required)* | number | Centre y in mm |
| `radius` *(required)* | number | Outer radius in mm |
| `sides` | number | Number of sides (3-100) or star points Default: `5`. |
| `innerRadiusRatio` | number | Star: inner radius / outer radius (0-1) |
| `cornerRadius` | number | Round corners with this radius in mm Default: `0`. |
| `rotation` | number | Rotation in degrees Default: `0`. |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fill` | string | Fill swatch name, or "none" |
| `stroke` | string | Stroke swatch name, or "none" |
| `strokeWeight` | number | Stroke weight in points |
| `opacity` | number | Object opacity in percent (0-100) |
| `name` | string | Object name (shown in the Layers panel) |
| `label` | string | Script label (shown by list_page_items) |

#### `convert_shape`

Convert a shape with InDesign's Convert Shape: to rectangle, rounded_rectangle, beveled_rectangle, inverse_rounded_rectangle, oval, triangle, polygon (sides, insetPercent, cornerRadius), line, straight_line, open_path or closed_path. Every shape (rectangle, oval, polygon) is an editable path: use edit_path_points to change its points.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `to` *(required)* | `rectangle` \| `rounded_rectangle` \| `beveled_rectangle` \| `inverse_rounded_rectangle` \| `oval` \| `triangle` \| `polygon` \| `line` \| `straight_line` \| `open_path` \| `closed_path` |  |
| `sides` | number | polygon: number of sides |
| `insetPercent` | number | polygon: star inset in percent |
| `cornerRadius` | number | polygon / rounded shapes: corner radius in mm |

#### `set_corner_options`

Set corner effects and radius (mm): option none \| rounded \| inverse_rounded \| inset \| bevel \| fancy, for all corners or per corner (corners: { top_left: { option, radius }, ... }).

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `option` | `none` \| `rounded` \| `inverse_rounded` \| `inset` \| `bevel` \| `fancy` | Effect for all corners (unless corners is given) |
| `radius` | number | Radius in mm for all corners |
| `corners` | object | Per corner settings; keys top_left, top_right, bottom_left, bottom_right; each { option?, radius? } |

#### `pathfinder`

Pathfinder on a list of object ids; returns the id of the result. union / intersect / exclude_overlap combine all shapes. subtract: the FIRST id is kept and all other shapes are cut out of it. minus_back: the frontmost shape is kept and all shapes behind it are cut out. The source objects are consumed.

| Parameter | Type | Description |
|---|---|---|
| `operation` *(required)* | `union` \| `subtract` \| `intersect` \| `exclude_overlap` \| `minus_back` |  |
| `ids` *(required)* | number[] | Ids of the shapes (at least 2, on the same page) |

### Fill, stroke and gradients

#### `set_object_fill`

Change the fill of any existing object (rectangle, oval, path, text frame) in place, so the stacking order is kept. swatch (name or "none"), tint (0-100), opacity of the fill (0-100), or a gradient (name of a gradient swatch) with gradientAngle.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `ids` | number[] | Alternative to id: change several objects at once |
| `swatch` | string | Fill swatch name, or "none" |
| `tint` | number | Fill tint in percent (0-100) |
| `opacity` | number | Fill opacity in percent (0-100) |
| `gradient` | string | Name of a gradient swatch (see create_gradient_swatch) |
| `gradientAngle` | number | Gradient angle in degrees |

#### `set_object_stroke`

Change the stroke of any existing object in place: swatch (name or "none"), weight (pt), tint, dash (custom [dash, gap, ...] pt, or dashStyle "solid" / a stroke style name), alignment (center, inside, outside), join (miter, round, bevel), miterLimit, cap (butt, round, projecting), opacity of the stroke (0-100).

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `ids` | number[] | Alternative to id: change several objects at once |
| `swatch` | string | Stroke swatch name, or "none" |
| `weight` | number | Stroke weight in points |
| `tint` | number | Stroke tint in percent (0-100) |
| `dash` | number[] | Custom dash/gap lengths in pt, e.g. [6, 3] |
| `dashStyle` | string | "solid" or the name of a stroke style |
| `alignment` | `center` \| `inside` \| `outside` |  |
| `join` | `miter` \| `round` \| `bevel` |  |
| `miterLimit` | number | Miter limit (1-500) |
| `cap` | `butt` \| `round` \| `projecting` |  |
| `opacity` | number | Stroke opacity in percent (0-100) |

#### `create_gradient_swatch`

Create a linear or radial gradient swatch from colour stops. stops: [{ swatch, position (0-100), midpoint? (13-87, between this and the previous stop) }]; the first stop must be at 0 and the last at 100. Apply it with set_object_fill (gradient, gradientAngle).

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Gradient name |
| `type` | `linear` \| `radial` |  Default: `"linear"`. |
| `stops` *(required)* | object[] | At least 2 stops |

#### `set_gradient_feather`

Fade an object out with a gradient feather (opacity gradient): type linear/radial, angle, startOpacity (default 100) and endOpacity (default 0). remove: true switches it off.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `type` | `linear` \| `radial` |  Default: `"linear"`. |
| `angle` | number | Angle in degrees Default: `90`. |
| `startOpacity` | number | Opacity at the start (0-100) Default: `100`. |
| `endOpacity` | number | Opacity at the end (0-100) Default: `0`. |
| `remove` | boolean | Remove the gradient feather Default: `false`. |

#### `set_object_opacity`

Set opacity (0-100) and/or blend mode of an object, its fill or its stroke (e.g. semi-transparent colour bands).

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |
| `opacity` | number | Opacity in percent (0-100) |
| `blendMode` | string | BlendMode name: NORMAL, MULTIPLY, SCREEN, OVERLAY, SOFT_LIGHT, HARD_LIGHT, COLOR_DODGE, COLOR_BURN, DARKEN, LIGHTEN, DIFFERENCE, EXCLUSION, HUE, SATURATION, COLOR, LUMINOSITY |
| `target` | `object` \| `fill` \| `stroke` | What the setting applies to Default: `"object"`. |

#### `set_object_overprint`

Set overprint for the fill and/or stroke of an object (print production)

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |
| `fill` | boolean | Overprint fill |
| `stroke` | boolean | Overprint stroke |

### Objects and layout

#### `create_rectangle`

Create a rectangle shape

| Parameter | Type | Description |
|---|---|---|
| `x` *(required)* | number | X position in mm |
| `y` *(required)* | number | Y position in mm |
| `width` *(required)* | number | Width in mm |
| `height` *(required)* | number | Height in mm |
| `pageIndex` | number | Page index Default: `0`. |
| `fillColor` | string | Fill color (RGB hex or swatch name) |
| `strokeColor` | string | Stroke color |
| `strokeWidth` | number | Stroke width in points Default: `1`. |
| `cornerRadius` | number | Corner radius in mm Default: `0`. |

#### `create_ellipse`

Create an ellipse shape

| Parameter | Type | Description |
|---|---|---|
| `x` *(required)* | number | X position in mm |
| `y` *(required)* | number | Y position in mm |
| `width` *(required)* | number | Width in mm |
| `height` *(required)* | number | Height in mm |
| `pageIndex` | number | Page index Default: `0`. |
| `fillColor` | string | Fill color |
| `strokeColor` | string | Stroke color |
| `strokeWidth` | number | Stroke width in points Default: `1`. |

#### `set_object_geometry`

Move and/or resize an object (x, y, width, height in mm, top-left of the unrotated frame); optionally set absolute rotation and shear (degrees). Omitted values are kept. Rotation is applied about the object centre.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |
| `x` | number | Left edge in mm |
| `y` | number | Top edge in mm |
| `width` | number | Width in mm |
| `height` | number | Height in mm |
| `rotation` | number | Absolute rotation in degrees (counter-clockwise positive) |
| `shear` | number | Absolute shear angle in degrees |

#### `rotate_object`

Rotate an object. Positive angles turn counter-clockwise (as in the InDesign UI). By default the angle is added to the current rotation; with absolute=true it sets the rotation. reference chooses the pivot point.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |
| `angle` *(required)* | number | Degrees, counter-clockwise positive |
| `absolute` | boolean | Set the rotation instead of adding to it Default: `false`. |
| `reference` | `center` \| `top-left` \| `top` \| `top-right` \| `left` \| `right` \| `bottom-left` \| `bottom` \| `bottom-right` | Pivot point Default: `"center"`. |

#### `delete_object`

Delete a text frame, rectangle, ellipse, image frame or group

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `duplicate_object`

Duplicate an object (count copies, each moved by offsetX / offsetY mm relative to the previous one, optionally onto another page). Returns the ids of the copies.

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (returned by every create_* tool, place_image and list_page_items) |
| `pageIndex` | number | Page index (0-based); only used with index |
| `index` | number | Position in list_page_items for that page (shifts when objects are added or reordered) |
| `count` | number | Number of copies (1-100) Default: `1`. |
| `offsetX` | number | Horizontal offset in mm between copies Default: `0`. |
| `offsetY` | number | Vertical offset in mm between copies Default: `0`. |
| `toPageIndex` | number | Put the copies on this page (0-based) |

#### `send_to_back`

Send an object to the back of the stacking order

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `bring_to_front`

Bring an object to the front of the stacking order

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `send_backward`

Move an object one step back in the stacking order

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `bring_forward`

Move an object one step forward in the stacking order

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `group_objects`

Group objects (by id) and return the group id

| Parameter | Type | Description |
|---|---|---|
| `ids` *(required)* | number[] | Ids of the objects to group (at least 2) |

#### `ungroup`

Ungroup a group; returns the ids of the released objects

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |

#### `align_objects`

Align objects by id. alignment: left, right, top, bottom, horizontal_center, vertical_center. relativeTo: selection (bounds of the objects), page, margins, spread, bleed, key_object (with keyObjectId).

| Parameter | Type | Description |
|---|---|---|
| `ids` *(required)* | number[] | Object ids |
| `alignment` *(required)* | `left` \| `right` \| `top` \| `bottom` \| `horizontal_center` \| `vertical_center` |  |
| `relativeTo` | `selection` \| `page` \| `margins` \| `spread` \| `bleed` \| `key_object` |  Default: `"selection"`. |
| `keyObjectId` | number | relativeTo key_object: the object that stays where it is |

#### `distribute_objects`

Distribute objects by id. distribution: left_edges, horizontal_centers, right_edges, top_edges, vertical_centers, bottom_edges, horizontal_space, vertical_space. With spacing (mm) the gaps are fixed (horizontal_space / vertical_space).

| Parameter | Type | Description |
|---|---|---|
| `ids` *(required)* | number[] | Object ids (at least 3 for a useful result) |
| `distribution` *(required)* | `left_edges` \| `horizontal_centers` \| `right_edges` \| `top_edges` \| `vertical_centers` \| `bottom_edges` \| `horizontal_space` \| `vertical_space` |  |
| `spacing` | number | Fixed gap in mm (horizontal_space / vertical_space) |
| `relativeTo` | `selection` \| `page` \| `margins` \| `spread` \| `bleed` |  Default: `"selection"`. |

#### `set_text_frame_options`

Set text frame options: insets (mm), vertical alignment, columns, auto-size and text wrap (wrap also works on other objects so text flows around them).

| Parameter | Type | Description |
|---|---|---|
| `id` | number | Stable object id (from create_* tools, place_image or list_page_items). Preferred. |
| `pageIndex` | number | Page index (0-based); used with index Default: `0`. |
| `index` | number | Position in list_page_items for that page (changes when objects are added/reordered) |
| `inset` | number | Inset on all sides in mm |
| `insetTop` | number | Top inset in mm |
| `insetLeft` | number | Left inset in mm |
| `insetBottom` | number | Bottom inset in mm |
| `insetRight` | number | Right inset in mm |
| `verticalAlignment` | `TOP` \| `CENTER` \| `BOTTOM` \| `JUSTIFY` |  |
| `columns` | number | Number of columns |
| `columnGutter` | number | Column gutter in mm |
| `autosize` | `OFF` \| `HEIGHT_ONLY` \| `WIDTH_ONLY` \| `HEIGHT_AND_WIDTH` \| `HEIGHT_AND_WIDTH_PROPORTIONALLY` | Auto-sizing type |
| `textWrap` | `NONE` \| `BOUNDING_BOX` \| `CONTOUR` \| `JUMP_OBJECT` \| `NEXT_COLUMN` | How text wraps around this object |
| `textWrapOffset` | number | Text wrap offset on all sides in mm |

### Images and graphics

#### `place_image`

Place an image into a new frame (mm). Give width and height for a fixed frame, only width or height to keep the aspect ratio, or neither to use the image size. Returns the frame id.

| Parameter | Type | Description |
|---|---|---|
| `imagePath` *(required)* | string | Path to the image file |
| `x` | number | X position in mm Default: `10`. |
| `y` | number | Y position in mm Default: `10`. |
| `width` | number | Frame width in mm |
| `height` | number | Frame height in mm |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fitOption` | `FILL_PROPORTIONALLY` \| `PROPORTIONALLY` \| `CONTENT_TO_FRAME` \| `FRAME_TO_CONTENT` \| `CENTER_CONTENT` \| `NONE` | FILL_PROPORTIONALLY = crop to fill the frame; PROPORTIONALLY = fit inside; NONE = keep 100% Default: `"PROPORTIONALLY"`. |
| `contentOffsetX` | number | Move the image inside the frame by this many mm (after fitting) |
| `contentOffsetY` | number | Move the image inside the frame by this many mm (after fitting) |
| `contentScale` | number | Scale the image content to this percentage (after fitting) |
| `createFrame` | boolean | Create a frame first (false = place at x,y with its own size) Default: `true`. |

#### `place_image_in_shape`

Place an image into an existing shape (path, rectangle, oval, polygon) which acts as the mask. fitOption FILL_PROPORTIONALLY (crop to fill, default), PROPORTIONALLY, CONTENT_TO_FRAME, CENTER_CONTENT or NONE; then optional contentOffsetX/Y (mm) and contentScale (%). Returns the shape id.

| Parameter | Type | Description |
|---|---|---|
| `shapeId` *(required)* | number | Id of the shape that becomes the mask |
| `imagePath` *(required)* | string | Path of the image file |
| `fitOption` | `FILL_PROPORTIONALLY` \| `PROPORTIONALLY` \| `CONTENT_TO_FRAME` \| `CENTER_CONTENT` \| `NONE` |  Default: `"FILL_PROPORTIONALLY"`. |
| `contentOffsetX` | number | Move the image inside the shape (mm) |
| `contentOffsetY` | number | Move the image inside the shape (mm) |
| `contentScale` | number | Scale of the image content in percent |

#### `place_graphic`

Place a vector graphic (.svg, .ai, .eps, .pdf) into a new frame (mm). Give width and height for a fixed frame, only one for the aspect ratio, or neither for the natural size. PDF/AI: pdfPage (default 1) and crop (pdf, art, trim, bleed, media, content_visible, content_all). Returns the frame id and whether the result is linked; embed: true unlinks and embeds it.

| Parameter | Type | Description |
|---|---|---|
| `filePath` *(required)* | string | Path of the .svg, .ai, .eps or .pdf file |
| `x` | number | X position in mm Default: `10`. |
| `y` | number | Y position in mm Default: `10`. |
| `width` | number | Frame width in mm |
| `height` | number | Frame height in mm |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fitOption` | `PROPORTIONALLY` \| `FILL_PROPORTIONALLY` \| `CONTENT_TO_FRAME` \| `FRAME_TO_CONTENT` \| `CENTER_CONTENT` \| `NONE` | How the graphic fits the frame (with width and height): PROPORTIONALLY = fit inside, FILL_PROPORTIONALLY = crop to fill Default: `"PROPORTIONALLY"`. |
| `embed` | boolean | Unlink and embed the graphic in the document Default: `false`. |
| `pdfPage` | number | PDF / AI: page number to place (1-based) Default: `1`. |
| `crop` | `pdf` \| `art` \| `trim` \| `bleed` \| `media` \| `content_visible` \| `content_all` | PDF / AI: crop to Default: `"pdf"`. |

#### `create_graphic_from_svg`

Place SVG markup (a string) as a vector graphic: the server writes it to a temp .svg under $TMPDIR/indesign-mcp/svg/ (or to saveTo) and places it like place_graphic. Use it as the fallback when a native path cannot express something (text as outlines, complex filters). LIMITS: the result is NOT editable as paths in InDesign, and the RGB colours of the SVG are converted to CMYK only at output, through the CMYK profile of the document (accurate, but you cannot dictate exact CMYK numbers). For exact CMYK values use create_path_from_svg with swatches, or create_cmyk_pdf_shape. With embed: true the graphic is embedded and the temp file deleted; otherwise it stays linked to the file (temp files are removed after 3 days). External references, scripts, entities and external url() in the SVG are rejected.

| Parameter | Type | Description |
|---|---|---|
| `svg` *(required)* | string | Complete SVG markup, starting with <svg ...> |
| `x` | number | X position in mm Default: `10`. |
| `y` | number | Y position in mm Default: `10`. |
| `width` | number | Frame width in mm |
| `height` | number | Frame height in mm |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fitOption` | `PROPORTIONALLY` \| `FILL_PROPORTIONALLY` \| `CONTENT_TO_FRAME` \| `FRAME_TO_CONTENT` \| `CENTER_CONTENT` \| `NONE` | How the graphic fits the frame (with width and height): PROPORTIONALLY = fit inside, FILL_PROPORTIONALLY = crop to fill Default: `"PROPORTIONALLY"`. |
| `embed` | boolean | Unlink and embed the graphic in the document Default: `false`. |
| `saveTo` | string | Optional: keep the .svg at this path (inside the allowed folders) instead of a temp file |

#### `create_cmyk_pdf_shape`

Generate a small vector PDF with EXACT CMYK fill/stroke values from SVG path data and place it (like place_graphic). Use it for logos or multi-colour shapes when you need exact brand CMYK numbers (the SVG route converts RGB through the profile) and native paths (create_path_from_svg with one swatch per colour) do not fit. The result is a linked or embedded graphic, not an editable path.

| Parameter | Type | Description |
|---|---|---|
| `shapes` *(required)* | object[] | Shapes painted in order: { d: SVG path data, fill: [C,M,Y,K] 0-100, stroke: [C,M,Y,K], strokeWidth: pt in viewBox units } |
| `viewBox` | string | SVG viewBox "minX minY width height" of the path data (default: bounds of all shapes) |
| `x` | number | X position in mm Default: `10`. |
| `y` | number | Y position in mm Default: `10`. |
| `width` | number | Frame width in mm |
| `height` | number | Frame height in mm |
| `pageIndex` | number | Page index (0-based) Default: `0`. |
| `fitOption` | `PROPORTIONALLY` \| `FILL_PROPORTIONALLY` \| `CONTENT_TO_FRAME` \| `FRAME_TO_CONTENT` \| `CENTER_CONTENT` \| `NONE` | How the graphic fits the frame (with width and height): PROPORTIONALLY = fit inside, FILL_PROPORTIONALLY = crop to fill Default: `"PROPORTIONALLY"`. |
| `embed` | boolean | Unlink and embed the generated PDF (recommended: the temp file is removed afterwards) Default: `true`. |
| `saveTo` | string | Optional: keep the generated .pdf at this path instead of a temp file |

### Colour

#### `create_color_swatch`

Create a colour swatch. Defaults to CMYK (values 0-100). RGB values (0-255) or a hex code are converted to CMYK in print documents with the colour management of InDesign (the CMYK profile of the document) unless keepRgb is true. Presets: rich_black (60/40/40/100), overprint_black (0/0/0/100; overprint is set per object with set_object_overprint).

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Swatch name |
| `colorModel` | `CMYK` \| `RGB` |  Default: `"CMYK"`. |
| `colorValues` | number[] | [C,M,Y,K] 0-100 for CMYK, or [R,G,B] 0-255 for RGB |
| `hex` | string | Alternative to colorValues: RGB hex such as #FF6600 |
| `preset` | `rich_black` \| `overprint_black` | Predefined CMYK values (overrides colorValues) |
| `keepRgb` | boolean | Keep RGB swatches as RGB even in a print document Default: `false`. |
| `spotColor` | boolean | Create as spot colour Default: `false`. |
| `update` | boolean | If a swatch with this name exists, change its values instead of failing (see also update_color_swatch) Default: `false`. |
| `conversion` | `profile` \| `simple` | How RGB / hex values become CMYK: profile = InDesign's colour management with the document's CMYK profile (accurate, default); simple = plain formula without a profile (hues drift toward blue) Default: `"profile"`. |

#### `update_color_swatch`

Change an existing colour swatch in place (all objects using it follow): new values (colorValues CMYK 0-100 or RGB 0-255, or hex), spot/process, and/or a new name. RGB or hex values are converted to CMYK in print documents with the colour management of InDesign unless keepRgb is true. Built-in swatches (Black, Paper, Registration, None) cannot be changed.

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Name of the swatch to change |
| `newName` | string | Rename the swatch |
| `colorModel` | `CMYK` \| `RGB` | Model of colorValues Default: `"CMYK"`. |
| `colorValues` | number[] | [C,M,Y,K] 0-100 or [R,G,B] 0-255 |
| `hex` | string | Alternative to colorValues: RGB hex such as #FF6600 |
| `keepRgb` | boolean | Keep RGB values as RGB even in a print document Default: `false`. |
| `conversion` | `profile` \| `simple` | RGB/hex to CMYK: profile = InDesign colour management (accurate, default); simple = formula without profile Default: `"profile"`. |
| `spotColor` | boolean | true = spot colour, false = process colour |

#### `delete_color_swatch`

Delete a swatch. If objects or styles use it you must say what replaces it (replaceWith: a swatch name, or "none"); the error lists how it is used.

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Name of the swatch to delete |
| `replaceWith` | string | Swatch that takes its place where it is used, or "none" |

#### `list_color_swatches`

List all swatches with kind, model and values (colours), base colour and tint (tints), stops (gradients) and where each is used: number of fills, strokes and styles, plus the ids of up to 10 objects. Local text colour overrides inside stories are not counted.

| Parameter | Type | Description |
|---|---|---|
| `includeUsage` | boolean | Also count where each swatch is used (scans the whole document) Default: `true`. |

#### `apply_color`

Apply color to an object

| Parameter | Type | Description |
|---|---|---|
| `objectId` | number | Stable object id (preferred) |
| `objectIndex` | number | Index in list_page_items for that page (shifts when the stacking order changes) |
| `pageIndex` | number | Page index Default: `0`. |
| `swatchName` *(required)* | string | Color swatch name |
| `property` | `fill` \| `stroke` |  Default: `"fill"`. |

#### `convert_rgb_to_cmyk`

Convert an RGB colour (0-255 values or hex) to CMYK percentages. conversion "profile" (default) uses the colour management of InDesign with the CMYK profile of the active document (accurate: hues stay put, and it tells you when the colour is outside the print gamut and how it will really look). "simple" is a plain formula without a profile, needs no document, and shifts hues (greens toward blue).

| Parameter | Type | Description |
|---|---|---|
| `rgb` | number[] | [R,G,B] 0-255 (sRGB) |
| `hex` | string | Hex colour such as #FF6600 |
| `conversion` | `profile` \| `simple` | profile = InDesign colour management (needs an open document); simple = formula Default: `"profile"`. |

### Tables

#### `create_table`

Create a table in a new text frame (mm; rows include header and footer rows). Returns the frame id, the rows and the columns.

| Parameter | Type | Description |
|---|---|---|
| `x` *(required)* | number | X position in mm |
| `y` *(required)* | number | Y position in mm |
| `width` *(required)* | number | Table width in mm |
| `height` *(required)* | number | Table height in mm |
| `rows` *(required)* | number | Number of rows |
| `columns` *(required)* | number | Number of columns |
| `pageIndex` | number | Page index Default: `0`. |
| `headerRows` | number | Number of header rows Default: `1`. |
| `footerRows` | number | Number of footer rows Default: `0`. |

#### `populate_table`

Populate table with data

| Parameter | Type | Description |
|---|---|---|
| `tableIndex` *(required)* | number | Table index on page |
| `pageIndex` | number | Page index Default: `0`. |
| `data` *(required)* | array[] | Array of arrays with table data |
| `includeHeaders` | boolean | First row contains headers Default: `true`. |

### Layers

#### `create_layer`

Create a new layer with a name, optional guide colour, visibility and lock state. Fails if the name is taken.

| Parameter | Type | Description |
|---|---|---|
| `name` *(required)* | string | Layer name |
| `color` | string | Layer colour name, e.g. RED, BLUE, LIGHT_BLUE, GREEN, YELLOW (error if unknown) |
| `visible` | boolean | Layer visibility Default: `true`. |
| `locked` | boolean | Layer locked state Default: `false`. |

#### `set_active_layer`

Set the active layer

| Parameter | Type | Description |
|---|---|---|
| `layerName` *(required)* | string | Layer name to activate |

#### `list_layers`

List all layers in the document

_No parameters._

### Export and production

#### `export_pdf`

Export the active document as PDF. preset: Print/HighQualityPrint = [High Quality Print], Web/SmallestFileSize = [Smallest File Size], PressQuality = [Press Quality], PDFX1a/PDFX3/PDFX4 = the PDF/X presets, or the exact name of any PDF preset in InDesign. includeBleed uses the document bleed; includeSlug the slug. Overwrites the file, so confirmDestructive is required.

| Parameter | Type | Description |
|---|---|---|
| `filePath` *(required)* | string | Output PDF file path |
| `preset` | string | Preset alias or exact PDF preset name Default: `"HighQualityPrint"`. |
| `pageRange` | string | Pages, e.g. "1", "2-3", "1,3-5" or "all" Default: `"all"`. |
| `includeBleed` | boolean | Use the document bleed Default: `false`. |
| `includeSlug` | boolean | Include the slug area Default: `false`. |
| `confirmDestructive` | boolean | REQUIRED: Confirm file overwrite Default: `false`. |

#### `export_images`

Export pages as PNG or JPEG, one file per page: <folder>/<documentname>_page<N>.<ext>. The folder is created if missing. Returns the list of files written.

| Parameter | Type | Description |
|---|---|---|
| `folderPath` *(required)* | string | Output folder path |
| `format` | `PNG` \| `JPEG` |  Default: `"PNG"`. |
| `resolution` | number | Export resolution in dpi (e.g. 72, 100, 300) Default: `300`. |
| `pageRange` | string | Pages, e.g. "1", "2-3", "1,3-5" or "all" Default: `"all"`. |
| `includeBleed` | boolean | Include the document bleed Default: `false`. |
| `confirmDestructive` | boolean | REQUIRED: Confirm folder write access Default: `false`. |

#### `export_epub`

Export document as EPUB

| Parameter | Type | Description |
|---|---|---|
| `filePath` *(required)* | string | Output EPUB file path |
| `version` | `EPUB2` \| `EPUB3` |  Default: `"EPUB3"`. |
| `imageFormat` | `AUTOMATIC` \| `PNG` \| `JPEG` \| `GIF` | Image conversion Default: `"AUTOMATIC"`. |
| `confirmDestructive` | boolean | REQUIRED: Confirm file overwrite Default: `false`. |

#### `package_document`

Package document for print production

| Parameter | Type | Description |
|---|---|---|
| `folderPath` *(required)* | string | Output folder path |
| `includeLinkedFiles` | boolean | Include linked files Default: `true`. |
| `includeFonts` | boolean | Include fonts Default: `true`. |
| `createReport` | boolean | Create packaging report Default: `true`. |
| `confirmDestructive` | boolean | REQUIRED: Confirm package creation Default: `false`. |

#### `preflight_document`

Run preflight check on the document

| Parameter | Type | Description |
|---|---|---|
| `profile` | string | Preflight profile name |
| `scope` | `document` \| `selection` |  Default: `"document"`. |

#### `data_merge`

Merge the active document (which needs data field placeholders) with a CSV/TXT data source. Writes ONE merged file per format into outputFolder: <document>_merged.pdf and/or <document>_merged.indd

| Parameter | Type | Description |
|---|---|---|
| `dataSourcePath` *(required)* | string | Path to CSV data source |
| `outputFolder` *(required)* | string | Output folder (created if missing) |
| `fileFormat` | `INDD` \| `PDF` \| `BOTH` |  Default: `"PDF"`. |
| `recordRange` | string | Records: "all", a single number ("3") or a range ("1-10") Default: `"all"`. |
| `confirmDestructive` | boolean | REQUIRED: Confirm bulk file creation Default: `false`. |

### Utilities

#### `execute_indesign_code`

⚠️ Execute custom ExtendScript code in InDesign (REQUIRES INDESIGN_ALLOW_ARBITRARY_CODE=1)

| Parameter | Type | Description |
|---|---|---|
| `code` *(required)* | string | ExtendScript/JavaScript code to execute in InDesign. WARNING: Can access filesystem, network, and system APIs! |

<!-- TOOLS:END -->

### Vector shapes

Shapes are native, editable Bezier paths (not pictures). The quickest way for an agent to describe a curve is SVG path data, scaled from a `viewBox` into a target rectangle in mm:

```javascript
// A green band with a gently curved (S-shaped) top edge across the bottom of an A5 page with 3 mm bleed
create_color_swatch({ name: "Green", colorValues: [85, 10, 100, 20] })
create_path_from_svg({
  d: "M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z",
  viewBox: "0 0 216 50",          // the coordinate system the path data was written in
  x: -3, y: 55, width: 216, height: 48,   // where it goes: mm, negative values reach into the bleed
  fill: "Green", opacity: 85
})
// -> "Path created from SVG data: Polygon id=223 ... points=5 paths=1 closed=true opacity=85"
```

- `create_path` takes anchor points with optional handles (`leftDirection`, `rightDirection`, `pointType`); `edit_path_points` reads and edits them (add, move, delete, set_type, reverse, set_closed), so paths round-trip exactly.
- `create_line` (curves, dashes, arrowheads), `create_polygon` (regular polygons, stars, rounded corners), `convert_shape`, `set_corner_options`, `pathfinder`, `set_object_fill`, `set_object_stroke`, `create_gradient_swatch`, `set_gradient_feather`, `place_image_in_shape`, `duplicate_object`, `align_objects`, `distribute_objects` and `get_object_info` complete the set.
- `set_object_fill` / `set_object_stroke` change an object in place, so its stacking order is kept (no delete and recreate).
- **Colour:** paths use swatches, so CMYK values are exact. **Fallbacks** when a native path cannot express something: `create_cmyk_pdf_shape` (exact CMYK, not editable) and `create_graphic_from_svg` (SVG string; not editable as paths; its RGB colours are converted to CMYK at output through the document profile).
- **Pathfinder semantics:** `union`, `intersect` and `exclude_overlap` combine all shapes; `subtract` keeps the first id and cuts the others out of it; `minus_back` keeps the frontmost shape and cuts the shapes behind it out of it.

### Seeing and measuring

`render_preview` returns a page (or a single object) as an image directly in the tool result, and `view_document` includes a picture of the current page. `measure_text` gives the line count, the text and width (mm) of every line and the overflow state without exporting anything. Forced line breaks (`<br>` or `lineBreak: "forced"`) are real Shift+Enter breaks; `get_text_content` shows them as an arrow.

## Known limits

- **macOS only** (AppleScript is used to start scripts in InDesign). The application name defaults to `Adobe InDesign 2026`; set `INDESIGN_APP_NAME` for other versions.
- One script runs at a time; calls are queued. A modal dialog open in InDesign blocks all tools until it is closed (calls time out after 60 s).
- Fonts must be installed: unknown fonts are an error (`Font not installed: ...`), not silently ignored.
- Tool scripts run with InDesign's user interaction set to "never interact" (restored afterwards), so an invalid request produces a tool error instead of a modal dialog that would block InDesign.
- Tests and tools act on the *active* document: do not open, close or switch documents in InDesign while a tool run or `npm test` is in progress.
- `place_graphic` handles `.svg`, `.pdf`, `.ai` and `.eps`; only SVG and PDF are covered by the tests.
- Vector previews are limited to 4000 px on the longest side.
- `set_object_overprint` fails for objects with transparency or a blend mode (an InDesign restriction).
- Bright RGB blues, oranges and some greens are outside the CMYK gamut: the conversion returns the nearest printable colour (`convert_rgb_to_cmyk` reports the difference).
- `insert_markdown_text` supports headers, bold, italic and paragraphs only.
- Objects are addressed by id; `pageIndex`/`index` addressing shifts when the stacking order changes.
- - `data_merge` needs a document that already contains data field placeholders; it writes one merged file per format (`<document>_merged.pdf` / `.indd`), not one per record.
- `export_epub` uses InDesign's default EPUB settings apart from `version` and `imageFormat`; `package_document` needs a saved document and always ignores preflight errors.

## How scripts run

Each tool generates ExtendScript, writes it to a temp `.jsx` file and runs it with `do script (POSIX file "...") language javascript`. The AppleScript is a fixed three-line wrapper; tool arguments never become part of AppleScript source. Arguments are validated against the tool's schema before use, and embedded in scripts only as JSON string literals (`jsStr`), checked numbers (`jsNum`) or checked identifiers (`jsEnum`). The script's last expression is captured with `eval`, so multi-line final expressions work.

Errors are reported by number (`Code: 1234`) and never depend on InDesign's localised message text. Validation errors of the server itself (`No document open`, `Swatch not found: X`, `Paragraph style not found: X`, ...) are English and the same on every system.

### ExtendScript gotchas (for contributors)
- Never write `a === b ? x : c === d ? y : z`: InDesign's ExtendScript mis-parses chained ternaries with `===` (it returned the wrong branch). Use `if / else if`.
- Compare enums with `String(value).indexOf("NAME") === 0` rather than `===` against the enum constant inside larger expressions.
- Reserved words such as `char`, `int`, `long` cannot be variable names ("Illegal use of reserved word").
- Look objects up with `doc.pageItems.itemByID(id).getElements()[0]`; the plain result is a generic `PageItem` without type-specific properties.
- A PDF preset overrides `app.pdfExportPreferences` (bleed, slug); only `pageRange` is taken from the preferences.

### Debugging
Scripts live in `$TMPDIR/indesign-mcp/`. A successful run deletes its script; a failed run **keeps** it and the error message ends with `[debug script: <path>]`. Kept scripts older than three days are removed automatically. Open the `.jsx` in the ExtendScript debugger, or run it with `osascript -e 'tell application "Adobe InDesign 2026" to do script (POSIX file "<path>") language javascript'`.

## Testing

```bash
npm test
```

`tests/run-tests.mjs` starts InDesign if it is not running, drives the server over MCP and asserts the results against scratch documents: A5 landscape flyer with styles, frames, image placement, rotate, delete, PDF export (Print, Web, HighQualityPrint, PressQuality, page range, bleed, slug) and PNG/JPEG export at 72/100/300 dpi, plus error handling. It only closes documents it created itself. `npm run docs` regenerates the tool reference above.

## 💡 Use Cases

### **Automated Publishing**
- Generate newsletters, brochures, and reports from data
- Batch process multiple documents
- Consistent styling across document series

### **Data-Driven Documents**
- Mail merge for personalized materials
- Catalog generation from databases
- Financial reports with dynamic content

### **Professional Workflows**
- Template-based document creation
- Brand compliance automation
- Print production preparation

### **Educational Materials**
- Automated textbook layout
- Exercise sheet generation
- Multi-language document variants

## 🔧 Advanced Configuration

### Custom ExtendScript Integration
`execute_indesign_code` is **disabled by default**. It only runs when the server is started with `INDESIGN_ALLOW_ARBITRARY_CODE=1` (add it to the `env` block of your MCP client configuration). Enable it only if you trust everything that can call the server.
```javascript
execute_indesign_code({
  code: `
    // Custom InDesign scripting
    var doc = app.activeDocument;
    // Your custom automation logic here
  `
})
```

### Batch Processing Example
```javascript
// Process multiple files
const files = ["doc1.indd", "doc2.indd", "doc3.indd"];
for (const file of files) {
  await open_document({ filePath: file });
  await export_pdf({ 
    filePath: file.replace('.indd', '.pdf'),
    preset: 'HighQualityPrint',
    confirmDestructive: true
  });
  await close_document({ save: false, confirmDestructive: true });
}
```

## 🐛 Troubleshooting

### Common Issues

**"Adobe InDesign not found"**
- Ensure InDesign is installed and running, and that `INDESIGN_APP_NAME` matches its name (default `Adobe InDesign 2026`)
- Allow your MCP client to control InDesign in System Settings > Privacy & Security > Automation

**"Script execution failed"**
- Read the error: it carries the InDesign error code and, for unexpected failures, the path of the kept script (`[debug script: ...]`, see "Debugging")
- Make sure no dialog is open in InDesign

**"Tool not found"**
- Restart MCP client after configuration changes
- Verify server is running with `node index.js`

### Debug Mode
```bash
node --inspect index.js
```

## 🤝 Contributing

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **Push** to the branch (`git push origin feature/AmazingFeature`)
5. **Open** a Pull Request

## 📝 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Anthropic** for the Model Context Protocol
- **Adobe** for InDesign ExtendScript API
- **MCP Community** for tools and inspiration

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/lucdesign/indesign-mcp-server/issues)
- **Discussions**: [GitHub Discussions](https://github.com/lucdesign/indesign-mcp-server/discussions)

---

**Made with ❤️ for the publishing community**

*Transform your InDesign workflows with AI automation*