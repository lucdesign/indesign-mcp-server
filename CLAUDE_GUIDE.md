# InDesign MCP Server - Usage Guide for AI Assistants

How to work with this server reliably. The complete tool and parameter reference is in `README.md`.

## Rules of thumb

1. **Use ids, not indices.** Every `create_*` tool and `place_image` returns `id=<number>`. Pass it as `frameId` (text tools) or `id` (object tools). Indices (`frameIndex`, `index`) count from the *frontmost* object and shift whenever objects are added, deleted or reordered.
2. **Discover before you act.** `list_open_documents`, `get_document_info`, `list_page_items`, `list_text_frames`, `list_styles`, `list_color_swatches`, `get_selected_objects`.
3. **Create styles first, then use them.** Formatting belongs in paragraph/character styles. Pass `paragraphStyle` to `create_text_frame`; pass only the formatting parameters you really want as local overrides.
4. **Read the result.** Text tools report `overflows=true` when the text does not fit; fix it by resizing (`set_object_geometry`) or shortening the text.
5. **Everything is millimetres**, origin at the top-left of each page's trim area, negative values reach into the bleed. Angles are degrees, counter-clockwise positive.

## Typical workflow

```javascript
create_document({ preset: "A5", orientation: "Landscape", pages: 2, bleed: 3 })
create_paragraph_style({ name: "Title", fontFamily: "Helvetica Neue", fontStyle: "Bold", fontSize: 24, leading: 28 })
create_paragraph_style({ name: "Body", fontFamily: "Helvetica Neue", fontStyle: "Regular", fontSize: 10, leading: 13 })
create_character_style({ name: "Bold", fontFamily: "Helvetica Neue", fontStyle: "Bold" })

create_text_frame({ content: "Headline", paragraphStyle: "Title", x: 10, y: 10, width: 120, height: 20 })
// -> "Text frame created: id=250 ... overflows=false"
create_text_frame({ content: "Um 7.30 Uhr beginnt es", paragraphStyle: "Body", x: 10, y: 40, width: 120, height: 30 })
apply_character_style_to_range({ frameId: 251, styleName: "Bold", matchText: "7.30 Uhr" })

place_image({ imagePath: "/Users/me/photo.jpg", x: 140, y: 10, width: 60, height: 60, fitOption: "FILL_PROPORTIONALLY" })
rotate_object({ id: 252, angle: 5 })

export_pdf({ filePath: "/Users/me/flyer.pdf", preset: "HighQualityPrint", includeBleed: true, confirmDestructive: true })
```

## Vector shapes and previews

- Draw curves with `create_path_from_svg` (SVG path data + viewBox scaled into an x/y/width/height rectangle in mm) or `create_path` (anchor points with handles). Both return an id; `edit_path_points` reads and changes the points, `get_object_info` shows everything about an object.
- Change colours in place with `set_object_fill` / `set_object_stroke` (keeps the stacking order); use swatches, so CMYK is exact.
- Check your work: `render_preview` returns the page (or one object) as an image, `measure_text` reports line count/widths/overflow. Do this after layout changes instead of exporting files.
- Logos: try native paths first; `place_graphic` for .svg/.pdf/.ai/.eps files; `create_graphic_from_svg` only as a fallback (RGB converted, not editable).
- Do not open or switch documents in InDesign while tools run.

## Text

- `\n` in `content` is a **paragraph break**; `lineBreak: "forced"` or `<br>` gives a forced line break.
- `edit_text_frame` with `content` replaces the whole story.
- `insert_markdown_text` maps Markdown to existing styles (`Heading 1`-`6`, `Bold`, `Italic`, `Bold Italic`; override with `styleMap`). It fails, changing nothing, if a style is missing - create the styles first.
- `fix_typography_in_selection` and `clean_imported_text` use GREP find/change, so character styles and formatting survive.
- With the user's selection: `get_selected_objects` shows what is selected; `useSelectedFrame: true` works with the typography and Markdown tools.

## Errors and what to do

| Message | Meaning / next step |
|---|---|
| `No document open` | Open or create a document first. |
| `No object with id N` | The object was deleted; call `list_page_items`. |
| `Invalid index: N` | Index out of range; prefer ids. |
| `Paragraph style not found: X` / `Character style ...` / `Object style ...` / `Swatch not found: X` | Create it first (`create_paragraph_style`, `create_color_swatch`, ...) or use `list_styles` / `list_color_swatches`. |
| `Font not installed: F (S)` | Check the exact family and style names. |
| `Missing styles - ...` | `insert_markdown_text`: create the listed styles or pass `styleMap`. |
| `save_document`: `has never been saved: pass filePath` / `already exists ... confirmDestructive` / `already open in the document` | A new document needs a `filePath`; overwriting a different existing file needs `confirmDestructive: true`; a file that is open in another document cannot be a Save As target. |
| `Security confirmation required` | Destructive tools (export, save-as, delete page, close with changes, data merge, package) need `confirmDestructive: true`; only add it when the user asked for that action. |
| `... [debug script: path]` | Unexpected InDesign failure; the generated script was kept at that path. |

`execute_indesign_code` is disabled unless the server runs with `INDESIGN_ALLOW_ARBITRARY_CODE=1`. Do not ask users to enable it; the dedicated tools cover normal work.
