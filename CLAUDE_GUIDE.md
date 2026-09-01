# InDesign MCP Server - Claude Usage Guide

## Enhanced Features for Better Claude Integration

### 🎯 Problem Solved: Selected Text Frames & Markdown Formatting

This enhanced MCP server now provides tools that Claude needs to work efficiently with InDesign:

## Essential Workflow for Claude

### 1. Finding Text Frames
**ALWAYS start with discovery tools:**

```javascript
// Option A: List all text frames on a page
await list_text_frames({ pageIndex: 0 })

// Option B: Get info about selected objects
await get_selected_objects()
```

### 2. Working with Selected Frames
When user manually selects a text frame in InDesign:

```javascript
// Check what's selected
await get_selected_objects()
// This will show: "Text Frame - Content: Hello World (Frame Index: 2)"

// Then use the Frame Index with other functions
await edit_text_frame({ frameIndex: 2, content: "New content" })
```

### 3. Markdown Text Insertion
**NEW FEATURE:** Insert markdown text with automatic formatting:

```javascript
// Insert markdown into selected frame
await insert_markdown_text({
  markdownText: `# Header\n**Bold text** and *italic text*\nRegular paragraph`,
  useSelectedFrame: true
})

// Or insert into specific frame
await insert_markdown_text({
  markdownText: "# My Header\n\nSome content",
  frameIndex: 0,
  pageIndex: 0
})
```

## Markdown Features Supported

- **Headers:** `# Header 1`, `## Header 2`, etc.
- **Bold:** `**bold text**`
- **Italic:** `*italic text*`
- **Automatic style application** using existing InDesign styles

### Style Mapping
The system looks for these InDesign styles:
- **Headers:** "Header 1", "Heading 1", "H1", "Header1"
- **Bold:** "Bold", "Strong"
- **Italic:** "Italic", "Emphasis"

## Error Messages Guide for Claude

### Common Errors and Solutions:

**"Invalid text frame index: 5"**
→ Use `list_text_frames()` first to see available frames

**"No text frame selected"**
→ User needs to select a frame in InDesign, then use `get_selected_objects()`

**"No objects selected"**
→ Tell user to select a text frame in InDesign first

## Best Practices for Claude

### 1. Always Check Before Acting
```javascript
// Don't assume - always check first
const frames = await list_text_frames({ pageIndex: 0 });
// Then use the correct frameIndex
```

### 2. Use Selected Objects When Available
```javascript
// When user says "put text in this frame"
const selected = await get_selected_objects();
// Check if it shows a text frame with Frame Index
```

### 3. Provide Clear Instructions
When errors occur, tell user:
- "Please select the text frame in InDesign first"
- "I found 3 text frames on this page - which one do you want?"

## Function Reference

### Discovery Functions
- `get_selected_objects()` - See what user has selected
- `list_text_frames(pageIndex)` - List all text frames on page
- `list_styles()` - See available paragraph/character styles

### Text Functions
- `insert_markdown_text()` - **NEW:** Insert formatted markdown
- `create_text_frame()` - Create new text frame
- `edit_text_frame()` - Edit existing frame (requires frameIndex)

### Style Functions
- `apply_paragraph_style()` - Apply existing paragraph style
- `create_paragraph_style()` - Create new paragraph style

## Example Claude Workflows

### Workflow 1: User Selected Frame
```
User: "Put this markdown text in the selected frame"
Claude: 
1. get_selected_objects()
2. If text frame found → insert_markdown_text({ useSelectedFrame: true })
3. If no frame selected → Ask user to select frame first
```

### Workflow 2: Specific Frame
```
User: "Put text in the first frame on page 2"
Claude:
1. list_text_frames({ pageIndex: 1 })
2. insert_markdown_text({ frameIndex: 0, pageIndex: 1 })
```

### Workflow 3: No Frames Exist
```
User: "Add this text to the page"
Claude:
1. list_text_frames({ pageIndex: 0 })
2. If no frames → create_text_frame() first
3. Then insert_markdown_text()
```

## Pro Tips for Claude

1. **Frame Index is Zero-Based:** First frame = 0, second = 1, etc.
2. **Page Index is Zero-Based:** First page = 0, second = 1, etc.
3. **Always Use Discovery First:** Don't guess frame indices
4. **Selected Objects Have Priority:** If user selected something, use it
5. **Error Messages Are Helpful:** They tell you exactly what to do next

## Troubleshooting

**Claude keeps using wrong frameIndex:**
→ Make sure to call `list_text_frames()` or `get_selected_objects()` first

**Markdown formatting not working:**
→ Check if document has appropriate paragraph/character styles

**"No document open" errors:**
→ User needs to have InDesign document open first
