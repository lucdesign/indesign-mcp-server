# InDesign MCP Server

A comprehensive **Model Context Protocol (MCP) server** for **Adobe InDesign automation** with **35+ professional tools**. This server enables AI assistants like Claude to directly control Adobe InDesign, automating complex publishing workflows, document creation, and professional layout tasks.

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

- **Adobe InDesign 2025** (or compatible version)
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
  fontSize: 18,
  fontFamily: "Helvetica Neue",
  fontStyle: "Bold",
  alignment: "LEFT_ALIGN"
})
```

### Create and Apply Styles
```javascript
// Create a character style
create_character_style({
  name: "Emphasis",
  fontFamily: "Helvetica Neue",
  fontStyle: "Italic",
  textColor: "Red"
})

// Apply to specific text
find_replace_text({
  findText: "important",
  replaceText: "important",
  applyCharacterStyle: "Emphasis"
})
```

### Export Professional PDF
```javascript
export_pdf({
  filePath: "/path/to/output.pdf",
  preset: "HighQualityPrint",
  includeBleed: true,
  colorProfile: "ISO Coated v2"
})
```

## 🛠️ Available Tools

### **Document Management (5 tools)**
- `get_document_info` - Document information and statistics
- `create_document` - Advanced document creation
- `open_document` - Open existing documents
- `save_document` - Save with optional new path
- `close_document` - Close with save options

### **Page Management (4 tools)**
- `add_page` - Add pages with position control
- `delete_page` - Remove pages safely
- `duplicate_page` - Copy pages with content
- `navigate_to_page` - Page navigation

### **Text Management (3 tools)**
- `create_text_frame` - Advanced text frame creation
- `edit_text_frame` - Modify existing frames
- `find_replace_text` - Search and replace with GREP

### **Graphics Management (3 tools)**
- `place_image` - Image placement with fit options
- `create_rectangle` - Rectangle creation with styling
- `create_ellipse` - Ellipse creation with styling

### **Style Management (4 tools)**
- `create_paragraph_style` - Paragraph style creation
- `create_character_style` - Character style creation
- `apply_paragraph_style` - Apply styles to text
- `list_styles` - List all available styles

### **Color Management (3 tools)**
- `create_color_swatch` - Create CMYK/RGB/Spot colors
- `list_color_swatches` - List available colors
- `apply_color` - Apply colors to objects

### **Table Management (2 tools)**
- `create_table` - Table creation with headers/footers
- `populate_table` - Fill tables with data

### **Layer Management (3 tools)**
- `create_layer` - Create new layers
- `set_active_layer` - Switch active layer
- `list_layers` - List all layers

### **Export & Production (4 tools)**
- `export_pdf` - Professional PDF export
- `export_images` - Multi-format image export
- `export_epub` - Digital publishing export
- `package_document` - Print production packaging

### **Automation & Utilities (4 tools)**
- `execute_indesign_code` - Custom ExtendScript execution
- `preflight_document` - Quality checking
- `view_document` - Document visualization
- `data_merge` - Automated data integration

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
    preset: 'HighQualityPrint' 
  });
  await close_document({ save: false });
}
```

## 🐛 Troubleshooting

### Common Issues

**"Adobe InDesign not found"**
- Ensure InDesign 2025 is installed and running
- Check AppleScript permissions in System Preferences

**"Script execution failed"**
- Verify Adobe InDesign is the active application
- Check for document-specific errors in console

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