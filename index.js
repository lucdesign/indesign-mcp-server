#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { execSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

class InDesignMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'indesign-server-complete',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupToolHandlers();
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        // =================== DOCUMENT MANAGEMENT ===================
        {
          name: 'get_document_info',
          description: 'Get detailed information about the current InDesign document',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'create_document',
          description: 'Create a new InDesign document with advanced options',
          inputSchema: {
            type: 'object',
            properties: {
              preset: { type: 'string', description: 'Document preset (A4, A5, Letter, Custom, etc.)', default: 'A4' },
              width: { type: 'number', description: 'Document width in mm (for custom preset)' },
              height: { type: 'number', description: 'Document height in mm (for custom preset)' },
              orientation: { type: 'string', enum: ['Portrait', 'Landscape'], default: 'Portrait' },
              pages: { type: 'number', description: 'Number of pages', default: 1 },
              facingPages: { type: 'boolean', description: 'Enable facing pages', default: false },
              bleed: { type: 'number', description: 'Bleed in mm', default: 0 },
              slug: { type: 'number', description: 'Slug area in mm', default: 0 },
              marginTop: { type: 'number', description: 'Top margin in mm', default: 20 },
              marginBottom: { type: 'number', description: 'Bottom margin in mm', default: 20 },
              marginLeft: { type: 'number', description: 'Left margin in mm', default: 20 },
              marginRight: { type: 'number', description: 'Right margin in mm', default: 20 },
            },
          },
        },
        {
          name: 'open_document',
          description: 'Open an existing InDesign document',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Path to the InDesign document (.indd)' },
            },
            required: ['filePath'],
          },
        },
        {
          name: 'save_document',
          description: 'Save the current document',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Optional: Save as new file path' },
            },
          },
        },
        {
          name: 'close_document',
          description: 'Close the current document',
          inputSchema: {
            type: 'object',
            properties: {
              save: { type: 'boolean', description: 'Save before closing', default: false },
            },
          },
        },

        // =================== PAGE MANAGEMENT ===================
        {
          name: 'add_page',
          description: 'Add a new page to the document',
          inputSchema: {
            type: 'object',
            properties: {
              position: { type: 'string', enum: ['before', 'after', 'end'], default: 'end' },
              pageIndex: { type: 'number', description: 'Reference page index (for before/after)' },
              masterPage: { type: 'string', description: 'Master page to apply' },
            },
          },
        },
        {
          name: 'delete_page',
          description: 'Delete a page from the document',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index to delete' },
            },
            required: ['pageIndex'],
          },
        },
        {
          name: 'duplicate_page',
          description: 'Duplicate a page',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index to duplicate' },
              position: { type: 'string', enum: ['before', 'after', 'end'], default: 'after' },
            },
            required: ['pageIndex'],
          },
        },
        {
          name: 'navigate_to_page',
          description: 'Navigate to a specific page',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index to navigate to' },
            },
            required: ['pageIndex'],
          },
        },

        // =================== TEXT MANAGEMENT ===================
        {
          name: 'create_text_frame',
          description: 'Create a text frame with advanced formatting options',
          inputSchema: {
            type: 'object',
            properties: {
              content: { type: 'string', description: 'Text content for the frame' },
              x: { type: 'number', description: 'X position in mm', default: 10 },
              y: { type: 'number', description: 'Y position in mm', default: 10 },
              width: { type: 'number', description: 'Width in mm', default: 100 },
              height: { type: 'number', description: 'Height in mm', default: 50 },
              pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
              fontSize: { type: 'number', description: 'Font size in points', default: 12 },
              fontFamily: { type: 'string', description: 'Font family name', default: 'Helvetica Neue' },
              fontStyle: { type: 'string', description: 'Font style (Regular, Bold, Italic, etc.)', default: 'Regular' },
              textColor: { type: 'string', description: 'Text color (RGB hex or name)', default: 'Black' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'], default: 'LEFT_ALIGN' },
              paragraphStyle: { type: 'string', description: 'Paragraph style name to apply' },
              characterStyle: { type: 'string', description: 'Character style name to apply' },
            },
            required: ['content'],
          },
        },
        {
          name: 'edit_text_frame',
          description: 'Edit properties of an existing text frame',
          inputSchema: {
            type: 'object',
            properties: {
              frameIndex: { type: 'number', description: 'Text frame index on page' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              content: { type: 'string', description: 'New text content' },
              fontSize: { type: 'number', description: 'Font size in points' },
              fontFamily: { type: 'string', description: 'Font family name' },
              textColor: { type: 'string', description: 'Text color' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'] },
            },
            required: ['frameIndex'],
          },
        },
        {
          name: 'find_replace_text',
          description: 'Find and replace text in the document',
          inputSchema: {
            type: 'object',
            properties: {
              findText: { type: 'string', description: 'Text to find' },
              replaceText: { type: 'string', description: 'Replacement text' },
              caseSensitive: { type: 'boolean', description: 'Case sensitive search', default: false },
              wholeWord: { type: 'boolean', description: 'Whole word only', default: false },
              useGrep: { type: 'boolean', description: 'Use GREP (regular expressions)', default: false },
              scope: { type: 'string', enum: ['document', 'story', 'selection'], default: 'document' },
            },
            required: ['findText', 'replaceText'],
          },
        },

        // =================== GRAPHICS MANAGEMENT ===================
        {
          name: 'place_image',
          description: 'Place an image with advanced options',
          inputSchema: {
            type: 'object',
            properties: {
              imagePath: { type: 'string', description: 'Path to the image file' },
              x: { type: 'number', description: 'X position in mm', default: 10 },
              y: { type: 'number', description: 'Y position in mm', default: 10 },
              width: { type: 'number', description: 'Width in mm (optional, maintains aspect ratio if not specified)' },
              height: { type: 'number', description: 'Height in mm (optional, maintains aspect ratio if not specified)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              fitOption: { type: 'string', enum: ['PROPORTIONALLY', 'FRAME_TO_CONTENT', 'CONTENT_TO_FRAME', 'CENTER_CONTENT'], default: 'PROPORTIONALLY' },
              createFrame: { type: 'boolean', description: 'Create frame first', default: true },
            },
            required: ['imagePath'],
          },
        },
        {
          name: 'create_rectangle',
          description: 'Create a rectangle shape',
          inputSchema: {
            type: 'object',
            properties: {
              x: { type: 'number', description: 'X position in mm' },
              y: { type: 'number', description: 'Y position in mm' },
              width: { type: 'number', description: 'Width in mm' },
              height: { type: 'number', description: 'Height in mm' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              fillColor: { type: 'string', description: 'Fill color (RGB hex or swatch name)' },
              strokeColor: { type: 'string', description: 'Stroke color' },
              strokeWidth: { type: 'number', description: 'Stroke width in points', default: 1 },
              cornerRadius: { type: 'number', description: 'Corner radius in mm', default: 0 },
            },
            required: ['x', 'y', 'width', 'height'],
          },
        },
        {
          name: 'create_ellipse',
          description: 'Create an ellipse shape',
          inputSchema: {
            type: 'object',
            properties: {
              x: { type: 'number', description: 'X position in mm' },
              y: { type: 'number', description: 'Y position in mm' },
              width: { type: 'number', description: 'Width in mm' },
              height: { type: 'number', description: 'Height in mm' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              fillColor: { type: 'string', description: 'Fill color' },
              strokeColor: { type: 'string', description: 'Stroke color' },
              strokeWidth: { type: 'number', description: 'Stroke width in points', default: 1 },
            },
            required: ['x', 'y', 'width', 'height'],
          },
        },

        // =================== STYLE MANAGEMENT ===================
        {
          name: 'create_paragraph_style',
          description: 'Create a new paragraph style',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Style name' },
              fontFamily: { type: 'string', description: 'Font family' },
              fontSize: { type: 'number', description: 'Font size in points' },
              leading: { type: 'number', description: 'Leading (line spacing) in points' },
              spaceBefore: { type: 'number', description: 'Space before paragraph in mm' },
              spaceAfter: { type: 'number', description: 'Space after paragraph in mm' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'] },
              textColor: { type: 'string', description: 'Text color' },
              baseStyle: { type: 'string', description: 'Base style to inherit from' },
            },
            required: ['name'],
          },
        },
        {
          name: 'create_character_style',
          description: 'Create a new character style',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Style name' },
              fontFamily: { type: 'string', description: 'Font family' },
              fontStyle: { type: 'string', description: 'Font style (Regular, Bold, Italic)' },
              fontSize: { type: 'number', description: 'Font size in points' },
              textColor: { type: 'string', description: 'Text color' },
              tracking: { type: 'number', description: 'Character tracking' },
              baseStyle: { type: 'string', description: 'Base style to inherit from' },
            },
            required: ['name'],
          },
        },
        {
          name: 'apply_paragraph_style',
          description: 'Apply a paragraph style to text',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Paragraph style name' },
              frameIndex: { type: 'number', description: 'Text frame index' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              startIndex: { type: 'number', description: 'Start character index (optional)' },
              endIndex: { type: 'number', description: 'End character index (optional)' },
            },
            required: ['styleName', 'frameIndex'],
          },
        },
        {
          name: 'list_styles',
          description: 'List all available styles in the document',
          inputSchema: {
            type: 'object',
            properties: {
              styleType: { type: 'string', enum: ['paragraph', 'character', 'object', 'all'], default: 'all' },
            },
          },
        },

        // =================== COLOR MANAGEMENT ===================
        {
          name: 'create_color_swatch',
          description: 'Create a new color swatch',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Swatch name' },
              colorModel: { type: 'string', enum: ['CMYK', 'RGB', 'LAB'], default: 'CMYK' },
              colorValues: { type: 'array', description: 'Color values array [C,M,Y,K] or [R,G,B]', items: { type: 'number' } },
              spotColor: { type: 'boolean', description: 'Create as spot color', default: false },
            },
            required: ['name', 'colorValues'],
          },
        },
        {
          name: 'list_color_swatches',
          description: 'List all color swatches in the document',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'apply_color',
          description: 'Apply color to an object',
          inputSchema: {
            type: 'object',
            properties: {
              objectIndex: { type: 'number', description: 'Object index on page' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              swatchName: { type: 'string', description: 'Color swatch name' },
              property: { type: 'string', enum: ['fill', 'stroke'], default: 'fill' },
            },
            required: ['objectIndex', 'swatchName'],
          },
        },

        // =================== TABLE MANAGEMENT ===================
        {
          name: 'create_table',
          description: 'Create a table',
          inputSchema: {
            type: 'object',
            properties: {
              x: { type: 'number', description: 'X position in mm' },
              y: { type: 'number', description: 'Y position in mm' },
              width: { type: 'number', description: 'Table width in mm' },
              height: { type: 'number', description: 'Table height in mm' },
              rows: { type: 'number', description: 'Number of rows' },
              columns: { type: 'number', description: 'Number of columns' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              headerRows: { type: 'number', description: 'Number of header rows', default: 1 },
              footerRows: { type: 'number', description: 'Number of footer rows', default: 0 },
            },
            required: ['x', 'y', 'width', 'height', 'rows', 'columns'],
          },
        },
        {
          name: 'populate_table',
          description: 'Populate table with data',
          inputSchema: {
            type: 'object',
            properties: {
              tableIndex: { type: 'number', description: 'Table index on page' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              data: { type: 'array', description: 'Array of arrays with table data', items: { type: 'array' } },
              includeHeaders: { type: 'boolean', description: 'First row contains headers', default: true },
            },
            required: ['tableIndex', 'data'],
          },
        },

        // =================== LAYERS MANAGEMENT ===================
        {
          name: 'create_layer',
          description: 'Create a new layer',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Layer name' },
              color: { type: 'string', description: 'Layer color for guides' },
              visible: { type: 'boolean', description: 'Layer visibility', default: true },
              locked: { type: 'boolean', description: 'Layer locked state', default: false },
            },
            required: ['name'],
          },
        },
        {
          name: 'set_active_layer',
          description: 'Set the active layer',
          inputSchema: {
            type: 'object',
            properties: {
              layerName: { type: 'string', description: 'Layer name to activate' },
            },
            required: ['layerName'],
          },
        },
        {
          name: 'list_layers',
          description: 'List all layers in the document',
          inputSchema: { type: 'object', properties: {} },
        },

        // =================== EXPORT & PRINT ===================
        {
          name: 'export_pdf',
          description: 'Export document as PDF with advanced options',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Output PDF file path' },
              preset: { type: 'string', enum: ['Print', 'Web', 'SmallestFileSize', 'HighQualityPrint', 'PressQuality'], default: 'HighQualityPrint' },
              pageRange: { type: 'string', description: 'Page range (e.g., "1-5", "all")', default: 'all' },
              includeBleed: { type: 'boolean', description: 'Include bleed area', default: false },
              includeSlug: { type: 'boolean', description: 'Include slug area', default: false },
              colorProfile: { type: 'string', description: 'Color profile for export' },
              jpegQuality: { type: 'string', enum: ['Low', 'Medium', 'High', 'Maximum'], default: 'High' },
            },
            required: ['filePath'],
          },
        },
        {
          name: 'export_images',
          description: 'Export pages as images',
          inputSchema: {
            type: 'object',
            properties: {
              folderPath: { type: 'string', description: 'Output folder path' },
              format: { type: 'string', enum: ['PNG', 'JPEG', 'TIFF', 'GIF'], default: 'PNG' },
              resolution: { type: 'number', description: 'Export resolution in DPI', default: 300 },
              pageRange: { type: 'string', description: 'Page range', default: 'all' },
              includeBleed: { type: 'boolean', description: 'Include bleed area', default: false },
            },
            required: ['folderPath'],
          },
        },
        {
          name: 'export_epub',
          description: 'Export document as EPUB',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Output EPUB file path' },
              version: { type: 'string', enum: ['EPUB2', 'EPUB3'], default: 'EPUB3' },
              includeImages: { type: 'boolean', description: 'Include images', default: true },
              imageFormat: { type: 'string', enum: ['PNG', 'JPEG', 'GIF'], default: 'PNG' },
            },
            required: ['filePath'],
          },
        },
        {
          name: 'package_document',
          description: 'Package document for print production',
          inputSchema: {
            type: 'object',
            properties: {
              folderPath: { type: 'string', description: 'Output folder path' },
              includeLinkedFiles: { type: 'boolean', description: 'Include linked files', default: true },
              includeFonts: { type: 'boolean', description: 'Include fonts', default: true },
              createReport: { type: 'boolean', description: 'Create packaging report', default: true },
            },
            required: ['folderPath'],
          },
        },

        // =================== UTILITIES & AUTOMATION ===================
        {
          name: 'execute_indesign_code',
          description: 'Execute custom ExtendScript code in InDesign',
          inputSchema: {
            type: 'object',
            properties: {
              code: { type: 'string', description: 'ExtendScript/JavaScript code to execute in InDesign' },
            },
            required: ['code'],
          },
        },
        {
          name: 'preflight_document',
          description: 'Run preflight check on the document',
          inputSchema: {
            type: 'object',
            properties: {
              profile: { type: 'string', description: 'Preflight profile name' },
              scope: { type: 'string', enum: ['document', 'selection'], default: 'document' },
            },
          },
        },
        {
          name: 'view_document',
          description: 'Get visual representation and detailed info about the current document',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'zoom_to_page',
          description: 'Zoom and fit page in view',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index to zoom to' },
              fitOption: { type: 'string', enum: ['FIT_PAGE', 'FIT_SPREAD', 'ACTUAL_SIZE', 'ZOOM_TO_SELECTION'], default: 'FIT_PAGE' },
            },
          },
        },
        {
          name: 'data_merge',
          description: 'Perform data merge operation',
          inputSchema: {
            type: 'object',
            properties: {
              dataSourcePath: { type: 'string', description: 'Path to CSV data source' },
              outputFolder: { type: 'string', description: 'Output folder for merged documents' },
              fileFormat: { type: 'string', enum: ['INDD', 'PDF', 'BOTH'], default: 'PDF' },
              recordRange: { type: 'string', description: 'Record range (e.g., "1-10", "all")', default: 'all' },
            },
            required: ['dataSourcePath', 'outputFolder'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          // Document Management
          case 'get_document_info': return await this.getDocumentInfo();
          case 'create_document': return await this.createDocument(args);
          case 'open_document': return await this.openDocument(args);
          case 'save_document': return await this.saveDocument(args);
          case 'close_document': return await this.closeDocument(args);

          // Page Management
          case 'add_page': return await this.addPage(args);
          case 'delete_page': return await this.deletePage(args);
          case 'duplicate_page': return await this.duplicatePage(args);
          case 'navigate_to_page': return await this.navigateToPage(args);

          // Text Management
          case 'create_text_frame': return await this.createTextFrame(args);
          case 'edit_text_frame': return await this.editTextFrame(args);
          case 'find_replace_text': return await this.findReplaceText(args);

          // Graphics Management
          case 'place_image': return await this.placeImage(args);
          case 'create_rectangle': return await this.createRectangle(args);
          case 'create_ellipse': return await this.createEllipse(args);

          // Style Management
          case 'create_paragraph_style': return await this.createParagraphStyle(args);
          case 'create_character_style': return await this.createCharacterStyle(args);
          case 'apply_paragraph_style': return await this.applyParagraphStyle(args);
          case 'list_styles': return await this.listStyles(args);

          // Color Management
          case 'create_color_swatch': return await this.createColorSwatch(args);
          case 'list_color_swatches': return await this.listColorSwatches();
          case 'apply_color': return await this.applyColor(args);

          // Table Management
          case 'create_table': return await this.createTable(args);
          case 'populate_table': return await this.populateTable(args);

          // Layer Management
          case 'create_layer': return await this.createLayer(args);
          case 'set_active_layer': return await this.setActiveLayer(args);
          case 'list_layers': return await this.listLayers();

          // Export & Print
          case 'export_pdf': return await this.exportPDF(args);
          case 'export_images': return await this.exportImages(args);
          case 'export_epub': return await this.exportEPUB(args);
          case 'package_document': return await this.packageDocument(args);

          // Utilities
          case 'execute_indesign_code': return await this.executeInDesignCode(args.code);
          case 'preflight_document': return await this.preflightDocument(args);
          case 'view_document': return await this.viewDocument();
          case 'zoom_to_page': return await this.zoomToPage(args);
          case 'data_merge': return await this.dataMerge(args);

          default:
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
        }
      } catch (error) {
        throw new McpError(ErrorCode.InternalError, `Error executing tool ${name}: ${error.message}`);
      }
    });
  }

  // =================== CORE UTILITIES ===================
  async executeAppleScript(script) {
    try {
      const result = execSync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`, {
        encoding: 'utf8',
        timeout: 30000,
      });
      return result.trim();
    } catch (error) {
      throw new Error(`AppleScript execution failed: ${error.message}`);
    }
  }

  async executeInDesignScript(script) {
    const tempScript = path.join(__dirname, 'temp_script.jsx');
    
    // Enhanced error handling wrapper
    const wrappedScript = `
      try {
        ${script}
      } catch (error) {
        "ERROR: " + error.message + " (Line: " + (error.line || "unknown") + ")";
      }
    `;
    
    fs.writeFileSync(tempScript, wrappedScript);

    try {
      const appleScript = `
        tell application "Adobe InDesign 2025"
          activate
          do script POSIX file "${tempScript}" language javascript
        end tell
      `;
      
      const result = await this.executeAppleScript(appleScript);
      return result;
    } finally {
      if (fs.existsSync(tempScript)) {
        fs.unlinkSync(tempScript);
      }
    }
  }

  formatResponse(result, operation = "Operation") {
    return {
      content: [
        {
          type: 'text',
          text: `${operation}: ${result}`,
        },
      ],
    };
  }

  // =================== DOCUMENT MANAGEMENT ===================
  async getDocumentInfo() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var info = "=== DOCUMENT INFORMATION ===\\n";
        info += "Name: " + doc.name + "\\n";
        info += "Pages: " + doc.pages.length + "\\n";
        info += "Width: " + doc.documentPreferences.pageWidth + "\\n";
        info += "Height: " + doc.documentPreferences.pageHeight + "\\n";
        info += "Facing Pages: " + doc.documentPreferences.facingPages + "\\n";
        info += "Modified: " + doc.modified + "\\n";
        info += "File Path: " + (doc.fullName ? doc.fullName.fsName : "Unsaved") + "\\n";
        info += "\\n=== MARGINS ===\\n";
        info += "Top: " + doc.marginPreferences.top + "\\n";
        info += "Bottom: " + doc.marginPreferences.bottom + "\\n";
        info += "Left: " + doc.marginPreferences.left + "\\n";
        info += "Right: " + doc.marginPreferences.right + "\\n";
        info += "\\n=== CONTENT SUMMARY ===\\n";
        
        var totalTextFrames = 0;
        var totalImages = 0;
        var totalShapes = 0;
        
        for (var i = 0; i < doc.pages.length; i++) {
          totalTextFrames += doc.pages[i].textFrames.length;
          totalImages += doc.pages[i].rectangles.length; // Approximation
          totalShapes += doc.pages[i].ovals.length + doc.pages[i].polygons.length;
        }
        
        info += "Text Frames: " + totalTextFrames + "\\n";
        info += "Images/Rectangles: " + totalImages + "\\n";
        info += "Shapes: " + totalShapes + "\\n";
        info += "Layers: " + doc.layers.length + "\\n";
        info += "Color Swatches: " + doc.swatches.length;
        
        info;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Document Info");
  }

  async createDocument(args) {
    const {
      preset = 'A4',
      width,
      height,
      orientation = 'Portrait',
      pages = 1,
      facingPages = false,
      bleed = 0,
      slug = 0,
      marginTop = 20,
      marginBottom = 20,
      marginLeft = 20,
      marginRight = 20
    } = args;

    const script = `
      var doc = app.documents.add();
      
      // Set measurement units to millimeters
      doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
      doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
      
      // Set document dimensions
      ${preset === 'Custom' && width && height ? `
        doc.documentPreferences.pageWidth = "${width}mm";
        doc.documentPreferences.pageHeight = "${height}mm";
      ` : `
        // Standard presets
        if ("${preset}" === "A4") {
          doc.documentPreferences.pageWidth = "${orientation === 'Landscape' ? '297mm' : '210mm'}";
          doc.documentPreferences.pageHeight = "${orientation === 'Landscape' ? '210mm' : '297mm'}";
        } else if ("${preset}" === "A5") {
          doc.documentPreferences.pageWidth = "${orientation === 'Landscape' ? '210mm' : '148mm'}";
          doc.documentPreferences.pageHeight = "${orientation === 'Landscape' ? '148mm' : '210mm'}";
        } else if ("${preset}" === "A3") {
          doc.documentPreferences.pageWidth = "${orientation === 'Landscape' ? '420mm' : '297mm'}";
          doc.documentPreferences.pageHeight = "${orientation === 'Landscape' ? '297mm' : '420mm'}";
        } else if ("${preset}" === "Letter") {
          doc.documentPreferences.pageWidth = "${orientation === 'Landscape' ? '279.4mm' : '215.9mm'}";
          doc.documentPreferences.pageHeight = "${orientation === 'Landscape' ? '215.9mm' : '279.4mm'}";
        } else if ("${preset}" === "Legal") {
          doc.documentPreferences.pageWidth = "${orientation === 'Landscape' ? '355.6mm' : '215.9mm'}";
          doc.documentPreferences.pageHeight = "${orientation === 'Landscape' ? '215.9mm' : '355.6mm'}";
        }
      `}
      
      // Document setup
      doc.documentPreferences.facingPages = ${facingPages};
      doc.documentPreferences.pagesPerDocument = ${pages};
      
      // Bleed and slug
      if (${bleed} > 0) {
        doc.documentPreferences.documentBleedTopOffset = "${bleed}mm";
        doc.documentPreferences.documentBleedBottomOffset = "${bleed}mm";
        doc.documentPreferences.documentBleedInsideOrLeftOffset = "${bleed}mm";
        doc.documentPreferences.documentBleedOutsideOrRightOffset = "${bleed}mm";
      }
      
      if (${slug} > 0) {
        doc.documentPreferences.slugTopOffset = "${slug}mm";
        doc.documentPreferences.slugBottomOffset = "${slug}mm";
        doc.documentPreferences.slugInsideOrLeftOffset = "${slug}mm";
        doc.documentPreferences.slugRightOrOutsideOffset = "${slug}mm";
      }
      
      // Margins
      doc.marginPreferences.top = "${marginTop}mm";
      doc.marginPreferences.bottom = "${marginBottom}mm";
      doc.marginPreferences.left = "${marginLeft}mm";
      doc.marginPreferences.right = "${marginRight}mm";
      
      "Document created: " + "${preset}" + " (" + doc.documentPreferences.pageWidth + " x " + doc.documentPreferences.pageHeight + "), " + 
      doc.pages.length + " pages, " + (doc.documentPreferences.facingPages ? "facing pages" : "single pages");
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Document");
  }

  async openDocument(args) {
    const { filePath } = args;
    
    const script = `
      try {
        var file = File("${filePath}");
        if (!file.exists) {
          "File not found: ${filePath}";
        } else {
          var doc = app.open(file);
          "Document opened: " + doc.name + " (" + doc.pages.length + " pages)";
        }
      } catch (e) {
        "Error opening document: " + e.message;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Open Document");
  }

  async saveDocument(args) {
    const { filePath } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open to save";
      } else {
        var doc = app.activeDocument;
        try {
          ${filePath ? `
            var file = File("${filePath}");
            doc.save(file);
            "Document saved as: " + file.fsName;
          ` : `
            if (doc.saved) {
              doc.save();
              "Document saved: " + doc.name;
            } else {
              "Document has never been saved. Please provide a file path.";
            }
          `}
        } catch (e) {
          "Error saving document: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Save Document");
  }

  async closeDocument(args) {
    const { save = false } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open to close";
      } else {
        var doc = app.activeDocument;
        var docName = doc.name;
        try {
          doc.close(${save ? 'SaveOptions.YES' : 'SaveOptions.NO'});
          "Document closed: " + docName;
        } catch (e) {
          "Error closing document: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Close Document");
  }

  // =================== PAGE MANAGEMENT ===================
  async addPage(args) {
    const { position = 'end', pageIndex, masterPage } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var newPage;
          
          ${position === 'end' ? `
            newPage = doc.pages.add();
          ` : `
            var refPage = doc.pages[${pageIndex || 0}];
            newPage = doc.pages.add(${position === 'before' ? 'LocationOptions.BEFORE' : 'LocationOptions.AFTER'}, refPage);
          `}
          
          ${masterPage ? `
            var master = doc.masterSpreads.itemByName("${masterPage}");
            if (master.isValid) {
              newPage.appliedMaster = master;
            }
          ` : ''}
          
          "Page added at position " + (newPage.documentOffset + 1) + ". Total pages: " + doc.pages.length;
        } catch (e) {
          "Error adding page: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Add Page");
  }

  async deletePage(args) {
    const { pageIndex } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          if (${pageIndex} >= doc.pages.length || ${pageIndex} < 0) {
            "Invalid page index: ${pageIndex}. Document has " + doc.pages.length + " pages.";
          } else if (doc.pages.length === 1) {
            "Cannot delete the last page in the document.";
          } else {
            var pageToDelete = doc.pages[${pageIndex}];
            pageToDelete.remove();
            "Page " + (${pageIndex} + 1) + " deleted. Remaining pages: " + doc.pages.length;
          }
        } catch (e) {
          "Error deleting page: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Delete Page");
  }

  async duplicatePage(args) {
    const { pageIndex, position = 'after' } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          if (${pageIndex} >= doc.pages.length || ${pageIndex} < 0) {
            "Invalid page index: ${pageIndex}";
          } else {
            var sourcePage = doc.pages[${pageIndex}];
            var newPage = doc.pages.add(${position === 'before' ? 'LocationOptions.BEFORE' : 'LocationOptions.AFTER'}, sourcePage);
            
            // Copy all page items
            for (var i = 0; i < sourcePage.allPageItems.length; i++) {
              sourcePage.allPageItems[i].duplicate(newPage);
            }
            
            "Page " + (${pageIndex} + 1) + " duplicated. New page position: " + (newPage.documentOffset + 1);
          }
        } catch (e) {
          "Error duplicating page: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Duplicate Page");
  }

  async navigateToPage(args) {
    const { pageIndex } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          if (${pageIndex} >= doc.pages.length || ${pageIndex} < 0) {
            "Invalid page index: ${pageIndex}. Document has " + doc.pages.length + " pages.";
          } else {
            app.activeWindow.activePage = doc.pages[${pageIndex}];
            "Navigated to page " + (${pageIndex} + 1);
          }
        } catch (e) {
          "Error navigating to page: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Navigate to Page");
  }

  // =================== TEXT MANAGEMENT ===================
  async createTextFrame(args) {
    const {
      content,
      x = 10,
      y = 10,
      width = 100,
      height = 50,
      pageIndex = 0,
      fontSize = 12,
      fontFamily = 'Helvetica Neue',
      fontStyle = 'Regular',
      textColor = 'Black',
      alignment = 'LEFT_ALIGN',
      paragraphStyle,
      characterStyle
    } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open. Please create a document first.";
      } else {
        var doc = app.activeDocument;
        try {
          if (${pageIndex} >= doc.pages.length || ${pageIndex} < 0) {
            "Invalid page index: ${pageIndex}";
          } else {
            var page = doc.pages[${pageIndex}];
            
            // Create text frame
            var textFrame = page.textFrames.add();
            textFrame.geometricBounds = ["${y}mm", "${x}mm", "${y + height}mm", "${x + width}mm"];
            
            // Add content
            textFrame.contents = "${content.replace(/"/g, '\\"').replace(/\n/g, '\\n')}";
            
            // Apply formatting
            var story = textFrame.parentStory;
            
            // Font and size
            try {
              story.characters.everyItem().appliedFont = app.fonts.itemByName("${fontFamily}\\t${fontStyle}");
            } catch (e) {
              try {
                story.characters.everyItem().appliedFont = app.fonts.itemByName("${fontFamily}");
              } catch (e2) {
                // Use default font
              }
            }
            
            story.characters.everyItem().pointSize = ${fontSize};
            
            // Color
            try {
              story.characters.everyItem().fillColor = doc.swatches.itemByName("${textColor}");
            } catch (e) {
              // Use default color
            }
            
            // Alignment
            story.paragraphs.everyItem().justification = Justification.${alignment};
            
            // Apply styles if specified
            ${paragraphStyle ? `
              try {
                var pStyle = doc.paragraphStyles.itemByName("${paragraphStyle}");
                if (pStyle.isValid) {
                  story.paragraphs.everyItem().appliedParagraphStyle = pStyle;
                }
              } catch (e) {}
            ` : ''}
            
            ${characterStyle ? `
              try {
                var cStyle = doc.characterStyles.itemByName("${characterStyle}");
                if (cStyle.isValid) {
                  story.characters.everyItem().appliedCharacterStyle = cStyle;
                }
              } catch (e) {}
            ` : ''}
            
            "Text frame created on page " + (${pageIndex} + 1) + " with content: " + "${content.substring(0, 50)}${content.length > 50 ? '...' : ''}";
          }
        } catch (e) {
          "Error creating text frame: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Text Frame");
  }

  async editTextFrame(args) {
    const { frameIndex, pageIndex = 0, content, fontSize, fontFamily, textColor, alignment } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
            "Invalid text frame index: ${frameIndex}. Page has " + page.textFrames.length + " text frames.";
          } else {
            var textFrame = page.textFrames[${frameIndex}];
            var story = textFrame.parentStory;
            
            ${content !== undefined ? `textFrame.contents = "${content.replace(/"/g, '\\"').replace(/\n/g, '\\n')}";` : ''}
            ${fontSize !== undefined ? `story.characters.everyItem().pointSize = ${fontSize};` : ''}
            ${fontFamily !== undefined ? `
              try {
                story.characters.everyItem().appliedFont = app.fonts.itemByName("${fontFamily}");
              } catch (e) {}
            ` : ''}
            ${textColor !== undefined ? `
              try {
                story.characters.everyItem().fillColor = doc.swatches.itemByName("${textColor}");
              } catch (e) {}
            ` : ''}
            ${alignment !== undefined ? `story.paragraphs.everyItem().justification = Justification.${alignment};` : ''}
            
            "Text frame " + ${frameIndex} + " on page " + (${pageIndex} + 1) + " updated successfully";
          }
        } catch (e) {
          "Error editing text frame: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Edit Text Frame");
  }

  async findReplaceText(args) {
    const { findText, replaceText, caseSensitive = false, wholeWord = false, useGrep = false, scope = 'document' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          // Clear previous search settings
          app.findTextPreferences = NothingEnum.nothing;
          app.changeTextPreferences = NothingEnum.nothing;
          
          // Set find preferences
          ${useGrep ? `
            app.findGrepPreferences.findWhat = "${findText.replace(/"/g, '\\"')}";
            app.changeGrepPreferences.changeTo = "${replaceText.replace(/"/g, '\\"')}";
          ` : `
            app.findTextPreferences.findWhat = "${findText.replace(/"/g, '\\"')}";
            app.changeTextPreferences.changeTo = "${replaceText.replace(/"/g, '\\"')}";
            app.findTextPreferences.caseSensitive = ${caseSensitive};
            app.findTextPreferences.wholeWord = ${wholeWord};
          `}
          
          var foundItems;
          var changeCount = 0;
          
          ${scope === 'document' ? `
            foundItems = ${useGrep ? 'doc.findGrep()' : 'doc.findText()'};
            changeCount = ${useGrep ? 'doc.changeGrep()' : 'doc.changeText()'}.length;
          ` : `
            // Handle other scopes (story, selection) if needed
            foundItems = ${useGrep ? 'doc.findGrep()' : 'doc.findText()'};
            changeCount = ${useGrep ? 'doc.changeGrep()' : 'doc.changeText()'}.length;
          `}
          
          // Clear preferences
          app.findTextPreferences = NothingEnum.nothing;
          app.changeTextPreferences = NothingEnum.nothing;
          app.findGrepPreferences = NothingEnum.nothing;
          app.changeGrepPreferences = NothingEnum.nothing;
          
          "Found and replaced " + changeCount + " instances of '" + "${findText}" + "' with '" + "${replaceText}" + "'";
        } catch (e) {
          "Error in find/replace: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Find/Replace Text");
  }

  // =================== GRAPHICS MANAGEMENT ===================
  async placeImage(args) {
    const { imagePath, x = 10, y = 10, width, height, pageIndex = 0, fitOption = 'PROPORTIONALLY', createFrame = true } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open. Please create a document first.";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var imageFile = File("${imagePath}");
          
          if (!imageFile.exists) {
            "Image file not found: ${imagePath}";
          } else {
            ${createFrame ? `
              var rect = page.rectangles.add();
              ${width && height ? `
                rect.geometricBounds = ["${y}mm", "${x}mm", "${y + height}mm", "${x + width}mm"];
              ` : `
                rect.geometricBounds = ["${y}mm", "${x}mm", "${y + 50}mm", "${x + 50}mm"];
              `}
              rect.place(imageFile);
            ` : `
              page.place(imageFile, ["${x}mm", "${y}mm"]);
              var rect = page.rectangles[page.rectangles.length - 1];
            `}
            
            // Apply fit option
            switch ("${fitOption}") {
              case "PROPORTIONALLY":
                rect.fit(FitOptions.PROPORTIONALLY);
                break;
              case "FRAME_TO_CONTENT":
                rect.fit(FitOptions.FRAME_TO_CONTENT);
                break;
              case "CONTENT_TO_FRAME":
                rect.fit(FitOptions.CONTENT_TO_FRAME);
                break;
              case "CENTER_CONTENT":
                rect.fit(FitOptions.CENTER_CONTENT);
                break;
            }
            
            "Image placed: " + imageFile.name + " on page " + (${pageIndex} + 1);
          }
        } catch (e) {
          "Error placing image: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Place Image");
  }

  async createRectangle(args) {
    const { x, y, width, height, pageIndex = 0, fillColor, strokeColor, strokeWidth = 1, cornerRadius = 0 } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var rect = page.rectangles.add();
          
          rect.geometricBounds = ["${y}mm", "${x}mm", "${y + height}mm", "${x + width}mm"];
          
          ${cornerRadius > 0 ? `
            rect.cornerRadius = "${cornerRadius}mm";
          ` : ''}
          
          ${fillColor ? `
            try {
              rect.fillColor = doc.swatches.itemByName("${fillColor}");
            } catch (e) {
              // Try to create color if it doesn't exist
              try {
                var newSwatch = doc.colors.add();
                newSwatch.name = "${fillColor}";
                rect.fillColor = newSwatch;
              } catch (e2) {}
            }
          ` : ''}
          
          ${strokeColor ? `
            try {
              rect.strokeColor = doc.swatches.itemByName("${strokeColor}");
              rect.strokeWeight = "${strokeWidth}pt";
            } catch (e) {}
          ` : ''}
          
          "Rectangle created on page " + (${pageIndex} + 1) + " (${width}mm x ${height}mm)";
        } catch (e) {
          "Error creating rectangle: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Rectangle");
  }

  async createEllipse(args) {
    const { x, y, width, height, pageIndex = 0, fillColor, strokeColor, strokeWidth = 1 } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var ellipse = page.ovals.add();
          
          ellipse.geometricBounds = ["${y}mm", "${x}mm", "${y + height}mm", "${x + width}mm"];
          
          ${fillColor ? `
            try {
              ellipse.fillColor = doc.swatches.itemByName("${fillColor}");
            } catch (e) {}
          ` : ''}
          
          ${strokeColor ? `
            try {
              ellipse.strokeColor = doc.swatches.itemByName("${strokeColor}");
              ellipse.strokeWeight = "${strokeWidth}pt";
            } catch (e) {}
          ` : ''}
          
          "Ellipse created on page " + (${pageIndex} + 1) + " (${width}mm x ${height}mm)";
        } catch (e) {
          "Error creating ellipse: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Ellipse");
  }

  // =================== STYLE MANAGEMENT ===================
  async createParagraphStyle(args) {
    const { name, fontFamily, fontSize, leading, spaceBefore, spaceAfter, alignment, textColor, baseStyle } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var pStyle = doc.paragraphStyles.add();
          pStyle.name = "${name}";
          
          ${baseStyle ? `
            try {
              var base = doc.paragraphStyles.itemByName("${baseStyle}");
              if (base.isValid) {
                pStyle.basedOn = base;
              }
            } catch (e) {}
          ` : ''}
          
          ${fontFamily ? `
            try {
              pStyle.appliedFont = app.fonts.itemByName("${fontFamily}");
            } catch (e) {}
          ` : ''}
          
          ${fontSize ? `pStyle.pointSize = ${fontSize};` : ''}
          ${leading ? `pStyle.leading = ${leading};` : ''}
          ${spaceBefore ? `pStyle.spaceBefore = "${spaceBefore}mm";` : ''}
          ${spaceAfter ? `pStyle.spaceAfter = "${spaceAfter}mm";` : ''}
          ${alignment ? `pStyle.justification = Justification.${alignment};` : ''}
          
          ${textColor ? `
            try {
              pStyle.fillColor = doc.swatches.itemByName("${textColor}");
            } catch (e) {}
          ` : ''}
          
          "Paragraph style '${name}' created successfully";
        } catch (e) {
          "Error creating paragraph style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Paragraph Style");
  }

  async createCharacterStyle(args) {
    const { name, fontFamily, fontStyle, fontSize, textColor, tracking, baseStyle } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var cStyle = doc.characterStyles.add();
          cStyle.name = "${name}";
          
          ${baseStyle ? `
            try {
              var base = doc.characterStyles.itemByName("${baseStyle}");
              if (base.isValid) {
                cStyle.basedOn = base;
              }
            } catch (e) {}
          ` : ''}
          
          ${fontFamily ? `
            try {
              ${fontStyle ? `
                cStyle.appliedFont = app.fonts.itemByName("${fontFamily}\\t${fontStyle}");
              ` : `
                cStyle.appliedFont = app.fonts.itemByName("${fontFamily}");
              `}
            } catch (e) {}
          ` : ''}
          
          ${fontSize ? `cStyle.pointSize = ${fontSize};` : ''}
          ${tracking ? `cStyle.tracking = ${tracking};` : ''}
          
          ${textColor ? `
            try {
              cStyle.fillColor = doc.swatches.itemByName("${textColor}");
            } catch (e) {}
          ` : ''}
          
          "Character style '${name}' created successfully";
        } catch (e) {
          "Error creating character style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Character Style");
  }

  async applyParagraphStyle(args) {
    const { styleName, frameIndex, pageIndex = 0, startIndex, endIndex } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var textFrame = page.textFrames[${frameIndex}];
          var style = doc.paragraphStyles.itemByName("${styleName}");
          
          if (!style.isValid) {
            "Paragraph style '${styleName}' not found";
          } else {
            ${startIndex !== undefined && endIndex !== undefined ? `
              var textRange = textFrame.parentStory.characters.itemByRange(${startIndex}, ${endIndex});
              textRange.paragraphs.everyItem().appliedParagraphStyle = style;
            ` : `
              textFrame.parentStory.paragraphs.everyItem().appliedParagraphStyle = style;
            `}
            
            "Paragraph style '${styleName}' applied to text frame ${frameIndex}";
          }
        } catch (e) {
          "Error applying paragraph style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Apply Paragraph Style");
  }

  async listStyles(args) {
    const { styleType = 'all' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var result = "=== DOCUMENT STYLES ===\\n\\n";
        
        ${styleType === 'all' || styleType === 'paragraph' ? `
          result += "PARAGRAPH STYLES (" + doc.paragraphStyles.length + "):\\n";
          for (var i = 0; i < doc.paragraphStyles.length; i++) {
            result += "  • " + doc.paragraphStyles[i].name + "\\n";
          }
          result += "\\n";
        ` : ''}
        
        ${styleType === 'all' || styleType === 'character' ? `
          result += "CHARACTER STYLES (" + doc.characterStyles.length + "):\\n";
          for (var i = 0; i < doc.characterStyles.length; i++) {
            result += "  • " + doc.characterStyles[i].name + "\\n";
          }
          result += "\\n";
        ` : ''}
        
        ${styleType === 'all' || styleType === 'object' ? `
          result += "OBJECT STYLES (" + doc.objectStyles.length + "):\\n";
          for (var i = 0; i < doc.objectStyles.length; i++) {
            result += "  • " + doc.objectStyles[i].name + "\\n";
          }
        ` : ''}
        
        result;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "List Styles");
  }

  // =================== COLOR MANAGEMENT ===================
  async createColorSwatch(args) {
    const { name, colorModel = 'CMYK', colorValues, spotColor = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var newColor;
          
          if ("${colorModel}" === "CMYK") {
            newColor = doc.colors.add();
            newColor.name = "${name}";
            newColor.model = ColorModel.SPOT;
            newColor.colorValue = [${colorValues.join(', ')}];
            ${spotColor ? `newColor.model = ColorModel.SPOT;` : `newColor.model = ColorModel.PROCESS;`}
          } else if ("${colorModel}" === "RGB") {
            newColor = doc.colors.add();
            newColor.name = "${name}";
            newColor.model = ColorModel.PROCESS;
            newColor.space = ColorSpace.RGB;
            newColor.colorValue = [${colorValues.join(', ')}];
          }
          
          "Color swatch '${name}' created (${colorModel}: ${colorValues.join(', ')})";
        } catch (e) {
          "Error creating color swatch: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Color Swatch");
  }

  async listColorSwatches() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var result = "=== COLOR SWATCHES ===\\n\\n";
        
        result += "TOTAL SWATCHES: " + doc.swatches.length + "\\n\\n";
        
        for (var i = 0; i < doc.swatches.length; i++) {
          var swatch = doc.swatches[i];
          result += "• " + swatch.name;
          
          try {
            if (swatch.color) {
              result += " (" + swatch.color.model + ")";
            }
          } catch (e) {}
          
          result += "\\n";
        }
        
        result;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "List Color Swatches");
  }

  async applyColor(args) {
    const { objectIndex, pageIndex = 0, swatchName, property = 'fill' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var pageItem = page.allPageItems[${objectIndex}];
          var swatch = doc.swatches.itemByName("${swatchName}");
          
          if (!swatch.isValid) {
            "Color swatch '${swatchName}' not found";
          } else {
            if ("${property}" === "fill") {
              pageItem.fillColor = swatch;
            } else if ("${property}" === "stroke") {
              pageItem.strokeColor = swatch;
            }
            
            "Color '${swatchName}' applied to ${property} of object ${objectIndex}";
          }
        } catch (e) {
          "Error applying color: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Apply Color");
  }

  // =================== EXPORT FUNCTIONS ===================
  async exportPDF(args) {
    const { filePath, preset = 'HighQualityPrint', pageRange = 'all', includeBleed = false, includeSlug = false, colorProfile, jpegQuality = 'High' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open. Please create a document first.";
      } else {
        var doc = app.activeDocument;
        try {
          var pdfFile = File("${filePath}");
          var pdfPreset;
          
          // Try to get the specified preset
          try {
            pdfPreset = app.pdfExportPresets.itemByName("[${preset}]");
          } catch (e) {
            pdfPreset = app.pdfExportPresets[0]; // Use first available preset
          }
          
          // Customize export preferences
          ${pageRange !== 'all' ? `
            app.pdfExportPreferences.pageRange = "${pageRange}";
          ` : `
            app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;
          `}
          
          app.pdfExportPreferences.includeBleedMarks = ${includeBleed};
          app.pdfExportPreferences.includeSlugArea = ${includeSlug};
          
          ${colorProfile ? `
            app.pdfExportPreferences.outputIntention = OutputIntention.REPURPOSE;
          ` : ''}
          
          // Set JPEG quality
          if ("${jpegQuality}" === "Low") {
            app.pdfExportPreferences.jpegQuality = JPEGOptionsQuality.LOW;
          } else if ("${jpegQuality}" === "Medium") {
            app.pdfExportPreferences.jpegQuality = JPEGOptionsQuality.MEDIUM;
          } else if ("${jpegQuality}" === "High") {
            app.pdfExportPreferences.jpegQuality = JPEGOptionsQuality.HIGH;
          } else if ("${jpegQuality}" === "Maximum") {
            app.pdfExportPreferences.jpegQuality = JPEGOptionsQuality.MAXIMUM;
          }
          
          doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, pdfPreset);
          "PDF exported successfully to: ${filePath}";
        } catch (e) {
          "Error exporting PDF: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export PDF");
  }

  async exportImages(args) {
    const { folderPath, format = 'PNG', resolution = 300, pageRange = 'all', includeBleed = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var exportFolder = Folder("${folderPath}");
          if (!exportFolder.exists) {
            exportFolder.create();
          }
          
          var exportFormat;
          var fileExtension;
          
          switch ("${format}") {
            case "PNG":
              exportFormat = ExportFormat.PNG_FORMAT;
              fileExtension = ".png";
              app.pngExportPreferences.resolution = ${resolution};
              app.pngExportPreferences.useDocumentBleedWithPDF = ${includeBleed};
              break;
            case "JPEG":
              exportFormat = ExportFormat.JPG;
              fileExtension = ".jpg";
              app.jpegExportPreferences.resolution = ${resolution};
              app.jpegExportPreferences.useDocumentBleedWithPDF = ${includeBleed};
              break;
            default:
              exportFormat = ExportFormat.PNG_FORMAT;
              fileExtension = ".png";
          }
          
          var pages = [];
          ${pageRange === 'all' ? `
            for (var i = 0; i < doc.pages.length; i++) {
              pages.push(doc.pages[i]);
            }
          ` : `
            // Parse page range (simplified)
            var pageNumbers = "${pageRange}".split("-");
            var startPage = parseInt(pageNumbers[0]) - 1;
            var endPage = pageNumbers.length > 1 ? parseInt(pageNumbers[1]) - 1 : startPage;
            
            for (var i = startPage; i <= endPage && i < doc.pages.length; i++) {
              pages.push(doc.pages[i]);
            }
          `}
          
          for (var i = 0; i < pages.length; i++) {
            var page = pages[i];
            var fileName = doc.name.replace(/\.indd$/i, "") + "_page" + (page.documentOffset + 1) + fileExtension;
            var exportFile = File(exportFolder + "/" + fileName);
            
            page.exportFile(exportFormat, exportFile);
          }
          
          "Exported " + pages.length + " pages as ${format} files to: ${folderPath}";
        } catch (e) {
          "Error exporting images: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export Images");
  }

  async exportEPUB(args) {
    const { filePath, version = 'EPUB3', includeImages = true, imageFormat = 'PNG' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var epubFile = File("${filePath}");
          
          // Set EPUB export preferences
          var epubExportPrefs = app.epubExportPreferences;
          epubExportPrefs.epubVersion = ${version === 'EPUB3' ? 'EPubVersion.EPUB_VERSION_3' : 'EPubVersion.EPUB_VERSION_2'};
          epubExportPrefs.preserveLocalOverride = true;
          
          ${includeImages ? `
            epubExportPrefs.imageConversion = ImageConversion.AUTOMATIC;
            if ("${imageFormat}" === "PNG") {
              epubExportPrefs.pngQualityLevel = PNGQualityLevel.HIGH;
            } else if ("${imageFormat}" === "JPEG") {
              epubExportPrefs.jpegOptionsQuality = JPEGOptionsQuality.HIGH;
            }
          ` : `
            epubExportPrefs.imageConversion = ImageConversion.LINK_TO_SERVER;
          `}
          
          doc.exportFile(ExportFormat.EPUB, epubFile);
          "EPUB exported successfully to: ${filePath}";
        } catch (e) {
          "Error exporting EPUB: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export EPUB");
  }

  async packageDocument(args) {
    const { folderPath, includeLinkedFiles = true, includeFonts = true, createReport = true } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var packageFolder = Folder("${folderPath}");
          
          doc.packageForPrint(packageFolder, ${includeLinkedFiles}, ${includeFonts}, true, ${createReport}, "Package created by InDesign MCP Server");
          
          "Document packaged successfully to: ${folderPath}";
        } catch (e) {
          "Error packaging document: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Package Document");
  }

  // =================== UTILITIES ===================
  async executeInDesignCode(code) {
    const result = await this.executeInDesignScript(code);
    return this.formatResponse(result, "Execute Custom Code");
  }

  async viewDocument() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var info = "=== DOCUMENT VIEW ===\\n";
        info += "Document: " + doc.name + "\\n";
        info += "Current Page: " + (app.activeWindow.activePage ? (app.activeWindow.activePage.documentOffset + 1) : "None") + " of " + doc.pages.length + "\\n";
        info += "Zoom Level: " + Math.round(app.activeWindow.zoomPercentage) + "%\\n";
        info += "View: " + app.activeWindow.viewDisplaySetting + "\\n";
        
        try {
          var currentPage = app.activeWindow.activePage || doc.pages[0];
          info += "\\n=== CURRENT PAGE CONTENT ===\\n";
          info += "Text Frames: " + currentPage.textFrames.length + "\\n";
          info += "Images/Rectangles: " + currentPage.rectangles.length + "\\n";
          info += "Ellipses: " + currentPage.ovals.length + "\\n";
          info += "Groups: " + currentPage.groups.length + "\\n";
          info += "Total Objects: " + currentPage.allPageItems.length;
        } catch (e) {
          info += "\\nCould not analyze page content: " + e.message;
        }
        
        info;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Document View");
  }

  // =================== TABLE MANAGEMENT (Simplified implementations) ===================
  async createTable(args) {
    const { x, y, width, height, rows, columns, pageIndex = 0, headerRows = 1, footerRows = 0 } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var textFrame = page.textFrames.add();
          textFrame.geometricBounds = ["${y}mm", "${x}mm", "${y + height}mm", "${x + width}mm"];
          
          var table = textFrame.tables.add();
          table.rowCount = ${rows};
          table.columnCount = ${columns};
          
          ${headerRows > 0 ? `table.headerRowCount = ${headerRows};` : ''}
          ${footerRows > 0 ? `table.footerRowCount = ${footerRows};` : ''}
          
          "Table created with " + ${rows} + " rows and " + ${columns} + " columns on page " + (${pageIndex} + 1);
        } catch (e) {
          "Error creating table: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Table");
  }

  async populateTable(args) {
    const { tableIndex, pageIndex = 0, data, includeHeaders = true } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var tables = [];
          
          // Collect all tables from text frames
          for (var i = 0; i < page.textFrames.length; i++) {
            for (var j = 0; j < page.textFrames[i].tables.length; j++) {
              tables.push(page.textFrames[i].tables[j]);
            }
          }
          
          if (${tableIndex} >= tables.length) {
            "Table index ${tableIndex} not found. Page has " + tables.length + " tables.";
          } else {
            var table = tables[${tableIndex}];
            var tableData = ${JSON.stringify(data)};
            
            for (var row = 0; row < tableData.length && row < table.rowCount; row++) {
              for (var col = 0; col < tableData[row].length && col < table.columnCount; col++) {
                table.cells.item(row * table.columnCount + col).contents = tableData[row][col].toString();
              }
            }
            
            "Table populated with " + tableData.length + " rows of data";
          }
        } catch (e) {
          "Error populating table: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Populate Table");
  }

  // =================== LAYER MANAGEMENT (Simplified implementations) ===================
  async createLayer(args) {
    const { name, color, visible = true, locked = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var layer = doc.layers.add();
          layer.name = "${name}";
          layer.visible = ${visible};
          layer.locked = ${locked};
          
          ${color ? `
            try {
              layer.layerColor = UIColors.${color.toUpperCase()};
            } catch (e) {}
          ` : ''}
          
          "Layer '${name}' created successfully";
        } catch (e) {
          "Error creating layer: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Layer");
  }

  async setActiveLayer(args) {
    const { layerName } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var layer = doc.layers.itemByName("${layerName}");
          if (layer.isValid) {
            doc.activeLayer = layer;
            "Active layer set to: ${layerName}";
          } else {
            "Layer '${layerName}' not found";
          }
        } catch (e) {
          "Error setting active layer: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Set Active Layer");
  }

  async listLayers() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var result = "=== DOCUMENT LAYERS ===\\n\\n";
        
        for (var i = 0; i < doc.layers.length; i++) {
          var layer = doc.layers[i];
          result += "• " + layer.name;
          result += " (Visible: " + layer.visible + ", Locked: " + layer.locked + ")";
          if (layer === doc.activeLayer) {
            result += " [ACTIVE]";
          }
          result += "\\n";
        }
        
        result;
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "List Layers");
  }

  // =================== ADDITIONAL UTILITIES ===================
  async preflightDocument(args) {
    const { profile, scope = 'document' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var preflightProfile;
          
          ${profile ? `
            preflightProfile = app.preflightProfiles.itemByName("${profile}");
            if (!preflightProfile.isValid) {
              preflightProfile = app.preflightProfiles[0];
            }
          ` : `
            preflightProfile = app.preflightProfiles[0];
          `}
          
          var preflightResults = doc.preflightProcesses.add(preflightProfile);
          var errorCount = preflightResults.preflightResultsData.length;
          
          "Preflight check completed. Found " + errorCount + " issues.";
        } catch (e) {
          "Error running preflight: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Preflight Document");
  }

  async zoomToPage(args) {
    const { pageIndex, fitOption = 'FIT_PAGE' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          ${pageIndex !== undefined ? `
            if (${pageIndex} >= 0 && ${pageIndex} < doc.pages.length) {
              app.activeWindow.activePage = doc.pages[${pageIndex}];
            }
          ` : ''}
          
          switch ("${fitOption}") {
            case "FIT_PAGE":
              app.activeWindow.zoom(ZoomOptions.FIT_PAGE);
              break;
            case "FIT_SPREAD":
              app.activeWindow.zoom(ZoomOptions.FIT_SPREAD);
              break;
            case "ACTUAL_SIZE":
              app.activeWindow.zoom(ZoomOptions.ACTUAL_SIZE);
              break;
            default:
              app.activeWindow.zoom(ZoomOptions.FIT_PAGE);
          }
          
          "Zoom applied: ${fitOption}${pageIndex !== undefined ? ` on page ${pageIndex + 1}` : ''}";
        } catch (e) {
          "Error zooming: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Zoom to Page");
  }

  async dataMerge(args) {
    const { dataSourcePath, outputFolder, fileFormat = 'PDF', recordRange = 'all' } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var dataSource = File("${dataSourcePath}");
          if (!dataSource.exists) {
            "Data source file not found: ${dataSourcePath}";
          } else {
            // Set up data merge
            doc.dataMergeProperties.dataMergeSource = dataSource;
            
            var outputDir = Folder("${outputFolder}");
            if (!outputDir.exists) {
              outputDir.create();
            }
            
            // Export merged documents
            ${recordRange === 'all' ? `
              doc.dataMergeProperties.exportRecords(RecordsToMerge.ALL_RECORDS, outputDir, true);
            ` : `
              // Parse record range
              var ranges = "${recordRange}".split("-");
              var startRecord = parseInt(ranges[0]);
              var endRecord = ranges.length > 1 ? parseInt(ranges[1]) : startRecord;
              doc.dataMergeProperties.exportRecords(RecordsToMerge.RANGE, outputDir, true, startRecord, endRecord);
            `}
            
            "Data merge completed. Files saved to: ${outputFolder}";
          }
        } catch (e) {
          "Error in data merge: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Data Merge");
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Complete InDesign MCP server running on stdio');
  }
}

const server = new InDesignMCPServer();
server.run().catch(console.error);