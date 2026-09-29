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
import os from 'os';
import { fileURLToPath } from 'url';
import { jsStr, jsJson, jsNum, jsEnum, appleScriptStr, describeOsascriptError, buildWrapper } from './lib/script-runtime.js';
import { createExtraTools } from './tools/extra-tools.js';
import { createVectorTools } from './tools/vector-tools.js';
import { createGraphicsTools } from './tools/graphics-tools.js';
import { createResourceTools } from './tools/resource-tools.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PACKAGE_VERSION = (() => {
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8')).version; } catch { return '0.0.0'; }
})();

export class InDesignMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'indesign-server-complete',
        version: PACKAGE_VERSION,
      },
      {
        capabilities: {
          tools: { listChanged: true },
        },
      }
    );

    // Security: Define allowed directories for file operations
    this.allowedDirectories = [
      os.homedir(), // User home directory
      '/Users/Shared', // Shared directory
      // Add more as needed via environment variable
      ...(process.env.INDESIGN_ALLOWED_DIRS ? process.env.INDESIGN_ALLOWED_DIRS.split(':') : [])
    ];

    // Tool modules contribute their own definitions and handlers
    const modules = [createExtraTools(this), createVectorTools(this), createGraphicsTools(this), createResourceTools(this)];
    this.extra = {
      definitions: modules.flatMap((m) => m.definitions),
      handlers: Object.assign({}, ...modules.map((m) => m.handlers)),
    };
    this.setupToolHandlers();
  }

  // Security: Path validation to prevent directory traversal
  // Resolve symlinks of the deepest existing ancestor, so a link inside an allowed directory
  // cannot lead outside it. The not-yet-existing tail is appended unchanged.
  realPathOf(target) {
    let existing = path.resolve(target);
    const tail = [];
    while (!fs.existsSync(existing)) {
      const parent = path.dirname(existing);
      if (parent === existing) break;
      tail.unshift(path.basename(existing));
      existing = parent;
    }
    let real;
    try { real = fs.realpathSync(existing); } catch { real = existing; }
    return path.join(real, ...tail);
  }

  // Security: Path validation to prevent directory traversal (also through symlinks)
  validateFilePath(filePath) {
    if (!filePath || typeof filePath !== 'string') {
      throw new Error('Invalid file path provided');
    }

    // Resolve to an absolute path with symlinks resolved; this is the path that is used afterwards
    const resolvedPath = this.realPathOf(filePath);

    // Check if path is within allowed directories (also symlink-resolved)
    const isAllowed = this.allowedDirectories.some(allowedDir => {
      const resolvedAllowedDir = this.realPathOf(allowedDir);
      return resolvedPath.startsWith(resolvedAllowedDir + path.sep) || resolvedPath === resolvedAllowedDir;
    });

    if (!isAllowed) {
      throw new Error(`Access denied: Path '${filePath}' is outside allowed directories. Allowed: ${this.allowedDirectories.join(', ')}`);
    }

    // Prevent access to sensitive system files
    const prohibitedPaths = ['/etc', '/private/etc', '/System', '/usr/bin', '/bin', '/sbin'];
    const isProhibited = prohibitedPaths.some(prohibited =>
      resolvedPath.startsWith(prohibited + path.sep) || resolvedPath === prohibited
    );

    if (isProhibited) {
      throw new Error(`Access denied: Cannot access system directory '${resolvedPath}'`);
    }

    return resolvedPath;
  }

  // Security: User confirmation for destructive operations
  requireUserConfirmation(operation, target, details = '') {
    const warningMessage = `
⚠️  DESTRUCTIVE OPERATION WARNING ⚠️

Operation: ${operation}
Target: ${target}
${details ? `Details: ${details}` : ''}

This operation may:
- Overwrite existing files
- Permanently delete data
- Modify system files

Type 'CONFIRM' to proceed or 'CANCEL' to abort:`;

    // In a real implementation, this would show a dialog or CLI prompt
    // For MCP context, we throw an error requiring explicit confirmation
    throw new McpError(
      ErrorCode.InvalidRequest,
      `Security confirmation required for destructive operation: ${operation} on ${target}. 
      
Add 'confirmDestructive: true' parameter to bypass this safety check.
      
CAUTION: Only do this if you understand the risks and have verified the operation details.`
    );
  }

  // Security: Check if user has explicitly confirmed destructive operation
  validateDestructiveOperation(args, operation, target) {
    if (!args.confirmDestructive) {
      this.requireUserConfirmation(operation, target);
    }
    // User has explicitly confirmed - proceed with operation
  }

  // Check tool arguments against the tool's inputSchema before any script is generated.
  // Numbers (also numeric strings), booleans, strings, arrays and enums are type-checked, so
  // a malformed value can never end up as code in a generated script.
  validateArguments(toolName, args, definitions) {
    const def = definitions.find((d) => d.name === toolName);
    if (!def) return args || {};
    const schema = def.inputSchema || {};
    const input = args && typeof args === 'object' && !Array.isArray(args) ? { ...args } : {};
    for (const key of schema.required || []) {
      if (input[key] === undefined || input[key] === null) {
        throw new McpError(ErrorCode.InvalidParams, `Missing required parameter "${key}" for ${toolName}`);
      }
    }
    const bad = (key, expected, value) => {
      throw new McpError(ErrorCode.InvalidParams, `Invalid parameter "${key}" for ${toolName}: expected ${expected}, got ${JSON.stringify(value)}`);
    };
    const check = (key, spec, value) => {
      if (value === undefined || value === null || !spec) return value;
      switch (spec.type) {
        case 'number': {
          const n = typeof value === 'string' && value.trim() !== '' ? Number(value) : value;
          if (typeof n !== 'number' || !Number.isFinite(n)) bad(key, 'a number', value);
          return n;
        }
        case 'boolean':
          if (typeof value !== 'boolean') bad(key, 'true or false', value);
          return value;
        case 'string':
          if (typeof value !== 'string') bad(key, 'a string', value);
          if (spec.enum && !spec.enum.includes(value)) bad(key, `one of ${spec.enum.join(', ')}`, value);
          return value;
        case 'array':
          if (!Array.isArray(value)) bad(key, 'an array', value);
          return spec.items ? value.map((v, i) => check(`${key}[${i}]`, spec.items, v)) : value;
        case 'object':
          if (typeof value !== 'object' || Array.isArray(value)) bad(key, 'an object', value);
          return value;
        default:
          return value;
      }
    };
    for (const [key, spec] of Object.entries(schema.properties || {})) {
      if (key in input) input[key] = check(key, spec, input[key]);
    }
    return input;
  }

  setupToolHandlers() {
    const buildToolDefinitions = () => [
        ...this.extra.definitions,

        // =================== DOCUMENT MANAGEMENT ===================
        {
          name: 'get_document_info',
          description: 'Get detailed information about the active document: size in mm, pages, facing pages, bleed, slug, margins, colour intent/profiles and swatches',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'create_document',
          description: 'Create a new document (units: mm). Presets: A3, A4, A5, Letter, Legal, Custom. For Custom, width/height are in mm; an explicit orientation swaps them if needed (Landscape => width >= height, Portrait => height >= width); without orientation they are used as given.',
          inputSchema: {
            type: 'object',
            properties: {
              preset: { type: 'string', enum: ['A3', 'A4', 'A5', 'Letter', 'Legal', 'Custom'], description: 'Page size preset', default: 'A4' },
              width: { type: 'number', description: 'Page width in mm (Custom only)' },
              height: { type: 'number', description: 'Page height in mm (Custom only)' },
              orientation: { type: 'string', enum: ['Portrait', 'Landscape'], description: 'Presets default to Portrait. For Custom, only applied when given.' },
              pages: { type: 'number', description: 'Number of pages', default: 1 },
              facingPages: { type: 'boolean', description: 'Enable facing pages', default: false },
              bleed: { type: 'number', description: 'Bleed on all sides in mm', default: 0 },
              slug: { type: 'number', description: 'Slug on all sides in mm', default: 0 },
              marginTop: { type: 'number', description: 'Top margin in mm', default: 20 },
              marginBottom: { type: 'number', description: 'Bottom margin in mm', default: 20 },
              marginLeft: { type: 'number', description: 'Left (inside) margin in mm', default: 20 },
              marginRight: { type: 'number', description: 'Right (outside) margin in mm', default: 20 },
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
          description: 'Save a document (the active one, or the one named by "name"). Without filePath: a plain save (a document that was never saved needs a filePath). With filePath: Save As, the same document is renamed and stays open; the document\'s own file (same path) is just a plain save; overwriting a DIFFERENT existing file needs confirmDestructive. Saving to the file of another open document is refused.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Optional: path of the .indd file to save to (Save As)' },
              name: { type: 'string', description: 'Name of the document as shown by list_open_documents (default: the active document)' },
              confirmDestructive: { type: 'boolean', description: 'Required only to overwrite an existing file that is not the document\'s own file', default: false },
            },
          },
        },
        {
          name: 'close_document',
          description: 'Close a document (the active one, or the one named by "name"). Documents with unsaved changes require save=true or confirmDestructive=true.',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Document name as shown by list_open_documents (default: active document)' },
              save: { type: 'boolean', description: 'Save before closing (only for documents that already have a file)', default: false },
              confirmDestructive: { type: 'boolean', description: 'Required to discard unsaved changes', default: false },
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm page deletion', default: false },
            },
            required: ['pageIndex'],
          },
        },
        {
          name: 'duplicate_page',
          description: 'Duplicate a page with all its content (position: before/after the source page, or end of the document)',
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
          name: 'get_selected_objects',
          description: 'Get information about currently selected objects in InDesign. ESSENTIAL for working with user-selected text frames.',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'get_text_content',
          description: 'Get the text of a text frame (selection, frameId, or frameIndex). Also reports paragraph count and overflow state.',
          inputSchema: {
            type: 'object',
            properties: {
              normalizeSpaces: { type: 'boolean', description: 'Collapse line breaks and repeated spaces', default: true },
              frameId: { type: 'number', description: 'Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex.' },
              frameIndex: { type: 'number', description: 'Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId.' },
              pageIndex: { type: 'number', description: 'Page index (0-based), used with frameIndex', default: 0 },
              maxLength: { type: 'number', description: 'Truncate output to this many characters (0 = no limit)', default: 0 }
            }
          }
        },
        {
          name: 'list_text_frames',
          description: 'List the text frames on a page: stable id, index, geometry (mm), overflow state and a text preview. Ordering: index 0 is the frontmost frame (stacking order), NOT creation order - indices shift when frames are added or reordered, so use the id.',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 }
            }
          }
        },
        {
          name: 'analyze_embedded_objects',
          description: 'Analyze embedded objects (MathML formulas, graphics, etc.) in selected text frame or specified frame',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (preferred over frameIndex)' },
              frameIndex: { type: 'number', description: 'Text frame index (optional if frame is selected)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              maxObjects: { type: 'number', description: 'Maximum number of objects to analyze', default: 5 }
            }
          }
        },
        {
          name: 'insert_markdown_text',
          description: 'Insert Markdown into a text frame, mapped to EXISTING styles. Each non-blank line is one paragraph. Supported: "# " to "###### " headers -> paragraph styles "Heading 1".."Heading 6" (aliases tried: "Header N", "HN", "Ueberschrift N"); **bold** -> character style "Bold" (alias "Strong"); *italic* -> "Italic" (alias "Emphasis"); ***both*** -> both styles. Plain paragraphs keep the frame formatting unless bodyStyle is given. Override any style name with styleMap. Fails with a clear error (and changes nothing) if a required style is missing.',
          inputSchema: {
            type: 'object',
            properties: {
              markdownText: { type: 'string', description: 'Markdown text to insert' },
              frameId: { type: 'number', description: 'Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex.' },
              frameIndex: { type: 'number', description: 'Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId.' },
              pageIndex: { type: 'number', description: 'Page index (0-based), used with frameIndex', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Use the currently selected text frame', default: false },
              replaceContent: { type: 'boolean', description: 'Replace the whole story (true) or append (false)', default: true },
              bodyStyle: { type: 'string', description: 'Paragraph style for normal paragraphs (optional)' },
              styleMap: {
                type: 'object',
                description: 'Override style names: { h1..h6, body, bold, italic } (paragraph styles for h1-h6/body, character styles for bold/italic)',
                additionalProperties: { type: 'string' },
              },
            },
            required: ['markdownText']
          }
        },
        {
          name: 'fix_typography_in_selection',
          description: 'Fix typography in selected text or story. Corrects dates (DD.MM.YYYY with thin spaces), quotes, dashes, and other typographic elements.',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (preferred over frameIndex)' },
              frameIndex: { type: 'number', description: 'Text frame index to fix (use get_selected_objects to work with selection)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Use currently selected text frame', default: false },
              fixDates: { type: 'boolean', description: 'Fix date spacing (DD. MM. YYYY)', default: true },
              fixQuotes: { type: 'boolean', description: 'Fix quotes to typographic quotes', default: true },
              fixDashes: { type: 'boolean', description: 'Fix hyphens to em/en dashes', default: true },
              fixSpaces: { type: 'boolean', description: 'Fix multiple spaces and trailing spaces', default: true }
            }
          }
        },
        {
          name: 'find_typography_issues',
          description: 'Analyze text for common typography issues (wrong spaces in dates, straight quotes, double spaces, etc.)',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (preferred over frameIndex)' },
              frameIndex: { type: 'number', description: 'Text frame index to analyze' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Analyze currently selected text frame', default: false }
            }
          }
        },
        {
          name: 'clean_imported_text',
          description: 'Clean imported text from common typography sins: double paragraph breaks, line breaks instead of paragraphs, trailing spaces, hyphens instead of dashes, manual formatting, bullet lists, hardcoded chapter numbers, etc.',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (preferred over frameIndex)' },
              frameIndex: { type: 'number', description: 'Text frame index to clean' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Clean currently selected text frame', default: false },
              fixParagraphs: { type: 'boolean', description: 'Fix double paragraph breaks and line breaks', default: true },
              fixDashes: { type: 'boolean', description: 'Fix hyphens to proper n-dashes for ranges/thoughts', default: true },
              fixLists: { type: 'boolean', description: 'Remove manual bullet lists and dashes', default: true },
              fixFormatting: { type: 'boolean', description: 'Remove manual bold/italic (prepare for character styles)', default: true },
              fixChapterNumbers: { type: 'boolean', description: 'Remove hardcoded chapter numbers', default: true },
              fixSpaces: { type: 'boolean', description: 'Remove trailing spaces and multiple spaces', default: true }
            }
          }
        },
        {
          name: 'analyze_text_problems',
          description: 'Analyze imported text for common problems before cleaning. Shows what issues exist.',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (preferred over frameIndex)' },
              frameIndex: { type: 'number', description: 'Text frame index to analyze' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Analyze currently selected text frame', default: false }
            }
          }
        },
        {
          name: 'list_grep_searches',
          description: 'List all saved GREP searches in the document (like DATUM search for dates)',
          inputSchema: {
            type: 'object',
            properties: {}
          }
        },
        {
          name: 'create_text_frame',
          description: 'Create a text frame (mm, origin top-left of the page trim). Only parameters you pass are applied as local formatting; defaults (Helvetica Neue 12 pt, Black, left) are used only when no paragraphStyle is given. Returns the frame id and whether the text overflows. In content, "\\n" makes a paragraph break unless lineBreak is "forced"; "<br>" always makes a forced line break (Shift+Enter).',
          inputSchema: {
            type: 'object',
            properties: {
              content: { type: 'string', description: 'Text content. "\\n" = paragraph break (default), "<br>" = forced line break' },
              lineBreak: { type: 'string', enum: ['paragraph', 'forced'], description: 'How "\\n" in content is interpreted', default: 'paragraph' },
              x: { type: 'number', description: 'X position in mm', default: 10 },
              y: { type: 'number', description: 'Y position in mm', default: 10 },
              width: { type: 'number', description: 'Width in mm', default: 100 },
              height: { type: 'number', description: 'Height in mm', default: 50 },
              pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
              fontSize: { type: 'number', description: 'Font size in points (local override)' },
              fontFamily: { type: 'string', description: 'Font family name (local override)' },
              fontStyle: { type: 'string', description: 'Font style (Regular, Bold, Italic, ...) used with fontFamily', default: 'Regular' },
              textColor: { type: 'string', description: 'Swatch name (local override)' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'], description: 'Paragraph alignment (local override)' },
              paragraphStyle: { type: 'string', description: 'Paragraph style name to apply (error if missing)' },
              characterStyle: { type: 'string', description: 'Character style name to apply to all text' },
            },
            required: ['content'],
          },
        },
        {
          name: 'edit_text_frame',
          description: 'Edit a text frame. Setting content replaces the whole story. Returns whether the text overflows. Use frameId (from create_text_frame / list_text_frames) or pageIndex + frameIndex.',
          inputSchema: {
            type: 'object',
            properties: {
              frameId: { type: 'number', description: 'Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex.' },
              frameIndex: { type: 'number', description: 'Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId.' },
              pageIndex: { type: 'number', description: 'Page index (0-based), used with frameIndex', default: 0 },
              content: { type: 'string', description: 'New text (replaces the whole story). "\\n" = paragraph break, "<br>" = forced line break' },
              lineBreak: { type: 'string', enum: ['paragraph', 'forced'], description: 'How "\\n" in content is interpreted', default: 'paragraph' },
              fontSize: { type: 'number', description: 'Font size in points' },
              fontFamily: { type: 'string', description: 'Font family name' },
              fontStyle: { type: 'string', description: 'Font style used with fontFamily', default: 'Regular' },
              textColor: { type: 'string', description: 'Swatch name' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'] },
            },
          },
        },
        {
          name: 'find_replace_text',
          description: 'Find and replace text. scope: document, story (the story of the current selection) or selection (needs a text selection)',
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
          description: 'Place an image into a new frame (mm). Give width and height for a fixed frame, only width or height to keep the aspect ratio, or neither to use the image size. Returns the frame id.',
          inputSchema: {
            type: 'object',
            properties: {
              imagePath: { type: 'string', description: 'Path to the image file' },
              x: { type: 'number', description: 'X position in mm', default: 10 },
              y: { type: 'number', description: 'Y position in mm', default: 10 },
              width: { type: 'number', description: 'Frame width in mm' },
              height: { type: 'number', description: 'Frame height in mm' },
              pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
              fitOption: { type: 'string', enum: ['FILL_PROPORTIONALLY', 'PROPORTIONALLY', 'CONTENT_TO_FRAME', 'FRAME_TO_CONTENT', 'CENTER_CONTENT', 'NONE'], description: 'FILL_PROPORTIONALLY = crop to fill the frame; PROPORTIONALLY = fit inside; NONE = keep 100%', default: 'PROPORTIONALLY' },
              contentOffsetX: { type: 'number', description: 'Move the image inside the frame by this many mm (after fitting)' },
              contentOffsetY: { type: 'number', description: 'Move the image inside the frame by this many mm (after fitting)' },
              contentScale: { type: 'number', description: 'Scale the image content to this percentage (after fitting)' },
              createFrame: { type: 'boolean', description: 'Create a frame first (false = place at x,y with its own size)', default: true },
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
          description: 'Create a paragraph style. Returns the font that was actually applied.',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Style name' },
              baseStyle: { type: 'string', description: 'Base style to inherit from' },
              fontFamily: { type: 'string', description: 'Font family (error if not installed)' },
              fontStyle: { type: 'string', description: 'Font style: Regular, Medium, Bold, Italic, ...' },
              fontSize: { type: 'number', description: 'Font size in points' },
              leading: { type: 'number', description: 'Leading in points' },
              tracking: { type: 'number', description: 'Tracking in 1/1000 em' },
              kerning: { type: 'string', enum: ['Metrics', 'Optical'], description: 'Kerning method' },
              spaceBefore: { type: 'number', description: 'Space before in mm' },
              spaceBeforePt: { type: 'number', description: 'Space before in points (alternative to spaceBefore)' },
              spaceAfter: { type: 'number', description: 'Space after in mm' },
              spaceAfterPt: { type: 'number', description: 'Space after in points (alternative to spaceAfter)' },
              firstLineIndent: { type: 'number', description: 'First-line indent in mm' },
              leftIndent: { type: 'number', description: 'Left indent in mm' },
              rightIndent: { type: 'number', description: 'Right indent in mm' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY', 'LEFT_JUSTIFIED', 'CENTER_JUSTIFIED', 'RIGHT_JUSTIFIED', 'FULLY_JUSTIFIED'] },
              hyphenation: { type: 'boolean', description: 'Enable hyphenation' },
              keepWithNext: { type: 'number', description: 'Keep with next N lines (0 = off)' },
              keepLinesTogether: { type: 'boolean', description: 'Keep lines together' },
              keepFirstLines: { type: 'number', description: 'Keep first N lines together' },
              keepLastLines: { type: 'number', description: 'Keep last N lines together' },
              keepWithPrevious: { type: 'boolean', description: 'Keep with previous paragraph' },
              composer: { type: 'string', enum: ['paragraph', 'single-line'], description: 'Composer: "paragraph" = Adobe Paragraph Composer (default look), "single-line" = Adobe Single-line Composer' },
              capitalization: { type: 'string', enum: ['normal', 'small_caps', 'all_caps', 'cap_to_small_cap', 'lower_case'], description: 'Capitalisation' },
              ligatures: { type: 'boolean', description: 'Standard ligatures on/off' },
              openType: {
                type: 'object',
                description: 'OpenType features: booleans discretionaryLigatures, contextualAlternates, swash, titling, fractions, ordinals, slashedZero, historical; figureStyle (default, tabular_lining, proportional_oldstyle, proportional_lining, tabular_oldstyle); stylisticSets (bit mask number)',
                properties: {
                  discretionaryLigatures: { type: 'boolean' }, contextualAlternates: { type: 'boolean' }, swash: { type: 'boolean' }, titling: { type: 'boolean' },
                  fractions: { type: 'boolean' }, ordinals: { type: 'boolean' }, slashedZero: { type: 'boolean' }, historical: { type: 'boolean' },
                  figureStyle: { type: 'string', enum: ['default', 'tabular_lining', 'proportional_oldstyle', 'proportional_lining', 'tabular_oldstyle'] },
                  stylisticSets: { type: 'number' },
                },
              },
              hyphenateWordsLongerThan: { type: 'number', description: 'Hyphenation: minimum word length' },
              hyphenateAfterFirst: { type: 'number', description: 'Hyphenation: minimum letters before the hyphen' },
              hyphenateBeforeLast: { type: 'number', description: 'Hyphenation: minimum letters after the hyphen' },
              hyphenateLadderLimit: { type: 'number', description: 'Hyphenation: maximum consecutive hyphenated lines' },
              hyphenateCapitalizedWords: { type: 'boolean', description: 'Hyphenation: hyphenate capitalised words' },
              hyphenateLastWord: { type: 'boolean', description: 'Hyphenation: hyphenate the last word of a paragraph' },
              textColor: { type: 'string', description: 'Swatch name' },
            },
            required: ['name'],
          },
        },
        {
          name: 'modify_paragraph_style',
          description: 'Modify a paragraph style (same properties as create_paragraph_style). Returns the font that was actually applied.',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Paragraph style name to modify' },
              baseStyle: { type: 'string', description: 'New base style' },
              fontFamily: { type: 'string', description: 'Font family (error if not installed)' },
              fontStyle: { type: 'string', description: 'Font style: Regular, Medium, Bold, Italic, ...' },
              fontSize: { type: 'number', description: 'Font size in points' },
              leading: { type: 'number', description: 'Leading in points' },
              tracking: { type: 'number', description: 'Tracking in 1/1000 em' },
              kerning: { type: 'string', enum: ['Metrics', 'Optical'], description: 'Kerning method' },
              spaceBefore: { type: 'number', description: 'Space before in mm' },
              spaceBeforePt: { type: 'number', description: 'Space before in points' },
              spaceAfter: { type: 'number', description: 'Space after in mm' },
              spaceAfterPt: { type: 'number', description: 'Space after in points' },
              firstLineIndent: { type: 'number', description: 'First-line indent in mm' },
              leftIndent: { type: 'number', description: 'Left indent in mm' },
              rightIndent: { type: 'number', description: 'Right indent in mm' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY', 'LEFT_JUSTIFIED', 'CENTER_JUSTIFIED', 'RIGHT_JUSTIFIED', 'FULLY_JUSTIFIED'] },
              hyphenation: { type: 'boolean', description: 'Enable hyphenation' },
              keepWithNext: { type: 'number', description: 'Keep with next N lines (0 = off)' },
              keepLinesTogether: { type: 'boolean', description: 'Keep lines together' },
              keepFirstLines: { type: 'number', description: 'Keep first N lines together' },
              keepLastLines: { type: 'number', description: 'Keep last N lines together' },
              keepWithPrevious: { type: 'boolean', description: 'Keep with previous paragraph' },
              composer: { type: 'string', enum: ['paragraph', 'single-line'], description: 'Composer: "paragraph" = Adobe Paragraph Composer (default look), "single-line" = Adobe Single-line Composer' },
              capitalization: { type: 'string', enum: ['normal', 'small_caps', 'all_caps', 'cap_to_small_cap', 'lower_case'], description: 'Capitalisation' },
              ligatures: { type: 'boolean', description: 'Standard ligatures on/off' },
              openType: {
                type: 'object',
                description: 'OpenType features: booleans discretionaryLigatures, contextualAlternates, swash, titling, fractions, ordinals, slashedZero, historical; figureStyle (default, tabular_lining, proportional_oldstyle, proportional_lining, tabular_oldstyle); stylisticSets (bit mask number)',
                properties: {
                  discretionaryLigatures: { type: 'boolean' }, contextualAlternates: { type: 'boolean' }, swash: { type: 'boolean' }, titling: { type: 'boolean' },
                  fractions: { type: 'boolean' }, ordinals: { type: 'boolean' }, slashedZero: { type: 'boolean' }, historical: { type: 'boolean' },
                  figureStyle: { type: 'string', enum: ['default', 'tabular_lining', 'proportional_oldstyle', 'proportional_lining', 'tabular_oldstyle'] },
                  stylisticSets: { type: 'number' },
                },
              },
              hyphenateWordsLongerThan: { type: 'number', description: 'Hyphenation: minimum word length' },
              hyphenateAfterFirst: { type: 'number', description: 'Hyphenation: minimum letters before the hyphen' },
              hyphenateBeforeLast: { type: 'number', description: 'Hyphenation: minimum letters after the hyphen' },
              hyphenateLadderLimit: { type: 'number', description: 'Hyphenation: maximum consecutive hyphenated lines' },
              hyphenateCapitalizedWords: { type: 'boolean', description: 'Hyphenation: hyphenate capitalised words' },
              hyphenateLastWord: { type: 'boolean', description: 'Hyphenation: hyphenate the last word of a paragraph' },
              textColor: { type: 'string', description: 'Swatch name' },
            },
            required: ['styleName'],
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
          name: 'modify_character_style',
          description: 'Modify properties of an existing character style',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Character style name to modify' },
              fontFamily: { type: 'string', description: 'Font family' },
              fontStyle: { type: 'string', description: 'Font style (Regular, Bold, Italic)' },
              fontSize: { type: 'number', description: 'Font size in points' },
              textColor: { type: 'string', description: 'Text color (swatch name)' },
              tracking: { type: 'number', description: 'Character tracking' },
            },
            required: ['styleName'],
          },
        },
        {
          name: 'create_object_style',
          description: 'Create a new object style',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Style name' },
              fillColor: { type: 'string', description: 'Fill color (swatch name)' },
              strokeColor: { type: 'string', description: 'Stroke color (swatch name)' },
              strokeWidth: { type: 'number', description: 'Stroke width in points' },
              transparency: { type: 'number', description: 'Transparency percentage (0-100)' },
              baseStyle: { type: 'string', description: 'Base style to inherit from' },
            },
            required: ['name'],
          },
        },
        {
          name: 'modify_object_style',
          description: 'Modify properties of an existing object style',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Object style name to modify' },
              fillColor: { type: 'string', description: 'Fill color (swatch name)' },
              strokeColor: { type: 'string', description: 'Stroke color (swatch name)' },
              strokeWidth: { type: 'number', description: 'Stroke width in points' },
              transparency: { type: 'number', description: 'Transparency percentage (0-100)' },
            },
            required: ['styleName'],
          },
        },
        {
          name: 'apply_object_style',
          description: 'Apply an object style to an object (objectId, or pageIndex + objectIndex) or, if none is given, to the current selection',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Object style name' },
              objectId: { type: 'number', description: 'Stable object id (preferred)' },
              objectIndex: { type: 'number', description: 'Index in list_page_items for that page (optional if objects are selected)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
            },
            required: ['styleName'],
          },
        },
        {
          name: 'apply_paragraph_style',
          description: 'Apply a paragraph style to a text frame (or a character range of it)',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Paragraph style name' },
              frameId: { type: 'number', description: 'Stable object id of the text frame (returned by create_text_frame / list_text_frames). Preferred over frameIndex.' },
              frameIndex: { type: 'number', description: 'Index in page.textFrames (0 = frontmost). Changes when frames are added or reordered - prefer frameId.' },
              pageIndex: { type: 'number', description: 'Page index (0-based), used with frameIndex', default: 0 },
              startIndex: { type: 'number', description: 'Start character index (optional)' },
              endIndex: { type: 'number', description: 'End character index (optional)' },
            },
            required: ['styleName'],
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
          description: 'Create a colour swatch. Defaults to CMYK (values 0-100). RGB values (0-255) or a hex code are converted to CMYK in print documents with the colour management of InDesign (the CMYK profile of the document) unless keepRgb is true. Presets: rich_black (60/40/40/100), overprint_black (0/0/0/100; overprint is set per object with set_object_overprint).',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Swatch name' },
              colorModel: { type: 'string', enum: ['CMYK', 'RGB'], default: 'CMYK' },
              colorValues: { type: 'array', description: '[C,M,Y,K] 0-100 for CMYK, or [R,G,B] 0-255 for RGB', items: { type: 'number' } },
              hex: { type: 'string', description: 'Alternative to colorValues: RGB hex such as #FF6600' },
              preset: { type: 'string', enum: ['rich_black', 'overprint_black'], description: 'Predefined CMYK values (overrides colorValues)' },
              keepRgb: { type: 'boolean', description: 'Keep RGB swatches as RGB even in a print document', default: false },
              spotColor: { type: 'boolean', description: 'Create as spot colour', default: false },
              update: { type: 'boolean', description: 'If a swatch with this name exists, change its values instead of failing (see also update_color_swatch)', default: false },
              conversion: { type: 'string', enum: ['profile', 'simple'], description: 'How RGB / hex values become CMYK: profile = InDesign\'s colour management with the document\'s CMYK profile (accurate, default); simple = plain formula without a profile (hues drift toward blue)', default: 'profile' },
            },
            required: ['name'],
          },
        },
        {
          name: 'apply_color',
          description: 'Apply color to an object',
          inputSchema: {
            type: 'object',
            properties: {
              objectId: { type: 'number', description: 'Stable object id (preferred)' },
              objectIndex: { type: 'number', description: 'Index in list_page_items for that page (shifts when the stacking order changes)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              swatchName: { type: 'string', description: 'Color swatch name' },
              property: { type: 'string', enum: ['fill', 'stroke'], default: 'fill' },
            },
            required: ['swatchName'],
          },
        },

        // =================== TABLE MANAGEMENT ===================
        {
          name: 'create_table',
          description: 'Create a table in a new text frame (mm; rows include header and footer rows). Returns the frame id, the rows and the columns.',
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
          description: 'Create a new layer with a name, optional guide colour, visibility and lock state. Fails if the name is taken.',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Layer name' },
              color: { type: 'string', description: 'Layer colour name, e.g. RED, BLUE, LIGHT_BLUE, GREEN, YELLOW (error if unknown)' },
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
          description: 'Export the active document as PDF. preset: Print/HighQualityPrint = [High Quality Print], Web/SmallestFileSize = [Smallest File Size], PressQuality = [Press Quality], PDFX1a/PDFX3/PDFX4 = the PDF/X presets, or the exact name of any PDF preset in InDesign. includeBleed uses the document bleed; includeSlug the slug. Overwrites the file, so confirmDestructive is required.',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: { type: 'string', description: 'Output PDF file path' },
              preset: { type: 'string', description: 'Preset alias or exact PDF preset name', default: 'HighQualityPrint' },
              pageRange: { type: 'string', description: 'Pages, e.g. "1", "2-3", "1,3-5" or "all"', default: 'all' },
              includeBleed: { type: 'boolean', description: 'Use the document bleed', default: false },
              includeSlug: { type: 'boolean', description: 'Include the slug area', default: false },
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm file overwrite', default: false },
            },
            required: ['filePath'],
          },
        },
        {
          name: 'export_images',
          description: 'Export pages as PNG or JPEG, one file per page: <folder>/<documentname>_page<N>.<ext>. The folder is created if missing. Returns the list of files written.',
          inputSchema: {
            type: 'object',
            properties: {
              folderPath: { type: 'string', description: 'Output folder path' },
              format: { type: 'string', enum: ['PNG', 'JPEG'], default: 'PNG' },
              resolution: { type: 'number', description: 'Export resolution in dpi (e.g. 72, 100, 300)', default: 300 },
              pageRange: { type: 'string', description: 'Pages, e.g. "1", "2-3", "1,3-5" or "all"', default: 'all' },
              includeBleed: { type: 'boolean', description: 'Include the document bleed', default: false },
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm folder write access', default: false },
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
              imageFormat: { type: 'string', enum: ['AUTOMATIC', 'PNG', 'JPEG', 'GIF'], description: 'Image conversion', default: 'AUTOMATIC' },
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm file overwrite', default: false },
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm package creation', default: false },
            },
            required: ['folderPath'],
          },
        },

        // =================== UTILITIES & AUTOMATION ===================
        {
          name: 'execute_indesign_code',
          description: '⚠️ Execute custom ExtendScript code in InDesign (REQUIRES INDESIGN_ALLOW_ARBITRARY_CODE=1)',
          inputSchema: {
            type: 'object',
            properties: {
              code: { 
                type: 'string', 
                description: 'ExtendScript/JavaScript code to execute in InDesign. WARNING: Can access filesystem, network, and system APIs!' 
              },
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
              fitOption: { type: 'string', enum: ['FIT_PAGE', 'FIT_SPREAD', 'ACTUAL_SIZE'], default: 'FIT_PAGE' },
            },
          },
        },
        {
          name: 'data_merge',
          description: 'Merge the active document (which needs data field placeholders) with a CSV/TXT data source. Writes ONE merged file per format into outputFolder: <document>_merged.pdf and/or <document>_merged.indd',
          inputSchema: {
            type: 'object',
            properties: {
              dataSourcePath: { type: 'string', description: 'Path to CSV data source' },
              outputFolder: { type: 'string', description: 'Output folder (created if missing)' },
              fileFormat: { type: 'string', enum: ['INDD', 'PDF', 'BOTH'], default: 'PDF' },
              recordRange: { type: 'string', description: 'Records: "all", a single number ("3") or a range ("1-10")', default: 'all' },
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm bulk file creation', default: false },
            },
            required: ['dataSourcePath', 'outputFolder'],
          },
        },
    ];

    // A tool module may replace a core tool of the same name: its definition and handler win
    const allDefinitions = () => {
      if (!this._toolDefs) {
        const seen = new Set();
        this._toolDefs = buildToolDefinitions().filter((d) => !seen.has(d.name) && seen.add(d.name));
      }
      return this._toolDefs;
    };

    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: allDefinitions(),
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name } = request.params;
      let args = request.params.arguments;

      try {
        args = this.validateArguments(name, args, allDefinitions());
        if (this.extra.handlers[name]) return await this.extra.handlers[name](args || {});
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
          case 'get_selected_objects': return await this.getSelectedObjects();
          case 'get_text_content': return await this.getTextContent(args);
          case 'list_text_frames': return await this.listTextFrames(args);
          case 'analyze_embedded_objects': return await this.analyzeEmbeddedObjects(args);
          case 'insert_markdown_text': return await this.insertMarkdownText(args);
          case 'fix_typography_in_selection': return await this.fixTypographyInSelection(args);
          case 'find_typography_issues': return await this.findTypographyIssues(args);
          case 'clean_imported_text': return await this.cleanImportedText(args);
          case 'analyze_text_problems': return await this.analyzeTextProblems(args);
          case 'list_grep_searches': return await this.listGrepSearches();
          case 'create_text_frame': return await this.createTextFrame(args);
          case 'edit_text_frame': return await this.editTextFrame(args);
          case 'find_replace_text': return await this.findReplaceText(args);

          // Graphics Management
          case 'place_image': return await this.placeImage(args);
          case 'create_rectangle': return await this.createRectangle(args);
          case 'create_ellipse': return await this.createEllipse(args);

          // Style Management
          case 'create_paragraph_style': return await this.createParagraphStyle(args);
          case 'modify_paragraph_style': return await this.modifyParagraphStyle(args);
          case 'create_character_style': return await this.createCharacterStyle(args);
          case 'modify_character_style': return await this.modifyCharacterStyle(args);
          case 'create_object_style': return await this.createObjectStyle(args);
          case 'modify_object_style': return await this.modifyObjectStyle(args);
          case 'apply_paragraph_style': return await this.applyParagraphStyle(args);
          case 'apply_object_style': return await this.applyObjectStyle(args);
          case 'list_styles': return await this.listStyles(args);

          // Color Management
          case 'create_color_swatch': return await this.createColorSwatch(args);
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
            if (this.extra.handlers[name]) return await this.extra.handlers[name](args || {});
            throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
        }
      } catch (error) {
        throw new McpError(ErrorCode.InternalError, `Error executing tool ${name}: ${error.message}`);
      }
    });
  }

  // =================== CORE UTILITIES ===================
  // Directory for generated scripts. Scripts are deleted after a successful run and
  // kept after a failure so they can be inspected (see README, "Debugging").
  get scriptDir() {
    if (!this._scriptDir) {
      this._scriptDir = path.join(os.tmpdir(), 'indesign-mcp');
      fs.mkdirSync(this._scriptDir, { recursive: true });
    }
    return this._scriptDir;
  }

  get appName() {
    return process.env.INDESIGN_APP_NAME || 'Adobe InDesign 2026';
  }

  // Remove kept debug scripts older than three days
  _pruneScriptDir() {
    if (this._pruned) return;
    this._pruned = true;
    const cutoff = Date.now() - 3 * 24 * 3600 * 1000;
    try {
      for (const f of fs.readdirSync(this.scriptDir)) {
        const fp = path.join(this.scriptDir, f);
        const st = fs.statSync(fp);
        if (st.isFile() && st.mtimeMs < cutoff) fs.unlinkSync(fp);
      }
    } catch (e) { /* best effort */ }
  }

  // Run AppleScript source from a temp file (never via the command line)
  async executeAppleScript(script) {
    const file = path.join(this.scriptDir, `applescript_${Date.now()}_${Math.random().toString(36).slice(2)}.scpt`);
    fs.writeFileSync(file, script, 'utf8');
    try {
      const result = execSync(`osascript ${JSON.stringify(file)}`, {
        encoding: 'utf8',
        timeout: 60000,
        stdio: ['ignore', 'pipe', 'pipe'],
      });
      fs.unlinkSync(file);
      return result.trim();
    } catch (error) {
      const err = new Error(describeOsascriptError(error));
      err.debugFile = file;
      throw err;
    }
  }

  // Run ExtendScript in InDesign and return the script's last expression as a string.
  // Calls are serialised: InDesign handles one script at a time.
  executeInDesignScript(script) {
    const run = () => this._runInDesignScript(script);
    const next = (this._queue || Promise.resolve()).then(run, run);
    this._queue = next.catch(() => {});
    return next;
  }

  async _runInDesignScript(script) {
    const id = `${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const jsxFile = path.join(this.scriptDir, `script_${id}.jsx`);
    const resultFile = path.join(this.scriptDir, `result_${id}.txt`);

    fs.writeFileSync(jsxFile, buildWrapper(script, resultFile.replace(/\\/g, '/')), 'utf8');

    let failure = null;
    try {
      // The .jsx is started from a tiny AppleScript; no user data is part of it.
      await this.executeAppleScript(
        `tell application ${appleScriptStr(this.appName)}\n` +
        `  do script (POSIX file ${appleScriptStr(jsxFile)}) language javascript\n` +
        `end tell`
      );

      if (!fs.existsSync(resultFile)) {
        return '(no return value)';
      }
      const result = fs.readFileSync(resultFile, 'utf8');
      // Tool scripts report failures as plain strings; surface them as real errors
      // instead of letting formatResponse() present them as a success.
      if (/^(ERROR:|Error\b|No document open|No (text frame|objects?)\b|Cannot\b|Invalid\b|Document has never)/.test(result) ||
          /^[^\n]{0,200}\bnot found\b[^\n]*$/.test(result)) {
        failure = new Error(result.replace(/^ERROR:\s*/, ''));
        // "Code: n/a" = a validation error raised by our own script (bad id, missing style...),
        // not an InDesign failure: no need to keep the script.
        failure.expected = /^ERROR:/.test(result) ? /\(Code: (n\/a|1),/.test(result) : true;
        throw failure;
      }
      return result;
    } catch (error) {
      failure = error;
      if (!error.expected) {
        // Keep the script for debugging; point at it without dumping osascript output
        console.error(`[indesign-mcp] script failed, kept for debugging: ${jsxFile}`);
        error.message = `${error.message} [debug script: ${jsxFile}]`;
      }
      throw error;
    } finally {
      if (!failure || failure.expected) {
        for (const f of [jsxFile, resultFile]) {
          if (fs.existsSync(f)) fs.unlinkSync(f);
        }
      } else if (fs.existsSync(resultFile)) {
        fs.unlinkSync(resultFile);
      }
      this._pruneScriptDir();
    }
  }

  // Embed a value in generated ExtendScript as a safe string literal
  jsStr(value) { return jsStr(value); }
  jsJson(value) { return jsJson(value); }
  // Validate a value that is spliced into generated code as a bare identifier (e.g. an enum name)
  jsEnum(value, label) { return jsEnum(value, label); }
  // Coerce a value to a finite number for embedding in generated code
  jsNum(value, label) { return jsNum(value, label); }

  // JS expression that resolves a text frame from frameId, or pageIndex + frameIndex
  frameExpr(args, docVar = 'doc') {
    const num = (v, label) => (v === undefined || v === null ? 'null' : this.jsNum(v, label));
    return `__textFrame(${docVar}, ${num(args.frameId, 'frameId')}, ${num(args.pageIndex, 'pageIndex')}, ${num(args.frameIndex, 'frameIndex')})`;
  }

  // Convert user text to InDesign text: "\n" -> paragraph (\r) or forced line break (U+2028); "<br>" -> forced line break
  toInDesignText(content, lineBreak = 'paragraph') {
    if (lineBreak !== 'paragraph' && lineBreak !== 'forced') {
      throw new Error(`Invalid lineBreak: ${lineBreak} (use "paragraph" or "forced")`);
    }
    // In InDesign text a paragraph break is CR (13) and a forced line break (Shift+Enter) is LF (10).
    // (U+2028 is NOT a line break there; it is stored as an ordinary, space-like character.)
    const FORCED = '\u0000';
    return String(content)
      .replace(/\r\n?/g, '\n')
      .replace(/<br\s*\/?>/gi, FORCED)
      .replace(/\n/g, lineBreak === 'forced' ? FORCED : '\r')
      .replace(new RegExp(FORCED, 'g'), '\n');
  }

  // Statement block that sets `textFrame`: the selected frame, or frameId / pageIndex + frameIndex
  textFrameAcquire(args) {
    if (args.useSelectedFrame) {
      return `var textFrame;
          if (app.selection.length === 0 || app.selection[0].constructor.name !== "TextFrame") {
            throw new Error("No text frame selected. Please select a text frame first.");
          }
          textFrame = app.selection[0];`;
    }
    return `var textFrame = ${this.frameExpr(args)};`;
  }

  // Build ExtendScript that runs GREP find/change steps on `story`, keeping all formatting.
  // steps: [{ find, change, label }] (label uses %n for the count)
  grepSteps(steps) {
    return steps.map((st) => `
              n = __grep(story, ${this.jsStr(st.find)}, ${this.jsStr(st.change)});
              if (n > 0) { changes += n; changeLog += ${this.jsStr('✓ ' + st.label)}.replace("%n", n) + "\\n"; }`).join('');
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
      var doc = __requireDoc();
      var dp = doc.documentPreferences;
      var mp = doc.marginPreferences;
      var intentName = String(dp.intent);
      // (ExtendScript mis-parses chained ternaries with ===, so use plain ifs)
      var intent = "Mobile";
      if (intentName.indexOf("PRINT") === 0) { intent = "Print"; }
      else if (intentName.indexOf("WEB") === 0) { intent = "Web"; }
      var info = "=== DOCUMENT INFORMATION ===\\n";
      info += "Name: " + doc.name + "\\n";
      info += "Pages: " + doc.pages.length + "\\n";
      info += "Width: " + __r(dp.pageWidth) + " mm\\n";
      info += "Height: " + __r(dp.pageHeight) + " mm\\n";
      info += "Facing Pages: " + dp.facingPages + "\\n";
      info += "Modified: " + doc.modified + "\\n";
      var filePath = "Unsaved";
      if (__hasFile(doc)) filePath = doc.fullName.fsName;
      info += "File Path: " + filePath + "\\n";
      info += "\\n=== BLEED / SLUG (mm) ===\\n";
      info += "Bleed top/bottom/inside/outside: " + __r(dp.documentBleedTopOffset) + " / " + __r(dp.documentBleedBottomOffset) + " / " + __r(dp.documentBleedInsideOrLeftOffset) + " / " + __r(dp.documentBleedOutsideOrRightOffset) + "\\n";
      info += "Slug top/bottom/inside/outside: " + __r(dp.slugTopOffset) + " / " + __r(dp.slugBottomOffset) + " / " + __r(dp.slugInsideOrLeftOffset) + " / " + __r(dp.slugRightOrOutsideOffset) + "\\n";
      info += "\\n=== MARGINS (mm, document defaults) ===\\n";
      info += "Top: " + __r(mp.top) + "\\n";
      info += "Bottom: " + __r(mp.bottom) + "\\n";
      info += "Left: " + __r(mp.left) + "\\n";
      info += "Right: " + __r(mp.right) + "\\n";
      info += "\\n=== COLOUR ===\\n";
      info += "Intent: " + intent + "\\n";
      try { info += "CMYK profile: " + doc.cmykProfile + "\\n"; } catch (e) {}
      try { info += "RGB profile: " + doc.rgbProfile + "\\n"; } catch (e) {}
      info += "\\n=== CONTENT SUMMARY ===\\n";
      var totalTextFrames = 0, totalRects = 0, totalShapes = 0;
      for (var i = 0; i < doc.pages.length; i++) {
        totalTextFrames += doc.pages[i].textFrames.length;
        totalRects += doc.pages[i].rectangles.length;
        totalShapes += doc.pages[i].ovals.length + doc.pages[i].polygons.length;
      }
      info += "Text Frames: " + totalTextFrames + "\\n";
      info += "Rectangles (incl. image frames): " + totalRects + "\\n";
      info += "Ovals/Polygons: " + totalShapes + "\\n";
      info += "Layers: " + doc.layers.length + "\\n";
      info += "\\n=== SWATCHES (" + doc.swatches.length + ") ===\\n";
      for (var s = 0; s < doc.swatches.length; s++) {
        var sw = doc.swatches[s].getElements()[0];
        var line = "  " + sw.name;
        try {
          if (sw.constructor.name === "Color") {
            line += " [" + String(sw.space) + " " + (String(sw.model) === "SPOT" ? "spot" : "process") + ": ";
            var cv = sw.colorValue, cvs = [];
            for (var c = 0; c < cv.length; c++) cvs.push(__r(cv[c]));
            line += cvs.join(",") + "]";
          }
        } catch (e) {}
        info += line + "\\n";
      }
      info;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Document Info");
  }

  async createDocument(args) {
    const {
      preset = 'A4',
      width,
      height,
      orientation,
      pages = 1,
      facingPages = false,
      bleed = 0,
      slug = 0,
      marginTop = 20,
      marginBottom = 20,
      marginLeft = 20,
      marginRight = 20
    } = args;

    // Page sizes in mm (portrait)
    const presets = {
      A3: [297, 420],
      A4: [210, 297],
      A5: [148, 210],
      Letter: [215.9, 279.4],
      Legal: [215.9, 355.6],
    };

    if (orientation !== undefined && !['portrait', 'landscape'].includes(String(orientation).toLowerCase())) {
      throw new Error(`Invalid orientation: ${orientation} (use Portrait or Landscape)`);
    }
    const landscape = orientation !== undefined && String(orientation).toLowerCase() === 'landscape';
    const portrait = orientation !== undefined && !landscape;

    let pageWidth, pageHeight;
    if (preset === 'Custom') {
      if (width === undefined || height === undefined) {
        throw new Error('Custom preset requires width and height (in mm)');
      }
      pageWidth = this.jsNum(width, 'width');
      pageHeight = this.jsNum(height, 'height');
      // An explicit orientation is honoured by swapping; without one, width/height are final
      if ((landscape && pageWidth < pageHeight) || (portrait && pageWidth > pageHeight)) {
        [pageWidth, pageHeight] = [pageHeight, pageWidth];
      }
    } else if (presets[preset]) {
      [pageWidth, pageHeight] = presets[preset];
      if (landscape) {
        [pageWidth, pageHeight] = [pageHeight, pageWidth];
      }
    } else {
      throw new Error(`Unknown preset: ${preset} (use ${Object.keys(presets).join(', ')} or Custom)`);
    }

    const pageCount = Math.max(1, Math.floor(this.jsNum(pages, 'pages')));
    const bleedMm = this.jsNum(bleed, 'bleed');
    const slugMm = this.jsNum(slug, 'slug');
    const mTop = this.jsNum(marginTop, 'marginTop');
    const mBottom = this.jsNum(marginBottom, 'marginBottom');
    const mLeft = this.jsNum(marginLeft, 'marginLeft');
    const mRight = this.jsNum(marginRight, 'marginRight');

    const script = `
      var doc = app.documents.add();
      
      // Millimetres, and coordinates relative to each page (origin = top-left of the trim area)
      doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
      doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
      doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
      
      // Page size, page count, facing pages
      doc.documentPreferences.pageWidth = "${pageWidth}mm";
      doc.documentPreferences.pageHeight = "${pageHeight}mm";
      doc.documentPreferences.facingPages = ${facingPages ? 'true' : 'false'};
      doc.documentPreferences.pagesPerDocument = ${pageCount};
      
      // Bleed and slug
      doc.documentPreferences.documentBleedUniformSize = false;
      doc.documentPreferences.documentBleedTopOffset = "${bleedMm}mm";
      doc.documentPreferences.documentBleedBottomOffset = "${bleedMm}mm";
      doc.documentPreferences.documentBleedInsideOrLeftOffset = "${bleedMm}mm";
      doc.documentPreferences.documentBleedOutsideOrRightOffset = "${bleedMm}mm";
      doc.documentPreferences.documentSlugUniformSize = false;
      doc.documentPreferences.slugTopOffset = "${slugMm}mm";
      doc.documentPreferences.slugBottomOffset = "${slugMm}mm";
      doc.documentPreferences.slugInsideOrLeftOffset = "${slugMm}mm";
      doc.documentPreferences.slugRightOrOutsideOffset = "${slugMm}mm";
      
      // Margins: set on the document defaults, the master pages and the existing pages
      doc.marginPreferences.top = "${mTop}mm";
      doc.marginPreferences.bottom = "${mBottom}mm";
      doc.marginPreferences.left = "${mLeft}mm";
      doc.marginPreferences.right = "${mRight}mm";
      var targets = [];
      for (var m = 0; m < doc.masterSpreads.length; m++) {
        for (var mp = 0; mp < doc.masterSpreads[m].pages.length; mp++) targets.push(doc.masterSpreads[m].pages[mp]);
      }
      for (var pg = 0; pg < doc.pages.length; pg++) targets.push(doc.pages[pg]);
      for (var t = 0; t < targets.length; t++) {
        targets[t].marginPreferences.top = "${mTop}mm";
        targets[t].marginPreferences.bottom = "${mBottom}mm";
        targets[t].marginPreferences.left = "${mLeft}mm";
        targets[t].marginPreferences.right = "${mRight}mm";
      }
      
      "Document created: " + doc.name + " - " + ${this.jsStr(preset)} + " (" + __r(doc.documentPreferences.pageWidth) + " x " + __r(doc.documentPreferences.pageHeight) + " mm), " + 
      doc.pages.length + " pages, " + (doc.documentPreferences.facingPages ? "facing pages" : "single pages");
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Document");
  }

  async openDocument(args) {
    const { filePath } = args;
    
    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);
    
    const script = `
      try {
        var file = File(${this.jsStr(validatedPath)});
        if (!file.exists) {
          "File not found: ${validatedPath}";
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
    const { filePath, name, confirmDestructive = false } = args;

    // Security: the target must be inside the allowed folders (symlinks resolved)
    const validatedPath = filePath ? this.validateFilePath(filePath) : null;
    if (validatedPath && !/\.ind[dt]$/i.test(validatedPath)) throw new Error('filePath must end with .indd (or .indt for a template)');

    const script = `
      var doc = ${name ? `app.documents.itemByName(${this.jsStr(name)})` : '__requireDoc()'};
      ${name ? `if (!doc.isValid) throw new Error("Document not found: " + ${this.jsStr(name)});` : ''}
      var hasFile = __hasFile(doc);
      ${validatedPath ? `
        var target = File(${this.jsStr(validatedPath)});
        if (hasFile && doc.fullName.fsName === target.fsName) {
          // Save As onto its own file: an ordinary save (InDesign would refuse to replace an open file)
          doc.save();
          "Document saved: " + doc.name + " (" + doc.fullName.fsName + "; filePath is the file the document already has, so this was a plain save)";
        } else {
          for (var i = 0; i < app.documents.length; i++) {
            var other = app.documents[i];
            if (other.id === doc.id || !__hasFile(other)) continue;
            if (other.fullName.fsName === target.fsName) {
              throw new Error("The file " + target.fsName + " is already open in the document " + other.name + ". Close that document first, or choose another file name.");
            }
          }
          if (target.exists && !${confirmDestructive ? 'true' : 'false'}) {
            throw new Error("The file " + target.fsName + " already exists. Pass confirmDestructive: true to overwrite it.");
          }
          if (!target.parent.exists) target.parent.create();
          var oldName = doc.name;
          doc.save(target);
          "Document saved as: " + target.fsName + " (Save As: the document " + oldName + " is now named " + doc.name + "; " + app.documents.length + " document(s) open)";
        }
      ` : `
        if (!hasFile) throw new Error("Document " + doc.name + " has never been saved: pass filePath (the .indd file to create) to save it.");
        doc.save();
        "Document saved: " + doc.name + " (" + doc.fullName.fsName + ")";
      `}
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Save Document");
  }

  async closeDocument(args) {
    const { name, save = false, confirmDestructive = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open to close";
      } else {
        var doc = ${name ? `app.documents.itemByName(${this.jsStr(name)})` : 'app.activeDocument'};
        if (!doc.isValid) {
          "Document not found: " + ${this.jsStr(name || '')};
        } else {
          var docName = doc.name;
          if (${save ? 'true' : 'false'}) {
            if (!__hasFile(doc)) throw new Error("Document " + docName + " has never been saved; save_document with a filePath first.");
            doc.close(SaveOptions.YES);
            "Document saved and closed: " + docName;
          } else if (doc.modified && !${confirmDestructive ? 'true' : 'false'}) {
            throw new Error("Document " + docName + " has unsaved changes. Pass save=true, or confirmDestructive=true to discard them.");
          } else {
            doc.close(SaveOptions.NO);
            "Document closed: " + docName;
          }
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
            var master = doc.masterSpreads.itemByName(${this.jsStr(masterPage)});
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
    
    // Security: Require confirmation for page deletion
    this.validateDestructiveOperation(args, 'DELETE PAGE', `page ${pageIndex + 1} and all its content`);
    
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
    if (!['before', 'after', 'end'].includes(position)) throw new Error(`Invalid position: ${position} (use before, after or end)`);
    const idx = this.jsNum(pageIndex, 'pageIndex');
    const at = position === 'end' ? 'LocationOptions.AT_END' : position === 'before' ? 'LocationOptions.BEFORE' : 'LocationOptions.AFTER';

    const script = `
      var doc = __requireDoc();
      var source = __page(doc, ${idx});
      // Page.duplicate copies the page with all its content (each object exactly once, groups intact)
      var copy = ${position === 'end' ? `source.duplicate(${at})` : `source.duplicate(${at}, source)`};
      "Page " + (${idx} + 1) + " duplicated. New page position: " + (copy.documentOffset + 1) + " (" + copy.allPageItems.length + " objects). Total pages: " + doc.pages.length;
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
  
  async getSelectedObjects() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var selection = app.selection;
        
        if (selection.length === 0) {
          "No objects selected. Please select a text frame or other object first.";
        } else {
          var result = "=== SELECTED OBJECTS ===\\n";
          
          for (var i = 0; i < selection.length; i++) {
            var obj = selection[i];
            result += "Object " + i + ": ";
            
            if (obj.hasOwnProperty('contents')) {
              // Text frame
              result += "Text Frame";
              var content = String(obj.contents).substring(0, 50);
              if (String(obj.contents).length > 50) content += "...";
              result += " - Content: " + content;
              
              // Find frame index on current page
              var currentPage = app.activeWindow.activePage;
              for (var j = 0; j < currentPage.textFrames.length; j++) {
                if (currentPage.textFrames[j] === obj) {
                  result += " (Frame Index: " + j + ")";
                  break;
                }
              }
            } else if (obj.hasOwnProperty('geometricBounds')) {
              result += "Shape/Image";
            } else {
              result += "Unknown object type";
            }
            result += "\\n";
          }
          
          result;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Get Selected Objects");
  }

  async getTextContent(args) {
    const { normalizeSpaces = true, maxLength = 0 } = args;
    const wantsFrame = args.frameId !== undefined || args.frameIndex !== undefined;
    const limit = this.jsNum(maxLength, 'maxLength');

    const script = `
      var doc = __requireDoc();
      var frame = null, textContent = null, source = "";
      
      ${wantsFrame ? `
        frame = ${this.frameExpr(args)};
        textContent = frame.parentStory.contents;
        source = "Text frame id " + frame.id;
      ` : `
        // No frame given: use the selection
        if (app.selection.length > 0) {
          var sel = app.selection[0];
          if (sel.constructor.name === "TextFrame") {
            frame = sel; textContent = sel.parentStory.contents; source = "Selected text frame id " + sel.id;
          } else if (sel.parentTextFrames && sel.parentTextFrames.length > 0) {
            frame = sel.parentTextFrames[0]; textContent = frame.parentStory.contents; source = "Text frame id " + frame.id + " (from selected text)";
          } else {
            throw new Error("Selected object (" + sel.constructor.name + ") has no text. Select a text frame or pass frameId / frameIndex.");
          }
        } else {
          throw new Error("No text frame selected. Select one or pass frameId / frameIndex.");
        }
      `}
      
      var result = "=== TEXT CONTENT ===\\n";
      result += "Source: " + source + "\\n";
      result += "(whole story, including any overflow text)\\n";
      result += "Original length: " + textContent.length + " characters\\n";
      result += "Paragraphs: " + frame.parentStory.paragraphs.length + "\\n";
      result += __overflowNote(frame) + "\\n\\n";
      
      // Forced line breaks (Shift+Enter, character 10) are shown as an arrow in both modes, never as a space
      var processedText = textContent.replace(/\\n/g, "\u21b5");
      ${normalizeSpaces ? `
        processedText = processedText.replace(/\\r+/g, ' ');
        while (processedText.indexOf('  ') !== -1) { processedText = processedText.replace(/  /g, ' '); }
        while (processedText.charAt(0) === ' ') { processedText = processedText.substring(1); }
        while (processedText.charAt(processedText.length - 1) === ' ') { processedText = processedText.substring(0, processedText.length - 1); }
      ` : ''}
      ${limit > 0 ? `
        if (processedText.length > ${limit}) {
          processedText = processedText.substring(0, ${limit}) + "...";
          result += "Text truncated to ${limit} characters\\n\\n";
        }
      ` : ''}
      result += "TEXT:\\n" + processedText;
      result;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Get Text Content");
  }

  async listTextFrames(args) {
    const pageIndex = this.jsNum(args.pageIndex === undefined ? 0 : args.pageIndex, 'pageIndex');

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${pageIndex});
      var result = "=== TEXT FRAMES ON PAGE " + (${pageIndex} + 1) + " ===\\n";
      result += "(index 0 = frontmost frame; indices change when frames are added or reordered - use id)\\n";
      if (page.textFrames.length === 0) {
        result += "No text frames on this page.";
      } else {
        for (var i = 0; i < page.textFrames.length; i++) {
          var f = page.textFrames[i];
          var b = f.geometricBounds;
          var preview = f.contents.substring(0, 60);
          if (f.contents.length > 60) preview += "...";
          preview = preview.replace(/[\\r\\n\\u2028]/g, "↵").replace(/\\t/g, "→");
          result += "Frame " + i + ": id=" + f.id + " x=" + __r(b[1]) + " y=" + __r(b[0]) + " w=" + __r(b[3] - b[1]) + " h=" + __r(b[2] - b[0]) +
                    " rotation=" + __r(f.rotationAngle) + " " + __overflowNote(f) + " paragraphs=" + f.parentStory.paragraphs.length +
                    (f.contents.length > 0 ? " style=" + JSON_STR(f.paragraphs[0].appliedParagraphStyle.name) + " size=" + __r(f.characters[0].pointSize) + "pt font=" + JSON_STR(f.characters[0].appliedFont.name.replace("\\t", " ")) : "") +
                    " text=" + (preview.length === 0 ? "(empty)" : JSON_STR(preview)) + "\\n";
        }
      }
      result;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "List Text Frames");
  }

  async analyzeEmbeddedObjects(args) {
    const { frameIndex, pageIndex = 0, maxObjects = 5 } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        var frame = null;
        
        // An explicit frame wins; otherwise use the selection
        ${(args.frameId !== undefined || args.frameIndex !== undefined) ? `
          frame = ${this.frameExpr(args)};
        ` : `
          if (app.selection.length > 0 && app.selection[0].constructor.name === "TextFrame") {
            frame = app.selection[0];
          }
        `}
        
        if (!frame) {
          throw new Error("No text frame found. Select a frame or pass frameId / frameIndex.");
        } else {
          var result = "=== EMBEDDED OBJECTS ANALYSIS ===\\n\\n";
          
          // Check if frame contains a table
          var hasTable = false;
          var table = null;
          
          try {
            if (frame.texts && frame.texts.length > 0 && frame.texts[0].tables && frame.texts[0].tables.length > 0) {
              hasTable = true;
              table = frame.texts[0].tables[0];
            }
          } catch (e) {}
          
          if (hasTable) {
            result += "FRAME TYPE: Contains TABLE\\n";
            result += "Table: " + table.rows.length + " rows x " + table.columns.length + " columns\\n\\n";
            
            // Analyze first few cells
            result += "=== FIRST " + Math.min(${maxObjects}, table.cells.length) + " CELLS ===\\n";
            for (var i = 0; i < Math.min(${maxObjects}, table.cells.length); i++) {
              var cell = table.cells[i];
              result += "\\nCell " + i + " (Row " + cell.rowIndex + ", Col " + cell.columnIndex + "):\\n";
              result += "  Content: " + String(cell.contents).substring(0, 100) + "\\n";
              
              // Check for embedded objects in cell
              if (cell.epstexts && cell.epstexts.length > 0) {
                result += "  EPSTexts: " + cell.epstexts.length + "\\n";
              }
              if (cell.pageItems && cell.pageItems.length > 0) {
                result += "  Page Items: " + cell.pageItems.length + "\\n";
                
                // Check first page item
                var pItem = cell.pageItems[0];
                result += "  First Item Type: " + pItem.constructor.name + "\\n";
                
                // If it's a group, check inside
                if (pItem.constructor.name === "Group" && pItem.allPageItems) {
                  result += "  Group contains: " + pItem.allPageItems.length + " items\\n";
                  if (pItem.allPageItems.length > 0) {
                    result += "  Sub-item type: " + pItem.allPageItems[0].constructor.name + "\\n";
                  }
                }
              }
            }
          } else {
            result += "FRAME TYPE: Regular text frame\\n\\n";
          }
          
          // Check for different types of embedded content
          result += "EPSTexts (formulas/EPS): " + frame.epstexts.length + "\\n";
          result += "Page Items (anchored): " + frame.pageItems.length + "\\n";
          result += "All Page Items: " + frame.allPageItems.length + "\\n\\n";
          
          // Analyze EPSTexts (MathML formulas are often EPS)
          if (frame.epstexts.length > 0) {
            result += "=== EPS TEXTS (First " + Math.min(${maxObjects}, frame.epstexts.length) + ") ===\\n";
            for (var i = 0; i < Math.min(${maxObjects}, frame.epstexts.length); i++) {
              var eps = frame.epstexts[i];
              result += "\\nObject " + i + ":\\n";
              result += "  Label: " + eps.label + "\\n";
              result += "  ID: " + eps.id + "\\n";
              
              // Try to get fill color
              try {
                if (eps.fillColor && eps.fillColor.name) {
                  result += "  Fill Color: " + eps.fillColor.name;
                  if (eps.fillColor.space) {
                    result += " (" + eps.fillColor.space + ")";
                  }
                  result += "\\n";
                }
              } catch (e) {
                result += "  Fill Color: (cannot access)\\n";
              }
              
              // Try to get bounds
              try {
                if (eps.geometricBounds) {
                  result += "  Bounds: [" + eps.geometricBounds.join(", ") + "]\\n";
                }
              } catch (e) {}
            }
          }
          
          // Analyze PageItems
          if (frame.pageItems.length > 0) {
            result += "\\n=== PAGE ITEMS (First " + Math.min(${maxObjects}, frame.pageItems.length) + ") ===\\n";
            for (var i = 0; i < Math.min(${maxObjects}, frame.pageItems.length); i++) {
              var item = frame.pageItems[i];
              result += "\\nItem " + i + ":\\n";
              result += "  Type: " + item.constructor.name + "\\n";
              result += "  Label: " + item.label + "\\n";
              
              // List available properties (only for first item)
              if (i === 0) {
                result += "  Properties: ";
                var props = [];
                for (var prop in item) {
                  try {
                    if (typeof item[prop] !== 'function') {
                      props.push(prop);
                    }
                  } catch (e) {
                    // Skip properties that throw errors
                  }
                }
                result += props.slice(0, 30).join(", ") + "\\n";
              }
              
              // Check if it's a group or has sub-items
              try {
                if (item.allPageItems && item.allPageItems.length > 0) {
                  result += "  Has " + item.allPageItems.length + " sub-items\\n";
                  var firstItem = item.allPageItems[0];
                  result += "  First sub-item type: " + firstItem.constructor.name + "\\n";
                }
              } catch (e) {}
              
              // Check content type
              try {
                if (item.contentType) {
                  result += "  Content Type: " + item.contentType + "\\n";
                }
              } catch (e) {}
              
              // Check if it's a rectangle with graphic
              try {
                if (item.graphics && item.graphics.length > 0) {
                  result += "  Has Graphics: " + item.graphics.length + "\\n";
                  var graphic = item.graphics[0];
                  result += "  Graphic Type: " + graphic.constructor.name + "\\n";
                  
                  // Try to get the actual file link
                  if (graphic.itemLink && graphic.itemLink.filePath) {
                    result += "  Linked File: " + graphic.itemLink.filePath + "\\n";
                  }
                }
              } catch (e) {}
              
              // Try alternative access via allGraphics
              try {
                if (item.allGraphics && item.allGraphics.length > 0) {
                  result += "  AllGraphics: " + item.allGraphics.length + "\\n";
                  var gfx = item.allGraphics[0];
                  result += "  Graphic Type: " + gfx.constructor.name + "\\n";
                  
                  // Try to access EPS/PDF content
                  if (gfx.itemLink) {
                    result += "  Link Name: " + gfx.itemLink.name + "\\n";
                    result += "  Link Status: " + gfx.itemLink.status + "\\n";
                  }
                  
                  // Try to get PDF/EPS data
                  if (gfx.pdfAttributes) {
                    result += "  Has PDF Attributes\\n";
                  }
                  if (gfx.epsText) {
                    result += "  Has EPS Text\\n";
                  }
                }
              } catch (e) {
                result += "  AllGraphics Error: " + e.message + "\\n";
              }
              
              // Try to access XML content (MathML)
              try {
                if (item.associatedXMLElement) {
                  var xmlElem = item.associatedXMLElement;
                  result += "  Has XML Element: YES\\n";
                  result += "  XML Tag: " + xmlElem.markupTag.name + "\\n";
                  
                  // Try to get MathML content
                  if (xmlElem.contents) {
                    var xmlContent = String(xmlElem.contents).substring(0, 500);
                    result += "  XML Content (first 500 chars):\\n    " + xmlContent.replace(/\\n/g, "\\n    ") + "\\n";
                  }
                }
              } catch (e) {
                result += "  XML Error: " + e.message + "\\n";
              }
              
              // Check if it contains EPSText
              try {
                if (item.epstexts && item.epstexts.length > 0) {
                  result += "  Contains EPSTexts: " + item.epstexts.length + "\\n";
                  var eps = item.epstexts[0];
                  result += "  EPS Label: " + eps.label + "\\n";
                  
                  // Try to get EPS content
                  if (eps.epsContent) {
                    var epsContent = String(eps.epsContent).substring(0, 500);
                    result += "  EPS Content (first 500 chars):\\n    " + epsContent.replace(/\\n/g, "\\n    ") + "\\n";
                  }
                }
              } catch (e) {}
              
              // Try to get fill color
              try {
                if (item.fillColor && item.fillColor.name) {
                  result += "  Fill Color: " + item.fillColor.name;
                  if (item.fillColor.space) {
                    result += " (" + item.fillColor.space + ")";
                  }
                  result += "\\n";
                }
              } catch (e) {}
            }
          }
          
          result;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Analyze Embedded Objects");
  }

  // Markdown -> { text, paras: [{ start, style }], ranges: [{ start, end, kind }] } (offsets in characters)
  parseMarkdown(markdown) {
    const lines = String(markdown).replace(/\r\n?/g, '\n').split('\n');
    const paras = [];
    const ranges = [];
    let text = '';

    const parseInline = (src) => {
      // Tokenise; unmatched emphasis markers are demoted to literal text.
      const forcedLiteral = new Set();
      for (;;) {
        const tokens = [];
        let buf = '';
        let i = 0;
        const flush = () => { if (buf) { tokens.push({ t: 'text', v: buf }); buf = ''; } };
        while (i < src.length) {
          const ch = src[i];
          if (ch === '\\' && i + 1 < src.length && /[*\\#]/.test(src[i + 1])) { buf += src[i + 1]; i += 2; continue; }
          let marker = null;
          for (const m of ['***', '**', '*']) { if (src.startsWith(m, i)) { marker = m; break; } }
          if (marker && !forcedLiteral.has(i)) {
            tokens.push({ t: 'text', v: buf }); buf = '';
            tokens[tokens.length - 1].v === '' && tokens.pop();
            tokens.push({ t: 'mark', v: marker, at: i, prev: src[i - 1], next: src[i + marker.length] });
            i += marker.length;
            continue;
          }
          buf += ch; i++;
        }
        flush();

        // Resolve marks
        let bold = false, italic = false, boldAt = -1, italicAt = -1;
        const out = []; // { v, bold, italic }
        for (const tok of tokens) {
          if (tok.t === 'text') { out.push({ v: tok.v, bold, italic }); continue; }
          const canOpen = tok.next !== undefined && !/\s/.test(tok.next);
          const canClose = tok.prev !== undefined && !/\s/.test(tok.prev);
          const wantBold = tok.v === '**' || tok.v === '***';
          const wantItalic = tok.v === '*' || tok.v === '***';
          const closing = (!wantBold || bold) && (!wantItalic || italic) && canClose;
          const opening = (!wantBold || !bold) && (!wantItalic || !italic) && canOpen;
          if (closing) {
            if (wantBold) { bold = false; boldAt = -1; }
            if (wantItalic) { italic = false; italicAt = -1; }
          } else if (opening) {
            if (wantBold) { bold = true; boldAt = tok.at; }
            if (wantItalic) { italic = true; italicAt = tok.at; }
          } else {
            out.push({ v: tok.v, bold, italic }); // literal
          }
        }
        if (bold || italic) {
          // Unclosed emphasis: make its opener literal and parse again
          if (bold && boldAt >= 0) forcedLiteral.add(boldAt);
          if (italic && italicAt >= 0) forcedLiteral.add(italicAt);
          continue;
        }
        return out.filter((r) => r.v !== '');
      }
    };

    for (const line of lines) {
      if (!line.trim()) continue;
      const h = line.match(/^(#{1,6})[ \t]+(.*)$/);
      const level = h ? h[1].length : 0;
      const runs = parseInline(h ? h[2] : line);
      const paraText = runs.map((r) => r.v).join('');
      if (!paraText.trim()) continue;
      if (paras.length > 0) text += '\r';
      const start = text.length;
      paras.push({ start, style: level ? `h${level}` : 'body' });
      let pos = start;
      for (const r of runs) {
        const kind = r.bold && r.italic ? 'boldItalic' : r.bold ? 'bold' : r.italic ? 'italic' : null;
        if (kind && !level) ranges.push({ start: pos, end: pos + r.v.length - 1, kind });
        else if (kind && level) ranges.push({ start: pos, end: pos + r.v.length - 1, kind });
        pos += r.v.length;
      }
      text += paraText;
    }
    return { text, paras, ranges };
  }

  async insertMarkdownText(args) {
    const { markdownText, useSelectedFrame = false, replaceContent = true, bodyStyle, styleMap = {} } = args;
    if (typeof markdownText !== 'string') {
      throw new Error('markdownText is required');
    }

    const parsed = this.parseMarkdown(markdownText);
    if (parsed.paras.length === 0) {
      throw new Error('markdownText contains no text');
    }

    // Style names (paragraph styles for h1-h6/body, character styles for bold/italic/boldItalic)
    const heading = (n) => [`Heading ${n}`, `Header ${n}`, `H${n}`, `Header${n}`, `Überschrift ${n}`, `Ueberschrift ${n}`];
    const candidates = {
      h1: heading(1), h2: heading(2), h3: heading(3), h4: heading(4), h5: heading(5), h6: heading(6),
      body: bodyStyle ? [bodyStyle] : [],
      bold: ['Bold', 'Strong', 'Fett'],
      italic: ['Italic', 'Emphasis', 'Kursiv'],
      boldItalic: ['Bold Italic', 'BoldItalic', 'Strong Emphasis', 'Fett Kursiv'],
    };
    for (const [key, name] of Object.entries(styleMap)) {
      if (!(key in candidates)) throw new Error(`Unknown styleMap key: ${key} (use h1-h6, body, bold, italic, boldItalic)`);
      candidates[key] = [String(name)];
    }
    const kinds = { h1: 'p', h2: 'p', h3: 'p', h4: 'p', h5: 'p', h6: 'p', body: 'p', bold: 'c', italic: 'c', boldItalic: 'c' };
    const used = new Set([...parsed.paras.map((p) => p.style), ...parsed.ranges.map((r) => r.kind)]);
    const required = [...used].filter((k) => candidates[k].length > 0).map((k) => ({ key: k, kind: kinds[k], names: candidates[k] }));

    const data = { text: parsed.text, paras: parsed.paras, ranges: parsed.ranges, required };

    const script = `
      var doc = __requireDoc();
      var D = ${this.jsJson(data)};
      var frame;
      ${useSelectedFrame ? `
        if (app.selection.length === 0 || app.selection[0].constructor.name !== "TextFrame") {
          throw new Error("No text frame selected. Select one or pass frameId / frameIndex.");
        }
        frame = app.selection[0];
      ` : `frame = ${this.frameExpr(args)};`}
      
      // 1. Resolve every style first, so a missing style changes nothing
      var resolved = {}, missing = [];
      for (var r = 0; r < D.required.length; r++) {
        var req = D.required[r], hit = null;
        for (var n = 0; n < req.names.length && !hit; n++) {
          var cand = (req.kind === "p" ? doc.paragraphStyles : doc.characterStyles).itemByName(req.names[n]);
          if (cand.isValid) hit = cand;
        }
        if (hit) resolved[req.key] = hit;
        else missing.push((req.kind === "p" ? "paragraph style" : "character style") + " for " + req.key + " (tried: " + req.names.join(", ") + ")");
      }
      if (missing.length > 0) {
        throw new Error("Missing styles - create them first (create_paragraph_style / create_character_style) or map other names with styleMap: " + missing.join("; "));
      }
      
      // 2. Insert the text
      var story = frame.parentStory;
      var offset = 0;
      ${replaceContent ? `
        story.contents = D.text;
      ` : `
        if (story.characters.length > 0) {
          offset = story.characters.length + 1;
          story.insertionPoints[-1].contents = "\\r" + D.text;
        } else {
          story.contents = D.text;
        }
      `}
      
      // 3. Paragraph styles, then character styles
      for (var p = 0; p < D.paras.length; p++) {
        if (resolved[D.paras[p].style]) {
          story.characters[offset + D.paras[p].start].paragraphs[0].appliedParagraphStyle = resolved[D.paras[p].style];
        }
      }
      for (var c = 0; c < D.ranges.length; c++) {
        var rg = D.ranges[c];
        story.characters.itemByRange(offset + rg.start, offset + rg.end).appliedCharacterStyle = resolved[rg.kind];
      }
      
      "Markdown inserted into frame id " + frame.id + ": " + D.paras.length + " paragraphs, " + D.ranges.length + " character runs. " + __overflowNote(frame);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Insert Markdown Text");
  }

  async fixTypographyInSelection(args) {
    const { fixDates = true, fixQuotes = true, fixDashes = true, fixSpaces = true } = args;
    const steps = [];
    if (fixDates) steps.push({ find: '(\\d{1,2})\\. (\\d{1,2})\\. (\\d{4})', change: '$1. $2. $3', label: 'Fixed %n date(s) with thin spaces' });
    if (fixQuotes) {
      // Straight or curly double quotes -> German pairs (opening after space/bracket/paragraph start)
      steps.push({ find: '(?<=[\\s(\\[])[“”\\x{0022}]', change: '„', label: 'Set %n opening quote(s)' });
      steps.push({ find: '^[“”\\x{0022}]', change: '„', label: 'Set %n opening quote(s) at paragraph start' });
      steps.push({ find: '(?<=[^\\s(\\[])[”\\x{0022}]', change: '“', label: 'Set %n closing quote(s)' });
    }
    if (fixDashes) {
      steps.push({ find: '--', change: '—', label: 'Fixed %n double hyphen(s) to em dash' });
      steps.push({ find: ' - ', change: ' – ', label: 'Fixed %n spaced hyphen(s) to en dash' });
    }
    if (fixSpaces) {
      steps.push({ find: ' {2,}', change: ' ', label: 'Fixed %n multiple space(s)' });
      steps.push({ find: ' +$', change: '', label: 'Removed %n trailing space(s)' });
    }

    // GREP find/change keeps character and paragraph formatting (rewriting story.contents would not)
    const script = `
      var doc = __requireDoc();
      ${this.textFrameAcquire(args)}
      var story = textFrame.parentStory;
      var changes = 0, n = 0;
      var changeLog = "=== TYPOGRAPHY FIXES ===\\n";
      ${this.grepSteps(steps)}
      if (changes > 0) {
        changeLog += "\\nTotal fixes applied: " + changes;
      } else {
        changeLog += "No typography issues found.";
      }
      changeLog;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Fix Typography");
  }

  async findTypographyIssues(args) {
    const { frameIndex, pageIndex = 0, useSelectedFrame = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          ${this.textFrameAcquire(args)}
          
          if (textFrame) {
            var content = textFrame.contents;
            var issues = "=== TYPOGRAPHY ANALYSIS ===\\n";
            var problemCount = 0;
            
            // Check for dates with wrong spacing using DATUM search if available
            try {
              var datumQuery = doc.findGrepPreferences.itemByName("DATUM");
              if (datumQuery.isValid) {
                app.findGrepPreferences = NothingEnum.nothing;
                app.findGrepPreferences.findWhat = datumQuery.findWhat;
                var foundDates = textFrame.parentStory.findGrep();
                
                var wrongSpaceDates = 0;
                var correctSpaceDates = 0;
                
                for (var d = 0; d < foundDates.length; d++) {
                  var dateText = foundDates[d].contents;
                  if (dateText.indexOf('\\u2009') === -1 && dateText.match(/\\d{1,2}\\. \\d{1,2}\\. \\d{4}/)) {
                    wrongSpaceDates++;
                  } else if (dateText.indexOf('\\u2009') !== -1) {
                    correctSpaceDates++;
                  }
                }
                
                if (wrongSpaceDates > 0) {
                  issues += "❌ " + wrongSpaceDates + " date(s) with normal spaces (found via DATUM search)\\n";
                  problemCount += wrongSpaceDates;
                }
                if (correctSpaceDates > 0) {
                  issues += "✅ " + correctSpaceDates + " date(s) with correct thin spaces\\n";
                }
                
                app.findGrepPreferences = NothingEnum.nothing;
              } else {
                throw new Error("DATUM search not found");
              }
            } catch (e) {
              // Fallback to manual pattern check
              var wrongDateSpaces = content.match(/\\d{1,2}\\. \\d{1,2}\\. \\d{4}/g);
              if (wrongDateSpaces) {
                issues += "❌ " + wrongDateSpaces.length + " date(s) with normal spaces (fallback pattern)\\n";
                problemCount += wrongDateSpaces.length;
              }
              
              var correctDates = content.match(/\\d{1,2}\\.\\u2009\\d{1,2}\\.\\u2009\\d{4}/g);
              if (correctDates) {
                issues += "✅ " + correctDates.length + " date(s) with correct thin spaces\\n";
              }
            }
            
            // Check for straight quotes
            var straightQuotes = (content.match(/"/g) || []).length;
            if (straightQuotes > 0) {
              issues += "❌ " + straightQuotes + " straight quote(s) found\\n";
              problemCount += straightQuotes;
            }
            
            // Check for double hyphens
            var doubleHyphens = (content.match(/--/g) || []).length;
            if (doubleHyphens > 0) {
              issues += "❌ " + doubleHyphens + " double hyphen(s) (should be em dash)\\n";
              problemCount += doubleHyphens;
            }
            
            // Check for space-hyphen-space
            var spaceHyphens = (content.match(/ - /g) || []).length;
            if (spaceHyphens > 0) {
              issues += "❌ " + spaceHyphens + " space-hyphen-space (should be en dash)\\n";
              problemCount += spaceHyphens;
            }
            
            // Check for multiple spaces
            var multiSpaces = (content.match(/  +/g) || []).length;
            if (multiSpaces > 0) {
              issues += "❌ " + multiSpaces + " multiple space(s) found\\n";
              problemCount += multiSpaces;
            }
            
            // Check for trailing spaces
            var lines = content.split('\\r');
            var trailingSpaces = 0;
            for (var i = 0; i < lines.length; i++) {
              if (lines[i].match(/ +$/)) {
                trailingSpaces++;
              }
            }
            if (trailingSpaces > 0) {
              issues += "❌ " + trailingSpaces + " line(s) with trailing spaces\\n";
              problemCount += trailingSpaces;
            }
            
            issues += "\\n=== SUMMARY ===\\n";
            if (problemCount > 0) {
              issues += "Found " + problemCount + " typography issue(s)\\n";
              issues += "Use fix_typography_in_selection() to auto-correct.";
            } else {
              issues += "No typography issues found. Text is clean!";
            }
            
            issues;
          }
        } catch (e) {
          "Error analyzing typography: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Typography Analysis");
  }

  async listGrepSearches() {
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var result = "=== SAVED GREP SEARCHES ===\\n";
          
          if (doc.findGrepPreferences.length === 0) {
            result += "No saved GREP searches found in this document.\\n";
            result += "TIP: Create and save GREP searches in Find/Change dialog.";
          } else {
            for (var i = 0; i < doc.findGrepPreferences.length; i++) {
              var grepSearch = doc.findGrepPreferences[i];
              result += "Search " + (i + 1) + ": " + (grepSearch.name || "Unnamed") + "\\n";
              result += "  Pattern: " + grepSearch.findWhat + "\\n";
              if (grepSearch.changeTo) {
                result += "  Replace: " + grepSearch.changeTo + "\\n";
              }
              result += "\\n";
            }
            
            // Special note about DATUM search
            try {
              var datumQuery = doc.findGrepPreferences.itemByName("DATUM");
              if (datumQuery.isValid) {
                result += "✅ DATUM search found - will be used for date typography fixes.";
              } else {
                result += "❌ DATUM search not found - typography fixes will use fallback patterns.";
              }
            } catch (e) {
              result += "❌ DATUM search not accessible - using fallback patterns.";
            }
          }
          
          result;
        } catch (e) {
          "Error listing GREP searches: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "GREP Searches");
  }

  async cleanImportedText(args) {
    const { fixParagraphs = true, fixDashes = true, fixLists = true, fixFormatting = true, fixChapterNumbers = true, fixSpaces = true } = args;
    const steps = [];
    if (fixSpaces) {
      steps.push({ find: ' +$', change: '', label: 'Removed %n trailing space(s)' });
      steps.push({ find: ' {2,}', change: ' ', label: 'Fixed %n multiple space(s)' });
    }
    if (fixParagraphs) {
      steps.push({ find: '\\r{2,}', change: '\\r', label: 'Fixed %n double paragraph break(s)' });
      steps.push({ find: '\\n', change: '\\r', label: 'Converted %n line break(s) to paragraphs' });
    }
    if (fixDashes) {
      steps.push({ find: '(\\d+)-(\\d+)', change: '$1–$2', label: 'Fixed %n number range(s) to n-dash' });
      steps.push({ find: ' - ', change: ' – ', label: 'Fixed %n thought dash(es) to n-dash' });
    }
    if (fixLists) {
      steps.push({ find: '^[•·-][ \\t]+', change: '', label: 'Removed %n manual bullet(s)/dash(es)' });
      steps.push({ find: '^\\t[•·-][ \\t]+', change: '', label: 'Removed %n tabbed bullet(s)' });
    }
    if (fixChapterNumbers) {
      steps.push({ find: '(?i)^(Kapitel|Chapter|Teil|Part)[ \\t]+\\d+[.:]*[ \\t]*', change: '', label: 'Removed %n hardcoded chapter number(s)' });
      // Roman numerals only with a dot/colon and a following space, so words like "Ich" or "Vier" stay intact
      steps.push({ find: '^[IVXLC]+[.:][ \\t]+', change: '', label: 'Removed %n roman numeral(s)' });
    }

    const script = `
      var doc = __requireDoc();
      ${this.textFrameAcquire(args)}
      var story = textFrame.parentStory;
      var changes = 0, n = 0;
      var changeLog = "=== TEXT CLEANING REPORT ===\\n";
      ${this.grepSteps(steps)}
      ${fixFormatting ? `
        // Remove local character overrides (manual bold/italic/size); applied styles stay in place
        story.texts[0].clearOverrides(OverrideType.CHARACTER_ONLY);
        changeLog += "✓ Cleared local character formatting overrides\\n";
        changes += 1;
      ` : ''}
      if (changes > 0) {
        changeLog += "\\n=== SUMMARY ===\\n";
        changeLog += "Total changes applied: " + changes + "\\n";
        changeLog += "Next steps: apply paragraph styles and character styles.";
      } else {
        changeLog += "No issues found - text is already clean!";
      }
      changeLog;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Text Cleaning");
  }

  async analyzeTextProblems(args) {
    const { frameIndex, pageIndex = 0, useSelectedFrame = false } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          ${this.textFrameAcquire(args)}
          
          if (textFrame) {
            var content = textFrame.contents;
            var issues = "=== TEXT PROBLEM ANALYSIS ===\\n";
            var problemCount = 0;
            
            // Check trailing spaces
            var trailingSpaces = (content.match(/ +\\r/g) || []).length;
            if (trailingSpaces > 0) {
              issues += "❌ " + trailingSpaces + " line(s) with trailing spaces\\n";
              problemCount += trailingSpaces;
            }
            
            // Check double paragraph breaks
            var doublePars = (content.match(/\\r\\r+/g) || []).length;
            if (doublePars > 0) {
              issues += "❌ " + doublePars + " double paragraph break(s) (fake spacing)\\n";
              problemCount += doublePars;
            }
            
            // Check line breaks instead of paragraphs
            var lineBreaks = (content.match(/\\n/g) || []).length;
            if (lineBreaks > 0) {
              issues += "❌ " + lineBreaks + " line break(s) should be paragraphs\\n";
              problemCount += lineBreaks;
            }
            
            // Check hyphen ranges
            var hyphenRanges = (content.match(/\\d+-\\d+/g) || []).length;
            if (hyphenRanges > 0) {
              issues += "❌ " + hyphenRanges + " number range(s) with hyphen (should be n-dash)\\n";
              problemCount += hyphenRanges;
            }
            
            // Check thought dashes
            var thoughtDashes = (content.match(/ - /g) || []).length;
            if (thoughtDashes > 0) {
              issues += "❌ " + thoughtDashes + " thought dash(es) with hyphen (should be n-dash)\\n";
              problemCount += thoughtDashes;
            }
            
            // Check manual bullets
            var bulletLists = (content.match(/^[•·-]\\s/gm) || []).length;
            var tabBullets = (content.match(/^\\t[•·-]\\s/gm) || []).length;
            if (bulletLists > 0 || tabBullets > 0) {
              issues += "❌ " + (bulletLists + tabBullets) + " manual bullet(s)/dash(es) found\\n";
              problemCount += (bulletLists + tabBullets);
            }
            
            // Check hardcoded chapter numbers
            var chapterNumbers = (content.match(/^(Kapitel|Chapter|Teil|Part)\\s+\\d+/gmi) || []).length;
            var romanNumbers = (content.match(/^[IVX]+[.:]/gm) || []).length;
            if (chapterNumbers > 0 || romanNumbers > 0) {
              issues += "❌ " + (chapterNumbers + romanNumbers) + " hardcoded chapter number(s)\\n";
              problemCount += (chapterNumbers + romanNumbers);
            }
            
            // Check multiple spaces
            var multipleSpaces = (content.match(/  +/g) || []).length;
            if (multipleSpaces > 0) {
              issues += "❌ " + multipleSpaces + " multiple space(s) found\\n";
              problemCount += multipleSpaces;
            }
            
            // Check for manual formatting (rough estimate)
            var story = textFrame.parentStory;
            var hasManualFormatting = false;
            try {
              for (var i = 0; i < Math.min(story.characters.length, 100); i++) {
                var ch = story.characters[i];
                if (ch.fontStyle !== "Regular" && ch.appliedCharacterStyle.name === "[None]") {
                  hasManualFormatting = true;
                  break;
                }
              }
              if (hasManualFormatting) {
                issues += "❌ Manual bold/italic formatting detected (should use character styles)\\n";
                problemCount += 1;
              }
            } catch (e) {}
            
            issues += "\\n=== SUMMARY ===\\n";
            if (problemCount > 0) {
              issues += "Found " + problemCount + " problem(s) in imported text\\n";
              issues += "⚡ Use clean_imported_text() to auto-fix these issues\\n";
              issues += "\\nThis looks like imported Word/text file content that needs cleaning.";
            } else {
              issues += "✅ No major problems found - text looks clean!\\n";
              issues += "Text appears to be properly formatted for InDesign.";
            }
            
            issues;
          }
        } catch (e) {
          "Error analyzing text: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Text Analysis");
  }

  async createTextFrame(args) {
    const {
      content,
      x = 10,
      y = 10,
      width = 100,
      height = 50,
      pageIndex = 0,
      fontSize,
      fontFamily,
      fontStyle = 'Regular',
      textColor,
      alignment,
      paragraphStyle,
      characterStyle,
      lineBreak = 'paragraph'
    } = args;

    if (content === undefined || content === null) {
      throw new Error('content is required');
    }
    const nx = this.jsNum(x, 'x');
    const ny = this.jsNum(y, 'y');
    const nw = this.jsNum(width, 'width');
    const nh = this.jsNum(height, 'height');
    const pageIdx = this.jsNum(pageIndex, 'pageIndex');
    const text = this.toInDesignText(content, lineBreak);

    // Direct formatting would override a paragraph style, so the old defaults are only
    // used when no style is requested. Values the caller passed always apply (and must be valid).
    const explicitFont = !!fontFamily;
    const font = fontFamily || (!paragraphStyle ? 'Helvetica Neue' : undefined);
    const explicitColor = !!textColor;
    const color = textColor || (!paragraphStyle ? 'Black' : undefined);
    const size = fontSize !== undefined ? this.jsNum(fontSize, 'fontSize') : (!paragraphStyle ? 12 : undefined);
    const align = alignment || (!paragraphStyle ? 'LEFT_ALIGN' : undefined);

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${pageIdx});
      var pStyle = ${paragraphStyle ? `__pstyle(doc, ${this.jsStr(paragraphStyle)})` : 'null'};
      var cStyle = ${characterStyle ? `__cstyle(doc, ${this.jsStr(characterStyle)})` : 'null'};
      var notes = [];
      
      var frame = page.textFrames.add();
      try {
        frame.geometricBounds = [${ny}, ${nx}, ${ny + nh}, ${nx + nw}];
        frame.contents = ${this.jsStr(text)};
        var story = frame.parentStory;
        var range = story.texts[0];
        
        // Style first, so explicit formatting can override it
        if (pStyle) range.appliedParagraphStyle = pStyle;
        if (cStyle) range.appliedCharacterStyle = cStyle;
        
        ${font ? `
          try {
            __applyFont(range, ${this.jsStr(font)}, ${this.jsStr(fontStyle)});
          } catch (fe) {
            ${explicitFont ? 'throw fe;' : 'notes.push("default font not available, kept the current font");'}
          }
        ` : ''}
        ${size !== undefined ? `range.pointSize = ${size};` : ''}
        ${color ? `
          try {
            range.fillColor = __swatch(doc, ${this.jsStr(color)});
          } catch (ce) {
            ${explicitColor ? 'throw ce;' : 'notes.push("default colour not available, kept the current colour");'}
          }
        ` : ''}
        ${align ? `range.justification = Justification.${this.jsEnum(align, 'alignment')};` : ''}
      } catch (e) {
        frame.remove();
        throw e;
      }
      
      var b = frame.geometricBounds;
      "Text frame created: id=" + frame.id + " page=" + (${pageIdx} + 1) + " x=" + __r(b[1]) + " y=" + __r(b[0]) + " w=" + __r(b[3] - b[1]) + " h=" + __r(b[2] - b[0]) +
        " paragraphs=" + story.paragraphs.length + " " + __overflowNote(frame) + (notes.length ? " (" + notes.join("; ") + ")" : "");
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Text Frame");
  }

  async editTextFrame(args) {
    const { content, fontSize, fontFamily, fontStyle = 'Regular', textColor, alignment, lineBreak = 'paragraph' } = args;
    const text = content !== undefined ? this.toInDesignText(content, lineBreak) : undefined;

    const script = `
      var doc = __requireDoc();
      var frame = ${this.frameExpr(args)};
      var story = frame.parentStory;
      
      ${text !== undefined ? `
        // Replace the whole story, not just the first paragraph
        var newText = ${this.jsStr(text)};
        story.contents = newText;
        if (story.characters.length > newText.length) {
          story.characters.itemByRange(newText.length, story.characters.length - 1).remove();
        }
      ` : ''}
      var range = story.texts[0];
      ${fontFamily !== undefined ? `__applyFont(range, ${this.jsStr(fontFamily)}, ${this.jsStr(fontStyle)});` : ''}
      ${fontSize !== undefined ? `range.pointSize = ${this.jsNum(fontSize, 'fontSize')};` : ''}
      ${textColor !== undefined ? `range.fillColor = __swatch(doc, ${this.jsStr(textColor)});` : ''}
      ${alignment !== undefined ? `range.justification = Justification.${this.jsEnum(alignment, 'alignment')};` : ''}
      
      "Text frame id=" + frame.id + " updated: paragraphs=" + story.paragraphs.length + " " + __overflowNote(frame);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Edit Text Frame");
  }

  async findReplaceText(args) {
    const { findText, replaceText, caseSensitive = false, wholeWord = false, useGrep = false, scope = 'document' } = args;
    if (!['document', 'story', 'selection'].includes(scope)) throw new Error(`Invalid scope: ${scope}`);

    const script = `
      var doc = __requireDoc();
      var target = doc;
      ${scope === 'story' ? `
        if (app.selection.length === 0 || !app.selection[0].parentStory) throw new Error("scope 'story' needs a selection inside a text story");
        target = app.selection[0].parentStory;
      ` : scope === 'selection' ? `
        if (app.selection.length === 0) throw new Error("scope 'selection' needs something selected");
        target = app.selection[0];
      ` : ''}
      
      app.findTextPreferences = NothingEnum.nothing;
      app.changeTextPreferences = NothingEnum.nothing;
      app.findGrepPreferences = NothingEnum.nothing;
      app.changeGrepPreferences = NothingEnum.nothing;
      var changeCount = 0;
      try {
        ${useGrep ? `
          app.findGrepPreferences.findWhat = ${this.jsStr(findText)};
          app.changeGrepPreferences.changeTo = ${this.jsStr(replaceText)};
          changeCount = target.changeGrep().length;
        ` : `
          app.findTextPreferences.findWhat = ${this.jsStr(findText)};
          app.changeTextPreferences.changeTo = ${this.jsStr(replaceText)};
          app.findChangeTextOptions.caseSensitive = ${caseSensitive ? 'true' : 'false'};
          app.findChangeTextOptions.wholeWord = ${wholeWord ? 'true' : 'false'};
          changeCount = target.changeText().length;
        `}
      } finally {
        app.findTextPreferences = NothingEnum.nothing;
        app.changeTextPreferences = NothingEnum.nothing;
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
      }
      "Replaced " + changeCount + " instance(s) of " + ${this.jsStr(findText)} + " with " + ${this.jsStr(replaceText)} + " (scope: ${scope})";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Find/Replace Text");
  }

  async placeImage(args) {
    const {
      imagePath, x = 10, y = 10, width, height, pageIndex = 0,
      fitOption = 'PROPORTIONALLY', createFrame = true,
      contentOffsetX, contentOffsetY, contentScale
    } = args;

    // Security: Validate image path
    const validatedPath = this.validateFilePath(imagePath);
    const fits = {
      FILL_PROPORTIONALLY: 'FitOptions.FILL_PROPORTIONALLY',
      PROPORTIONALLY: 'FitOptions.PROPORTIONALLY',
      CONTENT_TO_FRAME: 'FitOptions.CONTENT_TO_FRAME',
      FRAME_TO_CONTENT: 'FitOptions.FRAME_TO_CONTENT',
      CENTER_CONTENT: 'FitOptions.CENTER_CONTENT',
      NONE: null,
    };
    if (!(fitOption in fits)) {
      throw new Error(`Invalid fitOption: ${fitOption} (use ${Object.keys(fits).join(', ')})`);
    }
    const nx = this.jsNum(x, 'x');
    const ny = this.jsNum(y, 'y');
    const pageIdx = this.jsNum(pageIndex, 'pageIndex');
    const w = width !== undefined ? this.jsNum(width, 'width') : null;
    const h = height !== undefined ? this.jsNum(height, 'height') : null;

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${pageIdx});
      var imageFile = File(${this.jsStr(validatedPath)});
      if (!imageFile.exists) throw new Error("Image file not found: " + imageFile.fsName);
      
      var rect;
      ${createFrame ? `
        rect = page.rectangles.add();
        var fw = ${w !== null ? w : (h !== null ? h : 50)}, fh = ${h !== null ? h : (w !== null ? w : 50)};
        rect.geometricBounds = [${ny}, ${nx}, ${ny} + fh, ${nx} + fw];
        try {
          rect.place(imageFile);
        } catch (pe) {
          rect.remove();
          throw pe;
        }
        ${(w !== null) !== (h !== null) || (w === null && h === null) ? `
          // Only one dimension (or none) given: size the frame from the image aspect ratio
          rect.fit(FitOptions.FRAME_TO_CONTENT);
          var nb = rect.geometricBounds, iw = nb[3] - nb[1], ih = nb[2] - nb[0];
          ${w !== null ? `
            rect.geometricBounds = [${ny}, ${nx}, ${ny} + ${w} * ih / iw, ${nx} + ${w}];
            rect.fit(FitOptions.CONTENT_TO_FRAME);
          ` : h !== null ? `
            rect.geometricBounds = [${ny}, ${nx}, ${ny} + ${h}, ${nx} + ${h} * iw / ih];
            rect.fit(FitOptions.CONTENT_TO_FRAME);
          ` : `
            rect.geometricBounds = [${ny}, ${nx}, ${ny} + ih, ${nx} + iw];
            rect.fit(FitOptions.CONTENT_TO_FRAME);   // move the content with the frame
          `}
        ` : `
          ${fits[fitOption] ? `rect.fit(${fits[fitOption]});` : ''}
        `}
      ` : `
        var placed = page.place(imageFile, [${nx}, ${ny}]);
        rect = placed[0].parent;
      `}
      
      ${contentScale !== undefined ? `
        rect.graphics[0].horizontalScale = ${this.jsNum(contentScale, 'contentScale')};
        rect.graphics[0].verticalScale = ${this.jsNum(contentScale, 'contentScale')};
      ` : ''}
      ${contentOffsetX !== undefined || contentOffsetY !== undefined ? `
        rect.graphics[0].move(undefined, [${contentOffsetX !== undefined ? this.jsNum(contentOffsetX, 'contentOffsetX') : 0}, ${contentOffsetY !== undefined ? this.jsNum(contentOffsetY, 'contentOffsetY') : 0}]);
      ` : ''}
      
      "Image placed: " + imageFile.name + " - " + __describe(rect);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Place Image");
  }

  async createRectangle(args) {
    const { x, y, width, height, pageIndex = 0, fillColor, strokeColor, strokeWidth = 1, cornerRadius = 0 } = args;
    return this._createShape('rectangles', 'Rectangle', args);
  }

  async createEllipse(args) {
    return this._createShape('ovals', 'Ellipse', args);
  }

  async _createShape(collection, label, args) {
    const { x, y, width, height, pageIndex = 0, fillColor, strokeColor, strokeWidth = 1, cornerRadius = 0 } = args;
    const nx = this.jsNum(x, 'x');
    const ny = this.jsNum(y, 'y');
    const nw = this.jsNum(width, 'width');
    const nh = this.jsNum(height, 'height');
    const pageIdx = this.jsNum(pageIndex, 'pageIndex');
    const radius = collection === 'rectangles' ? this.jsNum(cornerRadius, 'cornerRadius') : 0;

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${pageIdx});
      ${fillColor ? `var fill = __swatch(doc, ${this.jsStr(fillColor)});` : ''}
      ${strokeColor ? `var stroke = __swatch(doc, ${this.jsStr(strokeColor)});` : ''}
      var shape = page.${collection}.add();
      shape.geometricBounds = [${ny}, ${nx}, ${ny + nh}, ${nx + nw}];
      ${radius > 0 ? `shape.topLeftCornerRadius = shape.topRightCornerRadius = shape.bottomLeftCornerRadius = shape.bottomRightCornerRadius = ${radius}; shape.topLeftCornerOption = shape.topRightCornerOption = shape.bottomLeftCornerOption = shape.bottomRightCornerOption = CornerOptions.ROUNDED_CORNER;` : ''}
      ${fillColor ? 'shape.fillColor = fill;' : ''}
      ${strokeColor ? `shape.strokeColor = stroke; shape.strokeWeight = "${this.jsNum(strokeWidth, 'strokeWidth')}pt";` : ''}
      "${label} created: " + __describe(shape);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, `Create ${label}`);
  }

  // =================== STYLE MANAGEMENT ===================
  // Shared by create_paragraph_style and modify_paragraph_style
  paragraphStyleScript(mode, args) {
    const name = mode === 'create' ? args.name : args.styleName;
    if (!name) throw new Error(mode === 'create' ? 'name is required' : 'styleName is required');
    const a = args;
    const has = (v) => v !== undefined && v !== null;
    const lines = [];

    if (has(a.baseStyle)) lines.push(`pStyle.basedOn = __pstyle(doc, ${this.jsStr(a.baseStyle)});`);
    if (has(a.fontFamily)) {
      lines.push(`__applyFont(pStyle, ${this.jsStr(a.fontFamily)}, ${this.jsStr(a.fontStyle || 'Regular')});`);
    } else if (has(a.fontStyle)) {
      lines.push(`pStyle.fontStyle = ${this.jsStr(a.fontStyle)};`);
    }
    if (has(a.fontSize)) lines.push(`pStyle.pointSize = ${this.jsNum(a.fontSize, 'fontSize')};`);
    if (has(a.leading)) lines.push(`pStyle.leading = ${this.jsNum(a.leading, 'leading')};`);
    if (has(a.tracking)) lines.push(`pStyle.tracking = ${this.jsNum(a.tracking, 'tracking')};`);
    if (has(a.kerning)) {
      if (!['Metrics', 'Optical'].includes(a.kerning)) throw new Error(`Invalid kerning: ${a.kerning} (use Metrics or Optical)`);
      lines.push(`pStyle.kerningMethod = ${this.jsStr(a.kerning)};`);
    }
    if (has(a.spaceBefore)) lines.push(`pStyle.spaceBefore = "${this.jsNum(a.spaceBefore, 'spaceBefore')}mm";`);
    if (has(a.spaceBeforePt)) lines.push(`pStyle.spaceBefore = "${this.jsNum(a.spaceBeforePt, 'spaceBeforePt')}pt";`);
    if (has(a.spaceAfter)) lines.push(`pStyle.spaceAfter = "${this.jsNum(a.spaceAfter, 'spaceAfter')}mm";`);
    if (has(a.spaceAfterPt)) lines.push(`pStyle.spaceAfter = "${this.jsNum(a.spaceAfterPt, 'spaceAfterPt')}pt";`);
    if (has(a.firstLineIndent)) lines.push(`pStyle.firstLineIndent = "${this.jsNum(a.firstLineIndent, 'firstLineIndent')}mm";`);
    if (has(a.leftIndent)) lines.push(`pStyle.leftIndent = "${this.jsNum(a.leftIndent, 'leftIndent')}mm";`);
    if (has(a.rightIndent)) lines.push(`pStyle.rightIndent = "${this.jsNum(a.rightIndent, 'rightIndent')}mm";`);
    if (has(a.alignment)) lines.push(`pStyle.justification = Justification.${this.jsEnum(a.alignment, 'alignment')};`);
    if (has(a.hyphenation)) lines.push(`pStyle.hyphenation = ${a.hyphenation ? 'true' : 'false'};`);
    if (has(a.keepWithNext)) lines.push(`pStyle.keepWithNext = ${this.jsNum(a.keepWithNext, 'keepWithNext')};`);
    if (has(a.keepLinesTogether)) lines.push(`pStyle.keepLinesTogether = ${a.keepLinesTogether ? 'true' : 'false'};`);
    if (has(a.keepFirstLines)) lines.push(`pStyle.keepFirstLines = ${this.jsNum(a.keepFirstLines, 'keepFirstLines')};`);
    if (has(a.keepLastLines)) lines.push(`pStyle.keepLastLines = ${this.jsNum(a.keepLastLines, 'keepLastLines')};`);
    if (has(a.textColor)) lines.push(`pStyle.fillColor = __swatch(doc, ${this.jsStr(a.textColor)});`);
    if (has(a.keepWithPrevious)) lines.push(`pStyle.keepWithPrevious = ${a.keepWithPrevious ? 'true' : 'false'};`);
    if (has(a.composer)) {
      const composers = { paragraph: 'Adobe Paragraph Composer', 'single-line': 'Adobe Single-line Composer' };
      if (!composers[a.composer]) throw new Error(`Invalid composer: ${a.composer} (use paragraph or single-line)`);
      lines.push(`pStyle.composer = ${this.jsStr(composers[a.composer])};`);
    }
    if (has(a.capitalization)) {
      const caps = { normal: 'NORMAL', small_caps: 'SMALL_CAPS', all_caps: 'ALL_CAPS', cap_to_small_cap: 'CAP_TO_SMALL_CAP', lower_case: 'LOWER_CASE' };
      if (!caps[a.capitalization]) throw new Error(`Invalid capitalization: ${a.capitalization} (use ${Object.keys(caps).join(', ')})`);
      lines.push(`pStyle.capitalization = Capitalization.${caps[a.capitalization]};`);
    }
    // Hyphenation details (hyphenation itself is on/off above)
    for (const [key, prop] of [['hyphenateWordsLongerThan', 'hyphenateWordsLongerThan'], ['hyphenateAfterFirst', 'hyphenateAfterFirst'], ['hyphenateBeforeLast', 'hyphenateBeforeLast'], ['hyphenateLadderLimit', 'hyphenateLadderLimit']]) {
      if (has(a[key])) lines.push(`pStyle.${prop} = ${this.jsNum(a[key], key)};`);
    }
    for (const key of ['hyphenateCapitalizedWords', 'hyphenateLastWord']) {
      if (has(a[key])) lines.push(`pStyle.${key} = ${a[key] ? 'true' : 'false'};`);
    }
    // Typography: ligatures and OpenType features
    if (has(a.ligatures)) lines.push(`pStyle.ligatures = ${a.ligatures ? 'true' : 'false'};`);
    if (has(a.openType)) {
      const flags = { discretionaryLigatures: 'otfDiscretionaryLigature', contextualAlternates: 'otfContextualAlternate', swash: 'otfSwash', titling: 'otfTitling', fractions: 'otfFraction', ordinals: 'otfOrdinal', slashedZero: 'otfSlashedZero', historical: 'otfHistorical' };
      const figures = { default: 'DEFAULT_VALUE', tabular_lining: 'TABULAR_LINING', proportional_oldstyle: 'PROPORTIONAL_OLDSTYLE', proportional_lining: 'PROPORTIONAL_LINING', tabular_oldstyle: 'TABULAR_OLDSTYLE' };
      for (const [key, value] of Object.entries(a.openType)) {
        if (flags[key]) {
          if (typeof value !== 'boolean') throw new Error(`openType.${key} must be true or false`);
          lines.push(`pStyle.${flags[key]} = ${value};`);
        } else if (key === 'figureStyle') {
          if (!figures[value]) throw new Error(`Invalid openType.figureStyle: ${value} (use ${Object.keys(figures).join(', ')})`);
          lines.push(`pStyle.otfFigureStyle = OTFFigureStyle.${figures[value]};`);
        } else if (key === 'stylisticSets') {
          lines.push(`pStyle.otfStylisticSets = ${this.jsNum(value, 'openType.stylisticSets')};`);
        } else {
          throw new Error(`Unknown openType feature: ${key} (use ${[...Object.keys(flags), 'figureStyle', 'stylisticSets'].join(', ')})`);
        }
      }
    }

    return `
      var doc = __requireDoc();
      var pStyle;
      ${mode === 'create' ? `
        if (doc.paragraphStyles.itemByName(${this.jsStr(name)}).isValid) {
          throw new Error("Paragraph style already exists: " + ${this.jsStr(name)} + " (use modify_paragraph_style)");
        }
        pStyle = doc.paragraphStyles.add();
        pStyle.name = ${this.jsStr(name)};
      ` : `
        pStyle = __pstyle(doc, ${this.jsStr(name)});
      `}
      try {
        ${lines.join('\n        ')}
      } catch (e) {
        ${mode === 'create' ? 'pStyle.remove();' : ''}
        throw e;
      }
      var fontName = "n/a";
      try { fontName = pStyle.appliedFont.name.replace("\\t", " "); } catch (e) { try { fontName = String(pStyle.appliedFont); } catch (e2) {} }
      "Paragraph style " + ${this.jsStr(name)} + " ${mode === 'create' ? 'created' : 'modified'}. font=" + fontName +
        " fontStyle=" + pStyle.fontStyle + " size=" + __r(pStyle.pointSize) + "pt leading=" + (pStyle.leading === Leading.AUTO ? "auto" : __r(pStyle.leading) + "pt") +
        " tracking=" + pStyle.tracking + " spaceBefore=" + __r(pStyle.spaceBefore) + "mm spaceAfter=" + __r(pStyle.spaceAfter) + "mm" +
        " firstLineIndent=" + __r(pStyle.firstLineIndent) + "mm leftIndent=" + __r(pStyle.leftIndent) + "mm hyphenation=" + pStyle.hyphenation +
        " composer=" + (function () { try { return pStyle.composer; } catch (e) { return "n/a"; } })() +
        " capitalization=" + (function () { try { return String(pStyle.capitalization); } catch (e) { return "n/a"; } })() +
        " ligatures=" + (function () { try { return pStyle.ligatures; } catch (e) { return "n/a"; } })();
    `;
  }

  async createParagraphStyle(args) {
    const result = await this.executeInDesignScript(this.paragraphStyleScript('create', args));
    return this.formatResponse(result, "Create Paragraph Style");
  }

  async modifyParagraphStyle(args) {
    const result = await this.executeInDesignScript(this.paragraphStyleScript('modify', args));
    return this.formatResponse(result, "Modify Paragraph Style");
  }

  // Shared by create_/modify_ character and object styles. kind: 'character' | 'object'
  simpleStyleScript(kind, mode, args) {
    const name = mode === 'create' ? args.name : args.styleName;
    if (!name) throw new Error(mode === 'create' ? 'name is required' : 'styleName is required');
    const a = args;
    const has = (v) => v !== undefined && v !== null;
    const coll = kind === 'character' ? 'characterStyles' : 'objectStyles';
    const label = kind === 'character' ? 'Character' : 'Object';
    const finder = kind === 'character' ? '__cstyle' : '__ostyle';
    const lines = [];

    if (has(a.baseStyle)) lines.push(`style.basedOn = ${finder}(doc, ${this.jsStr(a.baseStyle)});`);
    if (kind === 'character') {
      if (has(a.fontFamily)) lines.push(`__applyFont(style, ${this.jsStr(a.fontFamily)}, ${this.jsStr(a.fontStyle || 'Regular')});`);
      else if (has(a.fontStyle)) lines.push(`style.fontStyle = ${this.jsStr(a.fontStyle)};`);
      if (has(a.fontSize)) lines.push(`style.pointSize = ${this.jsNum(a.fontSize, 'fontSize')};`);
      if (has(a.tracking)) lines.push(`style.tracking = ${this.jsNum(a.tracking, 'tracking')};`);
      if (has(a.textColor)) lines.push(`style.fillColor = __swatch(doc, ${this.jsStr(a.textColor)});`);
    } else {
      if (has(a.fillColor)) lines.push(`style.fillColor = __swatch(doc, ${this.jsStr(a.fillColor)});`);
      if (has(a.strokeColor)) lines.push(`style.strokeColor = __swatch(doc, ${this.jsStr(a.strokeColor)});`);
      if (has(a.strokeWidth)) lines.push(`style.strokeWeight = ${this.jsNum(a.strokeWidth, 'strokeWidth')};`);
      if (has(a.transparency)) {
        const t = this.jsNum(a.transparency, 'transparency');
        if (t < 0 || t > 100) throw new Error('transparency must be 0-100');
        lines.push(`style.transparencySettings.blendingSettings.opacity = ${100 - t};`);
      }
    }
    if (mode === 'modify' && lines.length === 0) throw new Error(`No properties given to modify for ${kind} style ${name}`);

    return `
      var doc = __requireDoc();
      var style;
      ${mode === 'create' ? `
        if (doc.${coll}.itemByName(${this.jsStr(name)}).isValid) {
          throw new Error("${label} style already exists: " + ${this.jsStr(name)} + " (use modify_${kind}_style)");
        }
        style = doc.${coll}.add();
        style.name = ${this.jsStr(name)};
      ` : `
        style = ${finder}(doc, ${this.jsStr(name)});
      `}
      try {
        ${lines.join('\n        ')}
      } catch (e) {
        ${mode === 'create' ? 'style.remove();' : ''}
        throw e;
      }
      "${label} style " + ${this.jsStr(name)} + " ${mode === 'create' ? 'created' : 'modified'}." ${kind === 'character' ? `+ " fontStyle=" + style.fontStyle + " size=" + __r(style.pointSize) + "pt tracking=" + style.tracking` : `+ " stroke=" + __r(style.strokeWeight) + "pt opacity=" + __r(style.transparencySettings.blendingSettings.opacity) + "%"`};
    `;
  }

  async createCharacterStyle(args) {
    const result = await this.executeInDesignScript(this.simpleStyleScript('character', 'create', args));
    return this.formatResponse(result, "Create Character Style");
  }

  async modifyCharacterStyle(args) {
    const result = await this.executeInDesignScript(this.simpleStyleScript('character', 'modify', args));
    return this.formatResponse(result, "Modify Character Style");
  }

  async createObjectStyle(args) {
    const result = await this.executeInDesignScript(this.simpleStyleScript('object', 'create', args));
    return this.formatResponse(result, "Create Object Style");
  }

  async modifyObjectStyle(args) {
    const result = await this.executeInDesignScript(this.simpleStyleScript('object', 'modify', args));
    return this.formatResponse(result, "Modify Object Style");
  }

  async applyObjectStyle(args) {
    const { styleName } = args;
    const explicit = args.objectId !== undefined || args.objectIndex !== undefined;
    const num = (v, label) => (v === undefined || v === null ? 'null' : this.jsNum(v, label));

    const script = `
      var doc = __requireDoc();
      var oStyle = __ostyle(doc, ${this.jsStr(styleName)});
      var objectsToStyle = [];
      ${explicit ? `
        // An explicit object (id, or pageIndex + objectIndex) wins over the selection
        objectsToStyle.push(__resolve(doc, ${num(args.objectId, 'objectId')}, ${num(args.pageIndex, 'pageIndex')}, ${num(args.objectIndex, 'objectIndex')}));
      ` : `
        for (var i = 0; i < app.selection.length; i++) objectsToStyle.push(app.selection[i]);
        if (objectsToStyle.length === 0) {
          throw new Error("No objects selected. Select objects, or pass objectId / objectIndex.");
        }
      `}
      var appliedCount = 0;
      for (var j = 0; j < objectsToStyle.length; j++) {
        try {
          objectsToStyle[j].appliedObjectStyle = oStyle;
          appliedCount++;
        } catch (e) {}
      }
      if (appliedCount === 0) throw new Error("The style could not be applied to the given object(s)");
      "Object style " + ${this.jsStr(styleName)} + " applied to " + appliedCount + " object(s)";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Apply Object Style");
  }

  async applyParagraphStyle(args) {
    const { styleName, startIndex, endIndex } = args;
    const range = startIndex !== undefined && endIndex !== undefined;

    const script = `
      var doc = __requireDoc();
      var frame = ${this.frameExpr(args)};
      var style = __pstyle(doc, ${this.jsStr(styleName)});
      var story = frame.parentStory;
      ${range ? `
        story.characters.itemByRange(${this.jsNum(startIndex, 'startIndex')}, ${this.jsNum(endIndex, 'endIndex')}).paragraphs.everyItem().appliedParagraphStyle = style;
      ` : `
        story.paragraphs.everyItem().appliedParagraphStyle = style;
      `}
      "Paragraph style " + ${this.jsStr(styleName)} + " applied to text frame id=" + frame.id + " " + __overflowNote(frame);
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
  // Naive RGB (0-255) -> CMYK (0-100) WITHOUT a colour profile. Hues drift (greens toward blue), so it is only
  // used when conversion: "simple" is requested; the default is InDesign's colour management (see __cmsCmyk).
  rgbToCmyk(r, g, b) {
    const rr = r / 255, gg = g / 255, bb = b / 255;
    const k = 1 - Math.max(rr, gg, bb);
    if (k >= 1) return [0, 0, 0, 100];
    const c = (1 - rr - k) / (1 - k);
    const m = (1 - gg - k) / (1 - k);
    const y = (1 - bb - k) / (1 - k);
    return [c, m, y, k].map((v) => Math.round(v * 1000) / 10);
  }

  parseHex(hex) {
    const m = String(hex).trim().match(/^#?([0-9a-f]{6})$/i);
    if (!m) throw new Error(`Invalid hex colour: ${hex} (expected #RRGGBB)`);
    const n = parseInt(m[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }

  async createColorSwatch(args) {
    const { name, colorModel = 'CMYK', colorValues, hex, preset, keepRgb = false, spotColor = false, update = false, conversion = 'profile' } = args;
    if (!name) throw new Error('name is required');
    if (!['profile', 'simple'].includes(conversion)) throw new Error(`Invalid conversion: ${conversion} (use profile or simple)`);

    const presets = { rich_black: [60, 40, 40, 100], overprint_black: [0, 0, 0, 100] };
    let model = colorModel;
    let values = colorValues;
    if (preset) {
      if (!presets[preset]) throw new Error(`Unknown preset: ${preset} (use ${Object.keys(presets).join(', ')})`);
      model = 'CMYK';
      values = presets[preset];
    } else if (hex !== undefined) {
      model = 'RGB';
      values = this.parseHex(hex);
    }
    if (!['CMYK', 'RGB'].includes(model)) throw new Error(`Invalid colorModel: ${model} (use CMYK or RGB)`);
    if (!Array.isArray(values) || values.length !== (model === 'CMYK' ? 4 : 3) || values.some((v) => !Number.isFinite(Number(v)))) {
      throw new Error(`colorValues must be ${model === 'CMYK' ? '[C,M,Y,K] (0-100)' : '[R,G,B] (0-255)'}`);
    }
    values = values.map(Number);

    const script = `
      var doc = __requireDoc();
      var existing = doc.colors.itemByName(${this.jsStr(name)});
      if (doc.swatches.itemByName(${this.jsStr(name)}).isValid && !(${update ? 'true' : 'false'} && existing.isValid)) {
        throw new Error("Swatch already exists: " + ${this.jsStr(name)} + " (pass update: true to change its values, or use update_color_swatch)");
      }
      var model = ${this.jsStr(model)};
      var values = ${this.jsJson(values)};
      var note = "";
      // Print documents get CMYK unless keepRgb is set
      if (model === "RGB" && ${keepRgb ? 'false' : 'true'} && String(doc.documentPreferences.intent).indexOf("PRINT") === 0) {
        ${conversion === 'simple'
          ? `var conv = ${this.jsJson(model === 'RGB' ? this.rgbToCmyk(...values) : [0, 0, 0, 0])};
        note = " (RGB converted to CMYK with the simple formula, no colour profile - hues shift, prefer conversion profile)";`
          : `var conv = __cmsCmyk(doc, values);   // InDesign's colour engine: sRGB -> the document's CMYK profile
        note = " (RGB converted to CMYK with the document profile " + doc.cmykProfile + ")";`}
        values = conv; model = "CMYK";
      }
      var updated = existing.isValid;
      var color;
      if (updated) {
        color = existing;
      } else {
        color = doc.colors.add();
        color.name = ${this.jsStr(name)};
      }
      color.model = ${spotColor ? 'ColorModel.SPOT' : 'ColorModel.PROCESS'};
      color.space = (model === "CMYK") ? ColorSpace.CMYK : ColorSpace.RGB;
      color.colorValue = values;
      var out = [];
      for (var i = 0; i < color.colorValue.length; i++) out.push(__r(color.colorValue[i]));
      "Swatch " + ${this.jsStr(name)} + (updated ? " updated: " : " created: ") + model + " " + out.join(",") + (${spotColor ? 'true' : 'false'} ? " (spot)" : "") + note;
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Color Swatch");
  }

  async applyColor(args) {
    const { swatchName, property = 'fill' } = args;
    if (!['fill', 'stroke'].includes(property)) throw new Error(`Invalid property: ${property} (use fill or stroke)`);
    const num = (v, label) => (v === undefined || v === null ? 'null' : this.jsNum(v, label));

    const script = `
      var doc = __requireDoc();
      var item = __resolve(doc, ${num(args.objectId, 'objectId')}, ${num(args.pageIndex, 'pageIndex')}, ${num(args.objectIndex, 'objectIndex')});
      var swatch = __swatch(doc, ${this.jsStr(swatchName)});
      item.${property === 'fill' ? 'fillColor' : 'strokeColor'} = swatch;
      "Colour " + ${this.jsStr(swatchName)} + " applied to ${property} of " + __describe(item);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Apply Color");
  }

  // =================== EXPORT FUNCTIONS ===================
  // "1-3,5" / "all" -> array of 1-based page numbers, or null for all pages
  parsePageRange(pageRange) {
    if (pageRange === undefined || pageRange === null || String(pageRange).trim().toLowerCase() === 'all') return null;
    const pages = [];
    for (const part of String(pageRange).split(',')) {
      const m = part.trim().match(/^(\d+)(?:\s*-\s*(\d+))?$/);
      if (!m) throw new Error(`Invalid pageRange: ${pageRange} (use e.g. "1", "2-3", "1,3-5" or "all")`);
      const from = Number(m[1]);
      const to = m[2] ? Number(m[2]) : from;
      if (from < 1 || to < from) throw new Error(`Invalid pageRange: ${pageRange}`);
      for (let i = from; i <= to; i++) pages.push(i);
    }
    return pages;
  }

  // [1,2,3,5] -> "+1-+3,+5"
  absoluteRange(positions) {
    const parts = [];
    let start = positions[0], prev = positions[0];
    for (const n of [...positions.slice(1), null]) {
      if (n !== null && n === prev + 1) { prev = n; continue; }
      parts.push(start === prev ? `+${start}` : `+${start}-+${prev}`);
      start = prev = n;
    }
    return parts.join(',');
  }

  async exportPDF(args) {
    const { filePath, preset = 'HighQualityPrint', pageRange = 'all', includeBleed = false, includeSlug = false } = args;

    // Security: Require confirmation for file export
    this.validateDestructiveOperation(args, 'EXPORT PDF', filePath);

    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);
    this.parsePageRange(pageRange); // validate early

    const aliases = {
      Print: '[High Quality Print]',
      HighQualityPrint: '[High Quality Print]',
      PressQuality: '[Press Quality]',
      Web: '[Smallest File Size]',
      SmallestFileSize: '[Smallest File Size]',
      PDFX1a: '[PDF/X-1a:2001]',
      PDFX3: '[PDF/X-3:2002]',
      PDFX4: '[PDF/X-4:2008]',
    };
    const presetName = aliases[preset] || preset;
    // Page numbers are positions in the document (1 = first page); they are sent to InDesign as absolute
    // page numbers ("+2-+3"), so custom numbering, sections and repeated page names cannot change the meaning.
    const positions = this.parsePageRange(pageRange);
    const rangeString = positions === null ? '' : this.absoluteRange(positions);

    const script = `
      var doc = __requireDoc();
      var base = app.pdfExportPresets.itemByName(${this.jsStr(presetName)});
      if (!base.isValid) {
        var names = [];
        for (var i = 0; i < app.pdfExportPresets.length; i++) names.push(app.pdfExportPresets[i].name);
        throw new Error("PDF preset not found: " + ${this.jsStr(presetName)} + ". Available: " + names.join(", "));
      }
      
      // A preset overrides the export preferences (including bleed and slug), so export with a
      // temporary copy of it that carries the bleed/slug settings. Page range is a preference.
      var TMP = "__indesign_mcp_tmp__";
      var old = app.pdfExportPresets.itemByName(TMP);
      if (old.isValid) old.remove();
      var tmp = app.pdfExportPresets.add({ name: TMP });
      var pdfFile = File(${this.jsStr(validatedPath)});
      if (!pdfFile.parent.exists) pdfFile.parent.create();
      try {
        tmp.properties = base.properties;
        tmp.name = TMP;
        tmp.useDocumentBleedWithPDF = ${includeBleed ? 'true' : 'false'};
        tmp.includeSlugWithPDF = ${includeSlug ? 'true' : 'false'};
        ${rangeString ? `app.pdfExportPreferences.pageRange = ${this.jsStr(rangeString)};` : 'app.pdfExportPreferences.pageRange = PageRange.ALL_PAGES;'}
        doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, tmp);
      } finally {
        try { tmp.remove(); } catch (e) {}
      }
      if (!pdfFile.exists) throw new Error("PDF export did not create the file");
      "PDF exported: " + pdfFile.fsName + " (requested preset " + ${this.jsStr(preset)} + ", effective preset " + base.name + ", pages " + ${this.jsStr(positions === null ? 'all' : String(pageRange).replace(/\s+/g, ''))} + ", bleed " + ${includeBleed ? '"included"' : '"not included"'} + ", slug " + ${includeSlug ? '"included"' : '"not included"'} + ")";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export PDF");
  }

  async exportImages(args) {
    const { folderPath, format = 'PNG', resolution = 300, pageRange = 'all', includeBleed = false } = args;

    // Security: Require confirmation for folder write
    this.validateDestructiveOperation(args, 'EXPORT IMAGES', folderPath);

    // Security: Validate folder path
    const validatedPath = this.validateFilePath(folderPath);
    const fmt = String(format).toUpperCase();
    if (!['PNG', 'JPEG', 'JPG'].includes(fmt)) throw new Error(`Invalid format: ${format} (use PNG or JPEG)`);
    const dpi = this.jsNum(resolution, 'resolution');
    if (dpi < 1 || dpi > 2400) throw new Error(`Invalid resolution: ${resolution} (1-2400 dpi)`);
    const pages = this.parsePageRange(pageRange);

    const script = `
      var doc = __requireDoc();
      var exportFolder = Folder(${this.jsStr(validatedPath)});
      if (!exportFolder.exists && !exportFolder.create()) throw new Error("Cannot create folder: " + exportFolder.fsName);
      
      var isPng = ${fmt === 'PNG' ? 'true' : 'false'};
      var prefs = isPng ? app.pngExportPreferences : app.jpegExportPreferences;
      prefs.exportResolution = ${dpi};
      prefs.useDocumentBleeds = ${includeBleed ? 'true' : 'false'};
      // pageString is ignored (all pages are exported every time) unless the range mode is EXPORT_RANGE
      prefs.exportingSpread = false;
      if (isPng) prefs.pngExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
      else prefs.jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
      var ext = isPng ? ".png" : ".jpg";
      var base = doc.name.replace(/\\.indd$/i, "");
      
      var wanted = ${pages ? this.jsJson(pages) : 'null'};
      if (wanted === null) {
        wanted = [];
        for (var i = 0; i < doc.pages.length; i++) wanted.push(i + 1);
      }
      var written = [];
      for (var k = 0; k < wanted.length; k++) {
        var n = wanted[k];
        if (n > doc.pages.length) throw new Error("Page " + n + " does not exist (document has " + doc.pages.length + " pages)");
        // "+N" = the Nth page of the document, whatever the page names are (numbering start, sections, repeated names)
        prefs.pageString = "+" + n;
        var out = File(exportFolder.fsName + "/" + base + "_page" + n + ext);
        doc.exportFile(isPng ? ExportFormat.PNG_FORMAT : ExportFormat.JPG, out, false);
        if (!out.exists) throw new Error("Export did not create " + out.fsName);
        written.push(out.fsName);
      }
      "Exported " + written.length + " page(s) as ${fmt === 'PNG' ? 'PNG' : 'JPEG'} at ${dpi} dpi (pages: " + wanted.join(", ") + "):\\n" + written.join("\\n");
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export Images");
  }

  async exportEPUB(args) {
    const { filePath, version = 'EPUB3', imageFormat = 'AUTOMATIC' } = args;

    // Security: Require confirmation for file export
    this.validateDestructiveOperation(args, 'EXPORT EPUB', filePath);

    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);
    if (!['EPUB2', 'EPUB3'].includes(version)) throw new Error(`Invalid version: ${version}`);
    if (!['AUTOMATIC', 'PNG', 'JPEG', 'GIF'].includes(imageFormat)) throw new Error(`Invalid imageFormat: ${imageFormat}`);

    const script = `
      var doc = __requireDoc();
      var epubFile = File(${this.jsStr(validatedPath)});
      if (!epubFile.parent.exists) epubFile.parent.create();
      var prefs = doc.epubExportPreferences;
      prefs.version = EpubVersion.${version};
      prefs.imageConversion = ImageConversion.${imageFormat};
      doc.exportFile(ExportFormat.EPUB, epubFile, false);
      if (!epubFile.exists) throw new Error("EPUB export did not create the file");
      "EPUB exported: " + epubFile.fsName + " (" + prefs.version + ", images: ${imageFormat})";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Export EPUB");
  }

  async packageDocument(args) {
    const { folderPath, includeLinkedFiles = true, includeFonts = true, createReport = true } = args;

    // Security: Require confirmation for package creation
    this.validateDestructiveOperation(args, 'PACKAGE DOCUMENT', folderPath);

    // Security: Validate folder path
    const validatedPath = this.validateFilePath(folderPath);

    const script = `
      var doc = __requireDoc();
      if (!__hasFile(doc)) throw new Error("Document " + doc.name + " has never been saved. Save it first (save_document with a filePath).");
      var packageFolder = Folder(${this.jsStr(validatedPath)});
      if (!packageFolder.exists && !packageFolder.create()) throw new Error("Cannot create folder: " + packageFolder.fsName);
      // packageForPrint(to, copyingFonts, copyingLinkedGraphics, copyingProfiles, updatingGraphics,
      //                 includingHiddenLayers, ignorePreflightErrors, creatingReport, includeIdml, includePdf,
      //                 versionComments, forceSave)
      doc.packageForPrint(packageFolder, ${includeFonts ? 'true' : 'false'}, ${includeLinkedFiles ? 'true' : 'false'}, true, false, false, true, ${createReport ? 'true' : 'false'}, false, false, "Package created by InDesign MCP Server", false);
      "Document packaged to: " + packageFolder.fsName + " (" + packageFolder.getFiles().length + " items)";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Package Document");
  }

  // =================== UTILITIES ===================
  async executeInDesignCode(args) {
    const code = (typeof args === 'string') ? args : args.code;
    
    // Security: Check if arbitrary code execution is allowed
    const allowArbitraryCode = process.env.INDESIGN_ALLOW_ARBITRARY_CODE;
    if (!allowArbitraryCode || allowArbitraryCode === '0' || allowArbitraryCode.toLowerCase() === 'false') {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Arbitrary code execution is disabled for security reasons.

To enable this feature, set the environment variable:
INDESIGN_ALLOW_ARBITRARY_CODE=1

⚠️  WARNING: This allows execution of any ExtendScript code, which can:
- Access the file system
- Make network connections  
- Execute system commands via InDesign APIs
- Read/modify any InDesign document data

Only enable this if you trust all users and understand the security implications.

Usage: INDESIGN_ALLOW_ARBITRARY_CODE=1 node index.js`
      );
    }
    
    // User has explicitly enabled arbitrary code execution
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
    const response = this.formatResponse(result, "Document View");
    // Also return a picture of the current page, so the assistant can see the layout
    const pageNumber = Number((result.match(/Current Page: (\d+)/) || [])[1] || 1);
    try {
      const preview = await this.renderToPng({ pageIndex: pageNumber - 1, resolution: 72 });
      response.content.push({ type: 'image', data: preview.data.toString('base64'), mimeType: 'image/png' });
    } catch (error) {
      response.content[0].text += `\n(page preview unavailable: ${error.message})`;
    }
    return response;
  }

  // =================== TABLE MANAGEMENT (Simplified implementations) ===================
  async createTable(args) {
    const { pageIndex = 0, headerRows = 1, footerRows = 0 } = args;
    const nx = this.jsNum(args.x, 'x'), ny = this.jsNum(args.y, 'y');
    const nw = this.jsNum(args.width, 'width'), nh = this.jsNum(args.height, 'height');
    const rows = Math.floor(this.jsNum(args.rows, 'rows')), columns = Math.floor(this.jsNum(args.columns, 'columns'));
    const head = Math.floor(this.jsNum(headerRows, 'headerRows')), foot = Math.floor(this.jsNum(footerRows, 'footerRows'));
    const pageIdx = this.jsNum(pageIndex, 'pageIndex');
    if (columns < 1 || rows < head + foot + 1) {
      throw new Error(`rows (${rows}) must be at least headerRows + footerRows + 1 (${head + foot + 1}); columns must be >= 1`);
    }

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${pageIdx});
      var frame = page.textFrames.add();
      var table;
      try {
        frame.geometricBounds = [${ny}, ${nx}, ${ny + nh}, ${nx + nw}];
        table = frame.tables.add();
        table.columnCount = ${columns};
        table.bodyRowCount = ${rows - head - foot};
        ${head > 0 ? `table.headerRowCount = ${head};` : ''}
        ${foot > 0 ? `table.footerRowCount = ${foot};` : ''}
      } catch (e) {
        frame.remove();
        throw e;
      }
      "Table created: frame id=" + frame.id + " page=" + (${pageIdx} + 1) + " rows=" + table.rows.length + " columns=" + table.columns.length +
        " (header " + table.headerRowCount + ", footer " + table.footerRowCount + "). " + __overflowNote(frame);
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Table");
  }

  async populateTable(args) {
    const { tableIndex, pageIndex = 0, data, includeHeaders = true } = args;
    if (!Array.isArray(data) || data.some((r) => !Array.isArray(r))) throw new Error('data must be an array of arrays');
    const cells = data.map((row) => row.map((v) => (v === null || v === undefined ? '' : String(v))));

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${this.jsNum(pageIndex, 'pageIndex')});
      var tables = [];
      for (var i = 0; i < page.textFrames.length; i++) {
        for (var j = 0; j < page.textFrames[i].tables.length; j++) tables.push(page.textFrames[i].tables[j]);
      }
      var wanted = ${this.jsNum(tableIndex, 'tableIndex')};
      if (wanted < 0 || wanted >= tables.length) {
        throw new Error("Table index " + wanted + " not found. The page has " + tables.length + " table(s).");
      }
      var table = tables[wanted];
      var tableData = ${this.jsJson(cells)};
      // includeHeaders=false: leave the header rows alone and start in the first body row
      var offset = ${includeHeaders ? '0' : 'table.headerRowCount'};
      var written = 0;
      for (var r = 0; r < tableData.length && r + offset < table.rows.length; r++) {
        var rowCells = table.rows[r + offset].cells;
        for (var c = 0; c < tableData[r].length && c < rowCells.length; c++) {
          rowCells[c].contents = tableData[r][c];
          written++;
        }
      }
      "Table populated: " + written + " cell(s) written (table has " + table.rows.length + " rows x " + table.columns.length + " columns)";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Populate Table");
  }

  // =================== LAYER MANAGEMENT (Simplified implementations) ===================
  async createLayer(args) {
    const { name, color, visible = true, locked = false } = args;

    const script = `
      var doc = __requireDoc();
      if (doc.layers.itemByName(${this.jsStr(name)}).isValid) throw new Error("Layer already exists: " + ${this.jsStr(name)});
      var layer = doc.layers.add();
      try {
        layer.name = ${this.jsStr(name)};
        layer.visible = ${visible ? 'true' : 'false'};
        layer.locked = ${locked ? 'true' : 'false'};
        ${color ? `
          var colour;
          try { colour = UIColors[${this.jsStr(String(color).toUpperCase())}]; } catch (ce) { colour = undefined; }
          if (colour === undefined) throw new Error("Unknown layer colour: " + ${this.jsStr(color)} + " (examples: RED, BLUE, LIGHT_BLUE, GREEN, YELLOW)");
          layer.layerColor = colour;
        ` : ''}
      } catch (e) {
        layer.remove();
        throw e;
      }
      "Layer " + ${this.jsStr(name)} + " created";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Layer");
  }

  async setActiveLayer(args) {
    const { layerName } = args;

    const script = `
      var doc = __requireDoc();
      var layer = doc.layers.itemByName(${this.jsStr(layerName)});
      if (!layer.isValid) throw new Error("Layer not found: " + ${this.jsStr(layerName)});
      doc.activeLayer = layer;
      "Active layer set to: " + ${this.jsStr(layerName)};
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
            preflightProfile = app.preflightProfiles.itemByName(${this.jsStr(profile)});
            if (!preflightProfile.isValid) {
              preflightProfile = app.preflightProfiles[0];
            }
          ` : `
            preflightProfile = app.preflightProfiles[0];
          `}
          
          var proc = app.preflightProcesses.add(doc, preflightProfile);
          proc.waitForProcess();
          var summary = String(proc.processResults).replace(/\\s+$/, "");
          proc.remove();
          "Preflight with profile " + preflightProfile.name + ": " + (summary === "None" ? "no issues found" : "issues found:\\n" + summary);
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
    const zoom = { FIT_PAGE: 'FIT_PAGE', FIT_SPREAD: 'FIT_SPREAD', ACTUAL_SIZE: 'ACTUAL_SIZE' }[fitOption];
    if (!zoom) throw new Error(`Invalid fitOption: ${fitOption} (use FIT_PAGE, FIT_SPREAD or ACTUAL_SIZE)`);

    const script = `
      var doc = __requireDoc();
      ${pageIndex !== undefined ? `app.activeWindow.activePage = __page(doc, ${this.jsNum(pageIndex, 'pageIndex')});` : ''}
      app.activeWindow.zoom(ZoomOptions.${zoom});
      "Zoom applied: ${zoom}${pageIndex !== undefined ? ` on page ${this.jsNum(pageIndex, 'pageIndex') + 1}` : ''}";
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Zoom to Page");
  }

  async dataMerge(args) {
    const { dataSourcePath, outputFolder, fileFormat = 'PDF', recordRange = 'all' } = args;

    // Security: Require confirmation for bulk file creation
    this.validateDestructiveOperation(args, 'DATA MERGE', `${outputFolder} (creates merged files)`);

    // Security: Validate both data source and output paths
    const validatedDataSource = this.validateFilePath(dataSourcePath);
    const validatedOutputFolder = this.validateFilePath(outputFolder);
    if (!['PDF', 'INDD', 'BOTH'].includes(fileFormat)) throw new Error(`Invalid fileFormat: ${fileFormat} (use PDF, INDD or BOTH)`);

    const range = String(recordRange).trim().toLowerCase();
    let selection;
    if (range === 'all') selection = { mode: 'ALL_RECORDS' };
    else if (/^\d+$/.test(range)) selection = { mode: 'ONE_RECORD', number: Number(range) };
    else if (/^\d+\s*-\s*\d+$/.test(range)) selection = { mode: 'RANGE', range: range.replace(/\s+/g, '') };
    else throw new Error(`Invalid recordRange: ${recordRange} (use "all", "3" or "1-10")`);

    const script = `
      var doc = __requireDoc();
      var source = File(${this.jsStr(validatedDataSource)});
      if (!source.exists) throw new Error("Data source file not found: " + source.fsName);
      var outDir = Folder(${this.jsStr(validatedOutputFolder)});
      if (!outDir.exists && !outDir.create()) throw new Error("Cannot create folder: " + outDir.fsName);
      
      var dm = doc.dataMergeProperties;
      dm.selectDataSource(source);
      var prefs = dm.dataMergePreferences;
      prefs.recordSelection = RecordSelection.${selection.mode};
      ${selection.number !== undefined ? `prefs.recordNumber = ${selection.number};` : ''}
      ${selection.range !== undefined ? `prefs.recordRange = ${this.jsStr(selection.range)};` : ''}
      
      var base = doc.name.replace(/\\.indd$/i, "") + "_merged";
      var written = [];
      ${fileFormat === 'PDF' || fileFormat === 'BOTH' ? `
        var pdf = File(outDir.fsName + "/" + base + ".pdf");
        dm.exportFile(pdf);
        if (!pdf.exists) throw new Error("Data merge did not create " + pdf.fsName);
        written.push(pdf.fsName);
      ` : ''}
      ${fileFormat === 'INDD' || fileFormat === 'BOTH' ? `
        dm.mergeRecords();
        var merged = app.activeDocument;
        var indd = File(outDir.fsName + "/" + base + ".indd");
        merged.save(indd);
        merged.close(SaveOptions.NO);
        app.activeDocument = doc;
        if (!indd.exists) throw new Error("Data merge did not create " + indd.fsName);
        written.push(indd.fsName);
      ` : ''}
      "Data merge completed (records: ${selection.mode === 'ALL_RECORDS' ? 'all' : range}). Files written:\\n" + written.join("\\n");
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Data Merge");
  }

  async run() {
    const transport = new StdioServerTransport();
    // Tell the client to re-read the tool list after every (re)start, so cached schemas cannot go stale
    this.server.oninitialized = () => {
      this.server.sendToolListChanged().catch(() => {});
    };
    await this.server.connect(transport);
    console.error('Complete InDesign MCP server running on stdio');
  }
}

// Only start the server when run directly (tests import the class)
if (process.argv[1] && fs.realpathSync(process.argv[1]) === fs.realpathSync(__filename)) {
  const server = new InDesignMCPServer();
  server.run().catch(console.error);
}