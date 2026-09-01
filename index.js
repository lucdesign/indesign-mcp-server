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

    // Security: Define allowed directories for file operations
    this.allowedDirectories = [
      os.homedir(), // User home directory
      '/Users/Shared', // Shared directory
      // Add more as needed via environment variable
      ...(process.env.INDESIGN_ALLOWED_DIRS ? process.env.INDESIGN_ALLOWED_DIRS.split(':') : [])
    ];

    this.setupToolHandlers();
  }

  // Security: Path validation to prevent directory traversal
  validateFilePath(filePath) {
    if (!filePath || typeof filePath !== 'string') {
      throw new Error('Invalid file path provided');
    }

    // Resolve to absolute path
    const resolvedPath = path.resolve(filePath);
    
    // Check if path is within allowed directories
    const isAllowed = this.allowedDirectories.some(allowedDir => {
      const resolvedAllowedDir = path.resolve(allowedDir);
      return resolvedPath.startsWith(resolvedAllowedDir + path.sep) || resolvedPath === resolvedAllowedDir;
    });

    if (!isAllowed) {
      throw new Error(`Access denied: Path '${filePath}' is outside allowed directories. Allowed: ${this.allowedDirectories.join(', ')}`);
    }

    // Additional security checks
    if (resolvedPath.includes('..')) {
      throw new Error('Path traversal detected: .. not allowed in resolved path');
    }

    // Prevent access to sensitive system files
    const prohibitedPaths = ['/etc', '/System', '/usr/bin', '/bin', '/sbin'];
    const isProhibited = prohibitedPaths.some(prohibited => 
      resolvedPath.startsWith(prohibited + path.sep) || resolvedPath === prohibited
    );

    if (isProhibited) {
      throw new Error(`Access denied: Cannot access system directory '${resolvedPath}'`);
    }

    return resolvedPath;
  }

  // Security: User confirmation for destructive operations
  async requireUserConfirmation(operation, target, details = '') {
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm overwrite of existing files', default: false },
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm potential data loss', default: false },
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
          name: 'get_selected_objects',
          description: 'Get information about currently selected objects in InDesign. ESSENTIAL for working with user-selected text frames.',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'get_text_content',
          description: 'Extract raw text content from selected text frame, insertion point, or specific frame. Automatically handles different selection types and normalizes line breaks to spaces. RECOMMENDED for text extraction.',
          inputSchema: {
            type: 'object',
            properties: {
              normalizeSpaces: { type: 'boolean', description: 'Convert line breaks to spaces and remove multiple spaces', default: true },
              frameIndex: { type: 'number', description: 'Optional: specific frame index if nothing selected' },
              pageIndex: { type: 'number', description: 'Page index for frameIndex', default: 0 },
              maxLength: { type: 'number', description: 'Maximum text length to return (0 = unlimited)', default: 0 }
            }
          }
        },
        {
          name: 'list_text_frames',
          description: 'List all text frames on a page with their indices, content preview, and selection status',
          inputSchema: {
            type: 'object',
            properties: {
              pageIndex: { type: 'number', description: 'Page index to inspect', default: 0 }
            }
          }
        },
        {
          name: 'analyze_embedded_objects',
          description: 'Analyze embedded objects (MathML formulas, graphics, etc.) in selected text frame or specified frame',
          inputSchema: {
            type: 'object',
            properties: {
              frameIndex: { type: 'number', description: 'Text frame index (optional if frame is selected)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              maxObjects: { type: 'number', description: 'Maximum number of objects to analyze', default: 5 }
            }
          }
        },
        {
          name: 'insert_markdown_text',
          description: 'Insert markdown text into a text frame with automatic formatting using existing paragraph and character styles. Supports # headers, **bold**, *italic*, etc.',
          inputSchema: {
            type: 'object',
            properties: {
              markdownText: { type: 'string', description: 'Markdown text to insert' },
              frameIndex: { type: 'number', description: 'Text frame index (use list_text_frames or get_selected_objects first)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
              useSelectedFrame: { type: 'boolean', description: 'Use currently selected text frame instead of frameIndex', default: false },
              replaceContent: { type: 'boolean', description: 'Replace existing content or append', default: true }
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
          description: 'Edit properties of an existing text frame. WORKFLOW: First use list_text_frames() or get_selected_objects() to find the correct frameIndex.',
          inputSchema: {
            type: 'object',
            properties: {
              frameIndex: { type: 'number', description: 'Zero-based text frame index from list_text_frames() output. Example: Frame 0 = frameIndex: 0' },
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
          name: 'modify_paragraph_style',
          description: 'Modify properties of an existing paragraph style',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Paragraph style name to modify' },
              fontFamily: { type: 'string', description: 'Font family' },
              fontSize: { type: 'number', description: 'Font size in points' },
              leading: { type: 'number', description: 'Leading (line spacing) in points' },
              spaceBefore: { type: 'number', description: 'Space before paragraph in mm' },
              spaceAfter: { type: 'number', description: 'Space after paragraph in mm' },
              alignment: { type: 'string', enum: ['LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN', 'JUSTIFY'], description: 'Text alignment' },
              textColor: { type: 'string', description: 'Text color (swatch name)' },
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
          description: 'Apply an object style to selected objects',
          inputSchema: {
            type: 'object',
            properties: {
              styleName: { type: 'string', description: 'Object style name' },
              objectIndex: { type: 'number', description: 'Object index on page (optional if objects selected)' },
              pageIndex: { type: 'number', description: 'Page index', default: 0 },
            },
            required: ['styleName'],
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm file overwrite', default: false },
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
              includeImages: { type: 'boolean', description: 'Include images', default: true },
              imageFormat: { type: 'string', enum: ['PNG', 'JPEG', 'GIF'], default: 'PNG' },
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
              confirmDestructive: { type: 'boolean', description: 'REQUIRED: Confirm bulk file creation', default: false },
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
    // Security fix: Use temporary file instead of command line injection
    const tempAppleScript = path.join(__dirname, `temp_applescript_${Date.now()}_${Math.random().toString(36).substring(2)}.scpt`);
    
    try {
      // Write AppleScript to temporary file
      fs.writeFileSync(tempAppleScript, script, 'utf8');
      
      // Execute via file (immune to injection)
      const result = execSync(`osascript "${tempAppleScript}"`, {
        encoding: 'utf8',
        timeout: 30000,
      });
      
      return result.trim();
    } catch (error) {
      throw new Error(`AppleScript execution failed: ${error.message}`);
    } finally {
      // Always cleanup temp file
      if (fs.existsSync(tempAppleScript)) {
        fs.unlinkSync(tempAppleScript);
      }
    }
  }

  async executeInDesignScript(script) {
    const tempScript = path.join(__dirname, 'temp_script.jsx');
    const tempResult = path.join(__dirname, 'temp_result.txt');
    
    // Clean up any previous result file
    if (fs.existsSync(tempResult)) {
      fs.unlinkSync(tempResult);
    }
    
    // Wrap script to capture return value via temp file
    const resultPath = tempResult.replace(/\\/g, '/');
    
    // Auto-capture last expression: prepend __result__ = to last non-empty line
    const lines = script.split('\n');
    let lastExprIdx = -1;
    for (let i = lines.length - 1; i >= 0; i--) {
      const trimmed = lines[i].trim();
      if (trimmed && !trimmed.startsWith('//') && !trimmed.startsWith('}')) {
        lastExprIdx = i;
        break;
      }
    }
    if (lastExprIdx >= 0) {
      const lastLine = lines[lastExprIdx].trim();
      // Don't prepend for statements that aren't expressions
      if (!/^(var |let |const |if |for |while |switch |function |try |\/\/)/.test(lastLine)) {
        lines[lastExprIdx] = 'var __result__ = ' + lines[lastExprIdx];
      }
    }
    const processedScript = lines.join('\n');
    
    const wrappedScript = `
      try {
        ${processedScript}
      } catch (error) {
        var __result__ = "ERROR: " + error.message + " (Line: " + (error.line || "unknown") + ")";
      }
      if (typeof __result__ !== "undefined" && __result__ !== null) {
        var __f__ = new File("${resultPath}");
        __f__.encoding = "UTF-8";
        __f__.open("w");
        __f__.write(String(__result__));
        __f__.close();
      }
    `;
    
    fs.writeFileSync(tempScript, wrappedScript);

    try {
      const appleScript = `
        tell application "Adobe InDesign 2026"
          activate
          do script (POSIX file "${tempScript}") language javascript
        end tell
      `;
      
      await this.executeAppleScript(appleScript);
      
      // Read result from temp file
      if (fs.existsSync(tempResult)) {
        const result = fs.readFileSync(tempResult, 'utf8');
        return result;
      }
      return "(no return value)";
    } finally {
      if (fs.existsSync(tempScript)) {
        fs.unlinkSync(tempScript);
      }
      if (fs.existsSync(tempResult)) {
        fs.unlinkSync(tempResult);
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
    
    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);
    
    const script = `
      try {
        var file = File("${validatedPath.replace(/\\/g, '\\\\')}");
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
    const { filePath } = args;
    
    // Security: Require confirmation for destructive file operations
    if (filePath) {
      this.validateDestructiveOperation(args, 'SAVE DOCUMENT', filePath);
    }
    
    // Security: Validate file path if provided
    const validatedPath = filePath ? this.validateFilePath(filePath) : null;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open to save";
      } else {
        var doc = app.activeDocument;
        try {
          ${validatedPath ? `
            var file = File("${validatedPath.replace(/\\/g, '\\\\')}");
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
    
    // Security: Require confirmation if closing with save or potential data loss
    if (save) {
      this.validateDestructiveOperation(args, 'CLOSE AND SAVE DOCUMENT', 'current document');
    } else {
      this.validateDestructiveOperation(args, 'CLOSE WITHOUT SAVING', 'unsaved changes will be lost');
    }
    
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
    const { normalizeSpaces = true, frameIndex, pageIndex = 0, maxLength = 0 } = args;

    const script = `
      if (app.documents.length === 0) {
        "ERROR: No document open";
      } else {
        var doc = app.activeDocument;
        var textContent = null;
        var source = "";
        
        // Strategy 1: Try to get text from selection
        if (app.selection.length > 0) {
          var sel = app.selection[0];
          
          if (sel.hasOwnProperty('contents') && sel.contents && sel.contents.length > 0) {
            // Direct text frame selection with content
            textContent = sel.contents;
            source = "Selected text frame";
          } else if (sel.parentTextFrames && sel.parentTextFrames.length > 0) {
            // Insertion point or text selection - get parent frame
            textContent = sel.parentTextFrames[0].contents;
            source = "Parent text frame of selected insertion point";
          } else if (sel.constructor.name === "Text") {
            // Selected text content
            textContent = sel.contents;
            source = "Selected text content";
          } else {
            source = "Selected object (" + sel.constructor.name + ") has no text content";
          }
        }
        
        // Strategy 2: Use specific frame index if no selection or selection has no text
        if (!textContent && typeof ${frameIndex} === "number") {
          try {
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= 0 && ${frameIndex} < page.textFrames.length) {
              textContent = page.textFrames[${frameIndex}].contents;
              source = "Text frame " + ${frameIndex} + " on page " + (${pageIndex} + 1);
            } else {
              source = "ERROR: Frame index " + ${frameIndex} + " invalid. Page " + (${pageIndex} + 1) + " has " + page.textFrames.length + " frames.";
            }
          } catch (e) {
            source = "ERROR: " + e.message;
          }
        }
        
        // Process text content
        if (textContent) {
          var result = "=== TEXT CONTENT ===\\n";
          result += "Source: " + source + "\\n";
          result += "Original length: " + textContent.length + " characters\\n\\n";
          
          var processedText = textContent;
          
          // Normalize spaces if requested
          ${normalizeSpaces ? `
            // Convert all types of line breaks to spaces
            processedText = processedText.replace(/\\r\\n/g, ' ');  // Windows
            processedText = processedText.replace(/\\r/g, ' ');    // Mac  
            processedText = processedText.replace(/\\n/g, ' ');    // Unix
            
            // Remove multiple spaces
            while (processedText.indexOf('  ') !== -1) {
              processedText = processedText.replace(/  /g, ' ');
            }
            
            // Remove leading/trailing spaces (manual trim)
            while (processedText.charAt(0) === ' ') {
              processedText = processedText.substring(1);
            }
            while (processedText.charAt(processedText.length - 1) === ' ') {
              processedText = processedText.substring(0, processedText.length - 1);
            }
          ` : ''}
          
          // Apply length limit if specified
          ${maxLength > 0 ? `
            if (processedText.length > ${maxLength}) {
              processedText = processedText.substring(0, ${maxLength}) + "...";
              result += "Text truncated to " + ${maxLength} + " characters\\n\\n";
            }
          ` : ''}
          
          result += "TEXT:\\n" + processedText;
          result;
        } else {
          "ERROR: No text content found. " + source + "\\n\\nTip: Select a text frame, insertion point, or specify frameIndex parameter.";
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Get Text Content");
  }

  async listTextFrames(args) {
    const { pageIndex = 0 } = args;
    
    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          if (${pageIndex} >= doc.pages.length || ${pageIndex} < 0) {
            "Invalid page index: ${pageIndex}. Document has " + doc.pages.length + " pages.";
          } else {
            var page = doc.pages[${pageIndex}];
            var result = "=== TEXT FRAMES ON PAGE " + (${pageIndex} + 1) + " ===\\n";
            
            if (page.textFrames.length === 0) {
              result += "No text frames found on this page.\\n";
              result += "TIP: Create a text frame first or select an existing one.";
            } else {
              for (var i = 0; i < page.textFrames.length; i++) {
                var frame = page.textFrames[i];
                var content = frame.contents.substring(0, 60);
                if (frame.contents.length > 60) content += "...";
                
                result += "Frame " + i + ": ";
                if (content.length === 0) {
                  result += "(empty frame)";
                } else {
                  result += content.replace(/\\r/g, "↵").replace(/\\n/g, "↵").replace(/\\t/g, "→").replace(/\\u00A0/g, "␣").replace(/\\u2002/g, "⎵").replace(/\\u2003/g, "⎸").replace(/\\u2004/g, "⅓").replace(/\\u2005/g, "¼").replace(/\\u2006/g, "⅙").replace(/\\u2007/g, "♦").replace(/\\u2008/g, "⅛").replace(/\\u2009/g, "◦").replace(/\\u200A/g, "·").replace(/\\u200B/g, "‌").replace(/\\u202F/g, "◦").replace(/\\u205F/g, "▪");
                }
                result += "\\n";
              }
              result += "\\nUSAGE: Use frameIndex 0-" + (page.textFrames.length - 1) + " with edit_text_frame()";
            }
            
            result;
          }
        } catch (e) {
          "Error listing text frames: " + e.message;
        }
      }
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
        
        // Try to get frame from selection or frameIndex
        if (app.selection.length > 0 && app.selection[0].hasOwnProperty('contents')) {
          frame = app.selection[0];
        } else if (typeof ${frameIndex} === "number") {
          var page = doc.pages[${pageIndex}];
          if (${frameIndex} >= 0 && ${frameIndex} < page.textFrames.length) {
            frame = page.textFrames[${frameIndex}];
          }
        }
        
        if (!frame) {
          "ERROR: No text frame found. Select a frame or specify frameIndex.";
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

  async insertMarkdownText(args) {
    const { 
      markdownText, 
      frameIndex, 
      pageIndex = 0, 
      useSelectedFrame = false, 
      replaceContent = true 
    } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var textFrame;
          
          ${useSelectedFrame ? `
            // Use selected frame
            var selection = app.selection;
            if (selection.length === 0 || !selection[0].hasOwnProperty('contents')) {
              "No text frame selected. Please select a text frame first or use frameIndex parameter.";
            } else {
              textFrame = selection[0];
            }
          ` : `
            // Use frameIndex
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
              "Invalid text frame index: ${frameIndex}. Page ${pageIndex + 1} has " + page.textFrames.length + " text frames. Use list_text_frames() to see available frames.";
            } else {
              textFrame = page.textFrames[${frameIndex}];
            }
          `}
          
          if (textFrame) {
            // Convert markdown to formatted text
            var markdownContent = \`${markdownText.replace(/`/g, '\\`').replace(/\$/g, '\\$')}\`;
            
            ${replaceContent ? 'textFrame.contents = "";' : ''}
            
            // Simple markdown parsing
            var lines = markdownContent.split('\\n');
            var story = textFrame.parentStory;
            var insertionPoint = story.insertionPoints[-1];
            
            for (var i = 0; i < lines.length; i++) {
              var line = lines[i];
              
              // Skip empty lines but add paragraph break
              if (line.trim() === '') {
                if (i < lines.length - 1) {
                  insertionPoint.contents = '\\r';
                  insertionPoint = story.insertionPoints[-1];
                }
                continue;
              }
              
              // Headers
              if (line.match(/^#{1,6}\\s/)) {
                var level = line.match(/^#{1,6}/)[0].length;
                var headerText = line.replace(/^#{1,6}\\s/, '');
                
                insertionPoint.contents = headerText;
                
                // Apply header style based on level
                var headerStyle = null;
                var styleNames = ["Header 1", "Heading 1", "H1", "Header1"];
                for (var s = 0; s < styleNames.length; s++) {
                  try {
                    headerStyle = doc.paragraphStyles.itemByName(styleNames[s]);
                    if (headerStyle.isValid) break;
                  } catch (e) {}
                }
                
                if (headerStyle && headerStyle.isValid) {
                  var range = story.characters.itemByRange(
                    insertionPoint.index - headerText.length,
                    insertionPoint.index - 1
                  );
                  range.appliedParagraphStyle = headerStyle;
                }
                
              }
              // Bold text **text**
              else if (line.indexOf('**') !== -1) {
                var parts = line.split('**');
                for (var p = 0; p < parts.length; p++) {
                  insertionPoint.contents = parts[p];
                  
                  if (p % 2 === 1) { // Bold parts
                    var boldStyle = null;
                    try {
                      boldStyle = doc.characterStyles.itemByName("Bold");
                      if (!boldStyle.isValid) {
                        boldStyle = doc.characterStyles.itemByName("Strong");
                      }
                    } catch (e) {}
                    
                    if (boldStyle && boldStyle.isValid) {
                      var range = story.characters.itemByRange(
                        insertionPoint.index - parts[p].length,
                        insertionPoint.index - 1
                      );
                      range.appliedCharacterStyle = boldStyle;
                    } else {
                      // Fallback to manual bold
                      var range = story.characters.itemByRange(
                        insertionPoint.index - parts[p].length,
                        insertionPoint.index - 1
                      );
                      range.fontStyle = "Bold";
                    }
                  }
                  insertionPoint = story.insertionPoints[-1];
                }
              }
              // Italic text *text*
              else if (line.indexOf('*') !== -1 && line.indexOf('**') === -1) {
                var parts = line.split('*');
                for (var p = 0; p < parts.length; p++) {
                  insertionPoint.contents = parts[p];
                  
                  if (p % 2 === 1) { // Italic parts
                    var italicStyle = null;
                    try {
                      italicStyle = doc.characterStyles.itemByName("Italic");
                      if (!italicStyle.isValid) {
                        italicStyle = doc.characterStyles.itemByName("Emphasis");
                      }
                    } catch (e) {}
                    
                    if (italicStyle && italicStyle.isValid) {
                      var range = story.characters.itemByRange(
                        insertionPoint.index - parts[p].length,
                        insertionPoint.index - 1
                      );
                      range.appliedCharacterStyle = italicStyle;
                    } else {
                      // Fallback to manual italic
                      var range = story.characters.itemByRange(
                        insertionPoint.index - parts[p].length,
                        insertionPoint.index - 1
                      );
                      range.fontStyle = "Italic";
                    }
                  }
                  insertionPoint = story.insertionPoints[-1];
                }
              }
              // Regular paragraph
              else {
                insertionPoint.contents = line;
                insertionPoint = story.insertionPoints[-1];
              }
              
              // Add paragraph break except for last line
              if (i < lines.length - 1) {
                insertionPoint.contents = '\\r';
                insertionPoint = story.insertionPoints[-1];
              }
            }
            
            "Markdown text inserted successfully. Applied available paragraph and character styles.";
          }
        } catch (e) {
          "Error inserting markdown text: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Insert Markdown Text");
  }

  async fixTypographyInSelection(args) {
    const { 
      frameIndex, 
      pageIndex = 0, 
      useSelectedFrame = false,
      fixDates = true,
      fixQuotes = true, 
      fixDashes = true,
      fixSpaces = true
    } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var textFrame;
          
          ${useSelectedFrame ? `
            var selection = app.selection;
            if (selection.length === 0 || !selection[0].hasOwnProperty('contents')) {
              "No text frame selected. Please select a text frame first.";
            } else {
              textFrame = selection[0];
            }
          ` : `
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
              "Invalid text frame index: ${frameIndex}. Use list_text_frames() to see available frames.";
            } else {
              textFrame = page.textFrames[${frameIndex}];
            }
          `}
          
          if (textFrame) {
            var story = textFrame.parentStory;
            var content = story.contents;
            var changes = 0;
            var changeLog = "=== TYPOGRAPHY FIXES ===\\n";
            
            ${fixDates ? `
              // Use existing GREP search "DATUM" if available, otherwise fallback to pattern
              try {
                var datumQuery = doc.findGrepPreferences.itemByName("DATUM");
                if (datumQuery.isValid) {
                  // Clear preferences
                  app.findGrepPreferences = NothingEnum.nothing;
                  app.changeGrepPreferences = NothingEnum.nothing;
                  
                  // Load existing DATUM search
                  app.findGrepPreferences.findWhat = datumQuery.findWhat;
                  var foundDates = story.findGrep();
                  
                  if (foundDates.length > 0) {
                    for (var d = 0; d < foundDates.length; d++) {
                      var dateText = foundDates[d].contents;
                      // Replace normal spaces with thin spaces in the found date
                      var fixedDate = dateText.replace(/(\\d{1,2})\\. (\\d{1,2})\\. (\\d{4})/g, "$1.\\u2009$2.\\u2009$3");
                      foundDates[d].contents = fixedDate;
                    }
                    changes += foundDates.length;
                    changeLog += "Fixed " + foundDates.length + " date(s) using DATUM search with thin spaces\\n";
                  }
                  
                  // Clear preferences
                  app.findGrepPreferences = NothingEnum.nothing;
                  app.changeGrepPreferences = NothingEnum.nothing;
                } else {
                  throw new Error("DATUM search not found, using fallback");
                }
              } catch (e) {
                // Fallback to manual pattern search
                var datePattern = /(\\d{1,2})\\. (\\d{1,2})\\. (\\d{4})/g;
                var dateMatches = content.match(datePattern);
                if (dateMatches) {
                  var newContent = content.replace(datePattern, function(match, dd, mm, yyyy) {
                    return dd + ".\\u2009" + mm + ".\\u2009" + yyyy;
                  });
                  story.contents = newContent;
                  changes += dateMatches.length;
                  changeLog += "Fixed " + dateMatches.length + " date(s) with thin spaces (fallback pattern)\\n";
                }
              }
            ` : ''}
            
            ${fixQuotes ? `
              // Fix straight quotes to typographic quotes
              var beforeQuotes = (story.contents.match(/"/g) || []).length;
              if (beforeQuotes > 0) {
                // Simple quote replacement (could be enhanced)
                story.contents = story.contents.replace(/"/g, "„").replace(/„/g, """).replace(/"/g, "„");
                changeLog += "Fixed " + Math.floor(beforeQuotes/2) + " quote pair(s)\\n";
                changes += Math.floor(beforeQuotes/2);
              }
            ` : ''}
            
            ${fixDashes ? `
              // Fix double hyphens to em dashes
              var dashMatches = (story.contents.match(/--/g) || []).length;
              if (dashMatches > 0) {
                story.contents = story.contents.replace(/--/g, "—");
                changes += dashMatches;
                changeLog += "Fixed " + dashMatches + " double hyphen(s) to em dash\\n";
              }
              
              // Fix space-hyphen-space to en dash
              var enDashMatches = (story.contents.match(/ - /g) || []).length;
              if (enDashMatches > 0) {
                story.contents = story.contents.replace(/ - /g, " – ");
                changes += enDashMatches;
                changeLog += "Fixed " + enDashMatches + " hyphen(s) to en dash\\n";
              }
            ` : ''}
            
            ${fixSpaces ? `
              // Fix multiple spaces
              var multiSpaceMatches = (story.contents.match(/  +/g) || []).length;
              if (multiSpaceMatches > 0) {
                story.contents = story.contents.replace(/  +/g, " ");
                changes += multiSpaceMatches;
                changeLog += "Fixed " + multiSpaceMatches + " multiple space(s)\\n";
              }
              
              // Fix trailing spaces
              var lines = story.contents.split('\\r');
              var trailingSpaces = 0;
              for (var i = 0; i < lines.length; i++) {
                if (lines[i].match(/ +$/)) {
                  lines[i] = lines[i].replace(/ +$/, '');
                  trailingSpaces++;
                }
              }
              if (trailingSpaces > 0) {
                story.contents = lines.join('\\r');
                changes += trailingSpaces;
                changeLog += "Fixed " + trailingSpaces + " trailing space(s)\\n";
              }
            ` : ''}
            
            if (changes > 0) {
              changeLog += "\\nTotal fixes applied: " + changes;
            } else {
              changeLog += "No typography issues found.";
            }
            
            changeLog;
          }
        } catch (e) {
          "Error fixing typography: " + e.message;
        }
      }
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
          var textFrame;
          
          ${useSelectedFrame ? `
            var selection = app.selection;
            if (selection.length === 0 || !selection[0].hasOwnProperty('contents')) {
              "No text frame selected. Please select a text frame first.";
            } else {
              textFrame = selection[0];
            }
          ` : `
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
              "Invalid text frame index: ${frameIndex}. Use list_text_frames() to see available frames.";
            } else {
              textFrame = page.textFrames[${frameIndex}];
            }
          `}
          
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
    const { 
      frameIndex, 
      pageIndex = 0, 
      useSelectedFrame = false,
      fixParagraphs = true,
      fixDashes = true,
      fixLists = true,
      fixFormatting = true,
      fixChapterNumbers = true,
      fixSpaces = true
    } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var textFrame;
          
          ${useSelectedFrame ? `
            var selection = app.selection;
            if (selection.length === 0 || !selection[0].hasOwnProperty('contents')) {
              "No text frame selected. Please select a text frame first.";
            } else {
              textFrame = selection[0];
            }
          ` : `
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
              "Invalid text frame index: ${frameIndex}. Use list_text_frames() to see available frames.";
            } else {
              textFrame = page.textFrames[${frameIndex}];
            }
          `}
          
          if (textFrame) {
            var story = textFrame.parentStory;
            var changes = 0;
            var changeLog = "=== TEXT CLEANING REPORT ===\\n";
            
            ${fixSpaces ? `
              // 1. Remove trailing spaces at end of paragraphs
              var beforeTrailing = story.contents;
              story.contents = story.contents.replace(/ +\\r/g, '\\r');
              var trailingRemoved = beforeTrailing.length - story.contents.length;
              if (trailingRemoved > 0) {
                changes += trailingRemoved;
                changeLog += "✓ Removed " + trailingRemoved + " trailing space(s)\\n";
              }
              
              // Remove multiple spaces
              var beforeMultiple = story.contents;
              story.contents = story.contents.replace(/  +/g, ' ');
              var multipleRemoved = beforeMultiple.length - story.contents.length;
              if (multipleRemoved > 0) {
                changes += multipleRemoved;
                changeLog += "✓ Fixed " + multipleRemoved + " multiple space(s)\\n";
              }
            ` : ''}
            
            ${fixParagraphs ? `
              // 2. Fix double paragraph breaks (fake spacing)
              var doublePars = (story.contents.match(/\\r\\r+/g) || []).length;
              if (doublePars > 0) {
                story.contents = story.contents.replace(/\\r\\r+/g, '\\r');
                changes += doublePars;
                changeLog += "✓ Fixed " + doublePars + " double paragraph break(s)\\n";
              }
              
              // Fix line breaks that should be paragraphs (\\n to \\r)
              var lineBreaks = (story.contents.match(/\\n/g) || []).length;
              if (lineBreaks > 0) {
                story.contents = story.contents.replace(/\\n/g, '\\r');
                changes += lineBreaks;
                changeLog += "✓ Converted " + lineBreaks + " line break(s) to paragraphs\\n";
              }
            ` : ''}
            
            ${fixDashes ? `
              // 3. Fix hyphens to n-dashes for ranges and thoughts
              var hyphenRanges = (story.contents.match(/\\d+-\\d+/g) || []).length; // 1990-2000
              if (hyphenRanges > 0) {
                story.contents = story.contents.replace(/(\\d+)-(\\d+)/g, '$1–$2');
                changes += hyphenRanges;
                changeLog += "✓ Fixed " + hyphenRanges + " number range(s) to n-dash\\n";
              }
              
              var thoughtDashes = (story.contents.match(/ - /g) || []).length;
              if (thoughtDashes > 0) {
                story.contents = story.contents.replace(/ - /g, ' – ');
                changes += thoughtDashes;
                changeLog += "✓ Fixed " + thoughtDashes + " thought dash(es) to n-dash\\n";
              }
            ` : ''}
            
            ${fixLists ? `
              // 4. Remove manual bullet lists and dashes
              var bulletLists = (story.contents.match(/^[•·-]\\s/gm) || []).length;
              if (bulletLists > 0) {
                story.contents = story.contents.replace(/^[•·-]\\s+/gm, '');
                changes += bulletLists;
                changeLog += "✓ Removed " + bulletLists + " manual bullet(s)/dash(es)\\n";
              }
              
              var tabBullets = (story.contents.match(/^\\t[•·-]\\s/gm) || []).length;
              if (tabBullets > 0) {
                story.contents = story.contents.replace(/^\\t[•·-]\\s+/gm, '');
                changes += tabBullets;
                changeLog += "✓ Removed " + tabBullets + " tabbed bullet(s)\\n";
              }
            ` : ''}
            
            ${fixChapterNumbers ? `
              // 5. Remove hardcoded chapter numbers
              var chapterNumbers = (story.contents.match(/^(Kapitel|Chapter|Teil|Part)\\s+\\d+[.:]*\\s*/gmi) || []).length;
              if (chapterNumbers > 0) {
                story.contents = story.contents.replace(/^(Kapitel|Chapter|Teil|Part)\\s+\\d+[.:]*\\s*/gmi, '');
                changes += chapterNumbers;
                changeLog += "✓ Removed " + chapterNumbers + " hardcoded chapter number(s)\\n";
              }
              
              var romanNumbers = (story.contents.match(/^[IVX]+[.:]*\\s*/gm) || []).length;
              if (romanNumbers > 0) {
                story.contents = story.contents.replace(/^[IVX]+[.:]*\\s*/gm, '');
                changes += romanNumbers;
                changeLog += "✓ Removed " + romanNumbers + " roman numeral(s)\\n";
              }
            ` : ''}
            
            ${fixFormatting ? `
              // 6. Reset manual formatting (prepare for character styles)
              try {
                // Reset all character formatting to prepare for proper styles
                story.characters.everyItem().fontStyle = "Regular";
                story.characters.everyItem().appliedCharacterStyle = doc.characterStyles.item("[None]");
                changeLog += "✓ Reset all manual bold/italic formatting\\n";
                changes += 1;
              } catch (e) {
                changeLog += "⚠ Could not reset formatting: " + e.message + "\\n";
              }
            ` : ''}
            
            if (changes > 0) {
              changeLog += "\\n=== SUMMARY ===\\n";
              changeLog += "Total changes applied: " + changes + "\\n";
              changeLog += "Text is now ready for proper InDesign formatting!\\n";
              changeLog += "Next steps: Apply paragraph styles and character styles.";
            } else {
              changeLog += "No issues found - text is already clean!";
            }
            
            changeLog;
          }
        } catch (e) {
          "Error cleaning text: " + e.message;
        }
      }
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
          var textFrame;
          
          ${useSelectedFrame ? `
            var selection = app.selection;
            if (selection.length === 0 || !selection[0].hasOwnProperty('contents')) {
              "No text frame selected. Please select a text frame first.";
            } else {
              textFrame = selection[0];
            }
          ` : `
            var page = doc.pages[${pageIndex}];
            if (${frameIndex} >= page.textFrames.length || ${frameIndex} < 0) {
              "Invalid text frame index: ${frameIndex}. Use list_text_frames() to see available frames.";
            } else {
              textFrame = page.textFrames[${frameIndex}];
            }
          `}
          
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
                var char = story.characters[i];
                if (char.fontStyle !== "Regular" && char.appliedCharacterStyle.name === "[None]") {
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
            "Invalid text frame index: ${frameIndex}. Page ${pageIndex + 1} has " + page.textFrames.length + " text frames (valid indices: 0-" + (page.textFrames.length - 1) + "). Use list_text_frames() to see available frames.";
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

    // Security: Validate image path
    const validatedPath = this.validateFilePath(imagePath);

    const script = `
      if (app.documents.length === 0) {
        "No document open. Please create a document first.";
      } else {
        var doc = app.activeDocument;
        try {
          var page = doc.pages[${pageIndex}];
          var imageFile = File("${validatedPath.replace(/\\/g, '\\\\')}");
          
          if (!imageFile.exists) {
            "Image file not found: ${validatedPath}";
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

  async modifyCharacterStyle(args) {
    const { styleName, fontFamily, fontStyle, fontSize, textColor, tracking } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          // Find the character style
          var cStyle = doc.characterStyles.itemByName("${styleName}");
          if (!cStyle.isValid) {
            "Character style '${styleName}' not found";
          } else {
            var changes = [];
            
            ${fontFamily ? `
              cStyle.appliedFont = "${fontFamily}";
              changes.push("Font Family: ${fontFamily}");
            ` : ''}
            
            ${fontStyle ? `
              cStyle.fontStyle = "${fontStyle}";
              changes.push("Font Style: ${fontStyle}");
            ` : ''}
            
            ${fontSize ? `
              cStyle.pointSize = ${fontSize};
              changes.push("Font Size: ${fontSize}pt");
            ` : ''}
            
            ${textColor ? `
              try {
                var colorSwatch = doc.colors.itemByName("${textColor}");
                if (colorSwatch.isValid) {
                  cStyle.fillColor = colorSwatch;
                  changes.push("Text Color: ${textColor}");
                } else {
                  changes.push("Warning: Color '${textColor}' not found");
                }
              } catch (e) {
                changes.push("Warning: Could not apply color '${textColor}': " + e.message);
              }
            ` : ''}
            
            ${tracking ? `
              cStyle.tracking = ${tracking};
              changes.push("Tracking: ${tracking}");
            ` : ''}
            
            if (changes.length > 0) {
              "Character style '${styleName}' modified:\\n" + changes.join("\\n");
            } else {
              "No properties specified to modify for character style '${styleName}'";
            }
          }
        } catch (e) {
          "Error modifying character style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Modify Character Style");
  }

  async modifyParagraphStyle(args) {
    const { styleName, fontFamily, fontSize, leading, spaceBefore, spaceAfter, alignment, textColor } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          // Find the paragraph style
          var pStyle = doc.paragraphStyles.itemByName("${styleName}");
          if (!pStyle.isValid) {
            "Paragraph style '${styleName}' not found";
          } else {
            var changes = [];
            
            ${fontFamily ? `
              pStyle.appliedFont = "${fontFamily}";
              changes.push("Font Family: ${fontFamily}");
            ` : ''}
            
            ${fontSize ? `
              pStyle.pointSize = ${fontSize};
              changes.push("Font Size: ${fontSize}pt");
            ` : ''}
            
            ${leading ? `
              pStyle.leading = ${leading};
              changes.push("Leading: ${leading}pt");
            ` : ''}
            
            ${spaceBefore ? `
              pStyle.spaceBefore = "${spaceBefore}mm";
              changes.push("Space Before: ${spaceBefore}mm");
            ` : ''}
            
            ${spaceAfter ? `
              pStyle.spaceAfter = "${spaceAfter}mm";
              changes.push("Space After: ${spaceAfter}mm");
            ` : ''}
            
            ${alignment ? `
              pStyle.justification = Justification.${alignment};
              changes.push("Alignment: ${alignment}");
            ` : ''}
            
            ${textColor ? `
              try {
                var colorSwatch = doc.colors.itemByName("${textColor}");
                if (colorSwatch.isValid) {
                  pStyle.fillColor = colorSwatch;
                  changes.push("Text Color: ${textColor}");
                } else {
                  changes.push("Warning: Color '${textColor}' not found");
                }
              } catch (e) {
                changes.push("Warning: Could not apply color '${textColor}': " + e.message);
              }
            ` : ''}
            
            if (changes.length > 0) {
              "Paragraph style '${styleName}' modified:\\n" + changes.join("\\n");
            } else {
              "No properties specified to modify for paragraph style '${styleName}'";
            }
          }
        } catch (e) {
          "Error modifying paragraph style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Modify Paragraph Style");
  }

  async modifyObjectStyle(args) {
    const { styleName, fillColor, strokeColor, strokeWidth, transparency } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          // Find the object style
          var oStyle = doc.objectStyles.itemByName("${styleName}");
          if (!oStyle.isValid) {
            "Object style '${styleName}' not found";
          } else {
            var changes = [];
            
            ${fillColor ? `
              try {
                var fillSwatch = doc.colors.itemByName("${fillColor}");
                if (fillSwatch.isValid) {
                  oStyle.fillColor = fillSwatch;
                  changes.push("Fill Color: ${fillColor}");
                } else {
                  changes.push("Warning: Fill color '${fillColor}' not found");
                }
              } catch (e) {
                changes.push("Warning: Could not apply fill color '${fillColor}': " + e.message);
              }
            ` : ''}
            
            ${strokeColor ? `
              try {
                var strokeSwatch = doc.colors.itemByName("${strokeColor}");
                if (strokeSwatch.isValid) {
                  oStyle.strokeColor = strokeSwatch;
                  changes.push("Stroke Color: ${strokeColor}");
                } else {
                  changes.push("Warning: Stroke color '${strokeColor}' not found");
                }
              } catch (e) {
                changes.push("Warning: Could not apply stroke color '${strokeColor}': " + e.message);
              }
            ` : ''}
            
            ${strokeWidth ? `
              oStyle.strokeWeight = ${strokeWidth};
              changes.push("Stroke Width: ${strokeWidth}pt");
            ` : ''}
            
            ${transparency ? `
              oStyle.transparencySettings.blendingSettings.opacity = ${100 - transparency};
              changes.push("Transparency: ${transparency}%");
            ` : ''}
            
            if (changes.length > 0) {
              "Object style '${styleName}' modified:\\n" + changes.join("\\n");
            } else {
              "No properties specified to modify for object style '${styleName}'";
            }
          }
        } catch (e) {
          "Error modifying object style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Modify Object Style");
  }

  async createObjectStyle(args) {
    const { name, fillColor, strokeColor, strokeWidth, transparency, baseStyle } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var oStyle = doc.objectStyles.add();
          oStyle.name = "${name}";
          
          ${baseStyle ? `
            try {
              var base = doc.objectStyles.itemByName("${baseStyle}");
              if (base.isValid) {
                oStyle.basedOn = base;
              }
            } catch (e) {}
          ` : ''}
          
          ${fillColor ? `
            try {
              var fillSwatch = doc.colors.itemByName("${fillColor}");
              if (fillSwatch.isValid) {
                oStyle.fillColor = fillSwatch;
              }
            } catch (e) {}
          ` : ''}
          
          ${strokeColor ? `
            try {
              var strokeSwatch = doc.colors.itemByName("${strokeColor}");
              if (strokeSwatch.isValid) {
                oStyle.strokeColor = strokeSwatch;
              }
            } catch (e) {}
          ` : ''}
          
          ${strokeWidth ? `
            oStyle.strokeWeight = ${strokeWidth};
          ` : ''}
          
          ${transparency ? `
            oStyle.transparencySettings.blendingSettings.opacity = ${100 - transparency};
          ` : ''}
          
          "Object style '" + oStyle.name + "' created successfully";
        } catch (e) {
          "Error creating object style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Create Object Style");
  }

  async applyObjectStyle(args) {
    const { styleName, objectIndex, pageIndex = 0 } = args;

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var oStyle = doc.objectStyles.itemByName("${styleName}");
          if (!oStyle.isValid) {
            "Object style '${styleName}' not found";
          } else {
            var objectsToStyle = [];
            
            // Strategy 1: Use selection if available
            if (app.selection.length > 0) {
              for (var i = 0; i < app.selection.length; i++) {
                objectsToStyle.push(app.selection[i]);
              }
            }
            // Strategy 2: Use specific object index
            else if (typeof ${objectIndex} === "number") {
              var page = doc.pages[${pageIndex}];
              if (${objectIndex} >= 0 && ${objectIndex} < page.allPageItems.length) {
                objectsToStyle.push(page.allPageItems[${objectIndex}]);
              } else {
                "Invalid object index: ${objectIndex}. Page ${pageIndex + 1} has " + page.allPageItems.length + " objects.";
              }
            } else {
              "No objects selected and no objectIndex specified. Please select objects or provide objectIndex.";
            }
            
            if (objectsToStyle.length > 0) {
              var appliedCount = 0;
              for (var j = 0; j < objectsToStyle.length; j++) {
                try {
                  if (objectsToStyle[j].hasOwnProperty('appliedObjectStyle')) {
                    objectsToStyle[j].appliedObjectStyle = oStyle;
                    appliedCount++;
                  }
                } catch (e) {}
              }
              "Object style '${styleName}' applied to " + appliedCount + " object(s)";
            }
          }
        } catch (e) {
          "Error applying object style: " + e.message;
        }
      }
    `;

    const result = await this.executeInDesignScript(script);
    return this.formatResponse(result, "Apply Object Style");
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

    // Security: Require confirmation for file export
    this.validateDestructiveOperation(args, 'EXPORT PDF', filePath);

    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);

    const script = `
      if (app.documents.length === 0) {
        "No document open. Please create a document first.";
      } else {
        var doc = app.activeDocument;
        try {
          var pdfFile = File("${validatedPath.replace(/\\/g, '\\\\')}");
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

    // Security: Require confirmation for folder write
    this.validateDestructiveOperation(args, 'EXPORT IMAGES', folderPath);

    // Security: Validate folder path
    const validatedPath = this.validateFilePath(folderPath);

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var exportFolder = Folder("${validatedPath.replace(/\\/g, '\\\\')}");
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

    // Security: Require confirmation for file export
    this.validateDestructiveOperation(args, 'EXPORT EPUB', filePath);

    // Security: Validate file path
    const validatedPath = this.validateFilePath(filePath);

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var epubFile = File("${validatedPath.replace(/\\/g, '\\\\')}");
          
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
          "EPUB exported successfully to: ${validatedPath}";
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

    // Security: Require confirmation for package creation
    this.validateDestructiveOperation(args, 'PACKAGE DOCUMENT', folderPath);

    // Security: Validate folder path
    const validatedPath = this.validateFilePath(folderPath);

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var packageFolder = Folder("${validatedPath.replace(/\\/g, '\\\\')}");
          
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

    // Security: Require confirmation for bulk file creation
    this.validateDestructiveOperation(args, 'DATA MERGE', `${outputFolder} (will create multiple files)`);

    // Security: Validate both data source and output paths
    const validatedDataSource = this.validateFilePath(dataSourcePath);
    const validatedOutputFolder = this.validateFilePath(outputFolder);

    const script = `
      if (app.documents.length === 0) {
        "No document open";
      } else {
        var doc = app.activeDocument;
        try {
          var dataSource = File("${validatedDataSource.replace(/\\/g, '\\\\')}");
          if (!dataSource.exists) {
            "Data source file not found: ${validatedDataSource}";
          } else {
            // Set up data merge
            doc.dataMergeProperties.dataMergeSource = dataSource;
            
            var outputDir = Folder("${validatedOutputFolder.replace(/\\/g, '\\\\')}");
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
            
            "Data merge completed. Files saved to: ${validatedOutputFolder}";
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