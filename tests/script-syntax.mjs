#!/usr/bin/env node
// Offline check (no InDesign needed): generates the ExtendScript of EVERY tool with many
// argument combinations, with awkward strings (quotes, backslashes, umlauts, newlines) in
// every string parameter, and verifies that each script is syntactically valid JavaScript
// and uses only ES3 syntax. This catches the class of bug where generated code breaks the
// interpreter ("Illegal use of reserved word", unbalanced quotes, ...) before it ever runs.
//
//   node tests/script-syntax.mjs

import fs from 'fs';
import os from 'os';
import path from 'path';
import vm from 'vm';

const TMP = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'indesign-mcp-syntax-')));
process.env.INDESIGN_ALLOWED_DIRS = TMP;
process.env.INDESIGN_ALLOW_ARBITRARY_CODE = '1';

const { InDesignMCPServer } = await import('../index.js');
const { buildWrapper } = await import('../lib/script-runtime.js');

const server = new InDesignMCPServer();
let captured = [];
server.executeInDesignScript = async (script) => { captured.push(script); return 'ok'; };
server.executeAppleScript = async () => 'ok';

const callHandler = server.server._requestHandlers.get('tools/call');
const listHandler = server.server._requestHandlers.get('tools/list');
const { tools } = await listHandler({ method: 'tools/list', params: {} }, {});

// ---- schema declarations: what the handlers read must be declared in the input schema (C3) ----
const schemaProblems = [];
{
  const declared = Object.fromEntries(tools.map((t) => [t.name, new Set(Object.keys(t.inputSchema.properties || {}))]));
  const names = tools.map((t) => t.name);
  if (new Set(names).size !== names.length) schemaProblems.push(`duplicate tool names: ${names.filter((n, i) => names.indexOf(n) !== i).join(', ')}`);
  const paragraph = ['name', 'baseStyle', 'fontFamily', 'fontStyle', 'fontSize', 'leading', 'tracking', 'kerning', 'spaceBefore', 'spaceBeforePt', 'spaceAfter', 'spaceAfterPt',
    'firstLineIndent', 'leftIndent', 'rightIndent', 'alignment', 'hyphenation', 'keepWithNext', 'keepLinesTogether', 'keepFirstLines', 'keepLastLines', 'keepWithPrevious',
    'composer', 'capitalization', 'ligatures', 'openType', 'hyphenateWordsLongerThan', 'textColor'];
  const must = {
    create_text_frame: ['content', 'lineBreak', 'paragraphStyle', 'characterStyle', 'fontStyle', 'fontFamily', 'fontSize', 'textColor', 'alignment', 'x', 'y', 'width', 'height', 'pageIndex'],
    edit_text_frame: ['frameId', 'frameIndex', 'pageIndex', 'content', 'lineBreak', 'fontStyle', 'fontFamily', 'fontSize', 'textColor', 'alignment'],
    create_paragraph_style: paragraph,
    modify_paragraph_style: ['styleName', ...paragraph.filter((p) => p !== 'name')],
    export_pdf: ['filePath', 'preset', 'pageRange', 'includeBleed', 'includeSlug', 'confirmDestructive'],
    export_images: ['folderPath', 'format', 'resolution', 'pageRange', 'includeBleed', 'confirmDestructive'],
    create_color_swatch: ['name', 'colorModel', 'colorValues', 'hex', 'preset', 'keepRgb', 'spotColor', 'update'],
    get_text_content: ['frameId', 'frameIndex', 'pageIndex', 'normalizeSpaces', 'maxLength'],
    insert_markdown_text: ['markdownText', 'frameId', 'frameIndex', 'bodyStyle', 'styleMap', 'replaceContent'],
    place_image: ['imagePath', 'x', 'y', 'width', 'height', 'fitOption', 'contentOffsetX', 'contentScale'],
  };
  for (const [tool, props] of Object.entries(must)) {
    if (!declared[tool]) { schemaProblems.push(`tool ${tool} is missing`); continue; }
    for (const p of props) if (!declared[tool].has(p)) schemaProblems.push(`${tool}: parameter "${p}" is implemented but not declared in the input schema`);
  }
  // options that must NOT be advertised any more
  const gone = { export_pdf: ['colorProfile', 'jpegQuality'], export_epub: ['includeImages'], zoom_to_page: [] };
  for (const [tool, props] of Object.entries(gone)) for (const p of props) if (declared[tool]?.has(p)) schemaProblems.push(`${tool}: "${p}" is declared but no longer implemented`);
  // every parameter of every tool has a type and a description
  for (const t of tools) {
    for (const [key, spec] of Object.entries(t.inputSchema.properties || {})) {
      if (!spec.type) schemaProblems.push(`${t.name}.${key} has no type`);
      if (!spec.description && !spec.enum && key !== 'x' && key !== 'y') schemaProblems.push(`${t.name}.${key} has no description`);
    }
    if (!t.description || t.description.length < 20) schemaProblems.push(`${t.name} has no proper description`);
  }
}

const NASTY = 'Na"me \' with \\ back\\slash, $x ${y} `tick` ä ö ü ß   end';
const sample = (key, spec, mode) => {
  switch (spec.type) {
    case 'string':
      if (spec.enum) return mode === 'last' ? spec.enum[spec.enum.length - 1] : spec.enum[0];
      if (/path|folder/i.test(key)) return path.join(TMP, /folder/i.test(key) ? 'out' : 'file.pdf');
      if (/hex/i.test(key)) return '#FF6600';
      if (/^(recordRange|pageRange)$/.test(key)) return '1-2';
      return NASTY;
    case 'number': return 3;
    case 'boolean': return mode !== 'last';
    case 'array':
      if (spec.items?.type === 'array') return [['a', 'b'], ['c', null]];
      return [1, 2, 3, 4].slice(0, /rgb/.test(key) ? 3 : 4);
    case 'object': return { h1: NASTY };
    default: return 'x';
  }
};

const variants = (tool) => {
  const props = tool.inputSchema.properties || {};
  const required = tool.inputSchema.required || [];
  const list = [];
  const build = (keys, mode) => Object.fromEntries(keys.map((k) => [k, sample(k, props[k], mode)]));
  list.push(build(required, 'first'));
  list.push(build(Object.keys(props), 'first'));
  list.push(build(Object.keys(props), 'last'));
  for (const [k, spec] of Object.entries(props)) {
    if (spec.enum) for (const v of spec.enum) list.push({ ...build(required, 'first'), [k]: v });
    if (spec.type === 'boolean') { list.push({ ...build(required, 'first'), [k]: true }); list.push({ ...build(required, 'first'), [k]: false }); }
  }
  return list;
};

const problems = [];
let scripts = 0;
const check = (name, args, script) => {
  scripts++;
  try {
    new vm.Script(script, { filename: name });
  } catch (error) {
    problems.push(`${name} ${JSON.stringify(args).slice(0, 160)}\n      ${error.message}`);
    return;
  }
  const es6 = script.split('\n').find((l) => /=>|^\s*(let|const)\s|\bclass\s+\w+\s*\{/.test(l) && !/^\s*\/\//.test(l) && !/["'].*(=>).*["']/.test(l));
  if (es6) problems.push(`${name}: ES6+ syntax is not supported by ExtendScript: ${es6.trim().slice(0, 100)}`);
  if (/,\s*[\]}]/.test(script.replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, '""'))) {
    problems.push(`${name}: trailing comma (breaks ExtendScript arrays/objects): ${JSON.stringify(args).slice(0, 100)}`);
  }
};

check('wrapper', {}, buildWrapper('1 + 1;', '/tmp/r.txt'));

for (const tool of tools) {
  if (tool.name === 'execute_indesign_code') continue; // runs caller-supplied code by design
  for (const args of variants(tool)) {
    captured = [];
    try {
      await callHandler({ method: 'tools/call', params: { name: tool.name, arguments: args } }, {});
    } catch (error) {
      // Argument combinations may legitimately be rejected (validation); only the scripts matter.
      if (/SyntaxError|is not a function|is not defined|Cannot read prop|undefined \(reading/.test(error.message)) problems.push(`${tool.name}: server-side bug: ${error.message}`);
    }
    for (const script of captured) check(tool.name, args, script);
  }
}

fs.rmSync(TMP, { recursive: true, force: true });
for (const p of schemaProblems) problems.push(`schema: ${p}`);
console.log(`Checked ${scripts} generated scripts for ${tools.length} tools.`);
if (problems.length) {
  const unique = [...new Set(problems)];
  console.log(`\n${unique.length} problem(s):`);
  for (const p of unique) console.log(`  - ${p}`);
  process.exit(1);
}
console.log('All generated scripts are syntactically valid ES3-compatible JavaScript.');
process.exit(0);
