/**
 * Checks that every tool method produces parseable ExtendScript, including
 * the result assignment the driver adds.
 *
 * `node --check index.js` only validates the server, never the generated
 * script text — a broken template would otherwise surface inside InDesign.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import path from 'node:path';
import { runTool, parses, getAutoCapture } from './harness.mjs';

const autoCaptureResult = await getAutoCapture();

// Path arguments stay platform-neutral so the test also runs on macOS.
const TMP = os.tmpdir();

// One argument set; each method destructures what it needs.
const ARGS = {
  content: 'Test content',
  text: 'Test content',
  markdownText: '# Heading\n\nParagraph',
  findText: 'old',
  replaceText: 'new',
  x: 10, y: 10, width: 100, height: 50,
  pageIndex: 0, frameIndex: 0, objectIndex: 0, tableIndex: 0,
  startIndex: 0, endIndex: 5,
  fontSize: 12, fontFamily: 'Helvetica Neue', fontStyle: 'Regular',
  textColor: 'Black', fillColor: 'Paper', strokeColor: 'Black',
  strokeWidth: 1, cornerRadius: 2, transparency: 0,
  alignment: 'LEFT_ALIGN', fitOption: 'PROPORTIONALLY',
  colorModel: 'PROCESS', colorSpace: 'CMYK', colorValues: [0, 0, 0, 100],
  name: 'Test style', styleName: 'Test style', baseStyle: 'Base',
  swatchName: 'Test colour', property: 'fill',
  layerName: 'Test layer', color: 'RED',
  preset: 'A4', orientation: 'Portrait', pages: 1,
  marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20,
  bleed: 3, slug: 0, facingPages: false,
  rows: 2, columns: 2, headerRows: 0, footerRows: 0,
  data: [['a', 'b'], ['c', 'd']],
  tracking: 0, leading: 14, spaceBefore: 0, spaceAfter: 0,
  maxObjects: 10, maxLength: 0, normalizeSpaces: true,
  jpegQuality: 80, resolution: 300, imageFormat: 'PNG', format: 'PNG',
  pageRange: 'all', version: 'EPUB3', profile: 'Base',
  recordRange: 'all', includeImages: true,
  filePath: path.join(TMP, 'test.indd'),
  folderPath: TMP,
  outputFolder: TMP,
  dataSource: path.join(TMP, 'data.csv'),
  imagePath: path.join(TMP, 'image.jpg'),
  confirmDestructive: true,
};

// Every method that builds a script.
const METHODS = [
  'getDocumentInfo', 'createDocument', 'openDocument', 'saveDocument', 'closeDocument',
  'addPage', 'deletePage', 'duplicatePage', 'navigateToPage', 'getSelectedObjects',
  'getTextContent', 'listTextFrames', 'analyzeEmbeddedObjects', 'insertMarkdownText',
  'fixTypographyInSelection', 'findTypographyIssues', 'listGrepSearches',
  'cleanImportedText', 'analyzeTextProblems', 'createTextFrame', 'editTextFrame',
  'findReplaceText', 'placeImage', 'createRectangle', 'createEllipse',
  'createParagraphStyle', 'createCharacterStyle', 'modifyCharacterStyle',
  'modifyParagraphStyle', 'modifyObjectStyle', 'createObjectStyle', 'applyObjectStyle',
  'applyParagraphStyle', 'listStyles', 'createColorSwatch', 'listColorSwatches',
  'applyColor', 'exportPDF', 'exportImages', 'exportEPUB', 'packageDocument',
  'viewDocument', 'createTable', 'populateTable', 'createLayer', 'setActiveLayer',
  'listLayers', 'preflightDocument', 'zoomToPage', 'dataMerge',
];

for (const method of METHODS) {
  test(`${method}: produces parseable ExtendScript`, async () => {
    const { error, script } = await runTool(method, { ...ARGS });

    // A rejected argument is a valid outcome — a programming error is not.
    if (error) {
      assert.ok(
        !/is not a function|Cannot read propert|is not defined/.test(error.message),
        `${method} failed on a programming error: ${error.message}`
      );
      return;
    }
    assert.ok(script, `${method} produced no script`);

    const res = parses(autoCaptureResult(script));
    assert.ok(res.ok, `${method}: script does not parse — ${res.message}`);
  });
}

test('autoCaptureResult assigns a return value', async () => {
  const { script } = await runTool('createDocument', { ...ARGS });
  assert.match(
    autoCaptureResult(script),
    /__result__\s*=/,
    'no __result__ assigned — the tool would return nothing'
  );
});

test('autoCaptureResult handles a multi-line trailing expression', () => {
  const src = [
    'var doc = app.documents.add();',
    '"Created: " + doc.name + ", " +',
    'doc.pages.length + " pages";',
  ].join('\n');
  const wrapped = autoCaptureResult(src);
  const res = parses(wrapped);
  assert.ok(res.ok, 'multi-line trailing expression breaks: ' + res.message);
  assert.match(wrapped, /__result__ = "Created: "/);
});

test('autoCaptureResult leaves scripts that set __result__ alone', () => {
  const src = '__result__ = "fixed";';
  assert.equal(autoCaptureResult(src), src);
});
