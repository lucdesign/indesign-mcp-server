// Seeing and measuring (render_preview, measure_text) and vector graphics placement
// (place_graphic, create_graphic_from_svg, create_cmyk_pdf_shape).

import fs from 'fs';
import path from 'path';
import { jsStr, jsJson, jsNum } from '../lib/script-runtime.js';
import { svgPathToSubpaths, fitSubpaths, subpathBounds } from '../lib/svg-path.js';

const has = (v) => v !== undefined && v !== null;
const rand = () => Math.random().toString(36).slice(2);

const VECTOR_EXTENSIONS = ['.svg', '.ai', '.eps', '.pdf'];
const CROPS = { pdf: 'CROP_PDF', art: 'CROP_ART', trim: 'CROP_TRIM', bleed: 'CROP_BLEED', media: 'CROP_MEDIA', content_visible: 'CROP_CONTENT_VISIBLE_LAYERS', content_all: 'CROP_CONTENT_ALL_LAYERS' };
const FITS = { PROPORTIONALLY: 'FitOptions.PROPORTIONALLY', FILL_PROPORTIONALLY: 'FitOptions.FILL_PROPORTIONALLY', CONTENT_TO_FRAME: 'FitOptions.CONTENT_TO_FRAME', FRAME_TO_CONTENT: 'FitOptions.FRAME_TO_CONTENT', CENTER_CONTENT: 'FitOptions.CENTER_CONTENT', NONE: null };

// SVG markup is parsed by InDesign; refuse anything that could make it read other files or
// fetch URLs (external references, entities, scripts).
export function validateSvgMarkup(svg) {
  if (typeof svg !== 'string' || !/<svg[\s>]/i.test(svg)) throw new Error('svg must be SVG markup starting with an <svg> element');
  if (svg.length > 2 * 1024 * 1024) throw new Error('svg is too large (2 MB maximum)');
  const forbidden = [
    [/<!DOCTYPE/i, 'a DOCTYPE'], [/<!ENTITY/i, 'entity declarations'], [/<script[\s>]/i, '<script>'], [/<foreignObject[\s>]/i, '<foreignObject>'],
    [/<\?xml-stylesheet/i, 'xml-stylesheet'], [/@import/i, '@import'], [/\bon[a-z]+\s*=/i, 'event handler attributes'],
  ];
  for (const [re, what] of forbidden) if (re.test(svg)) throw new Error(`The SVG contains ${what}, which is not allowed`);
  for (const m of svg.matchAll(/(?:xlink:)?href\s*=\s*["']([^"']*)["']/gi)) {
    if (!/^(#|data:image\/(png|jpeg|jpg|gif|svg\+xml);base64,)/i.test(m[1].trim())) throw new Error(`The SVG references an external resource (${m[1].slice(0, 40)}); only #fragment and embedded data:image URIs are allowed`);
  }
  for (const m of svg.matchAll(/url\(\s*["']?([^)"']*)["']?\s*\)/gi)) {
    if (!/^(#|data:image\/)/i.test(m[1].trim())) throw new Error(`The SVG uses an external url(${m[1].slice(0, 40)}); only url(#id) is allowed`);
  }
}

// Minimal single-page PDF with exact DeviceCMYK colours (no compression, no fonts)
export function buildCmykPdf({ width, height, shapes }) {
  const n = (v) => (Math.round(v * 1000) / 1000).toString();
  const k = (c) => c.map((v) => n(v / 100)).join(' ');
  let content = '';
  for (const shape of shapes) {
    const y = (v) => height - v;
    let ops = '';
    for (const sp of shape.subpaths) {
      const pts = sp.points;
      ops += `${n(pts[0].x)} ${n(y(pts[0].y))} m\n`;
      const count = sp.closed ? pts.length : pts.length - 1;
      for (let i = 0; i < count; i++) {
        const a = pts[i], b = pts[(i + 1) % pts.length];
        ops += `${n(a.right.x)} ${n(y(a.right.y))} ${n(b.left.x)} ${n(y(b.left.y))} ${n(b.x)} ${n(y(b.y))} c\n`;
      }
      if (sp.closed) ops += 'h\n';
    }
    let paint = 'n';
    if (shape.fill && shape.stroke) paint = 'B*';
    else if (shape.fill) paint = 'f*';
    else if (shape.stroke) paint = 'S';
    content += 'q\n';
    if (shape.fill) content += `${k(shape.fill)} k\n`;
    if (shape.stroke) content += `${k(shape.stroke)} K\n${n(shape.strokeWidth || 1)} w\n`;
    content += ops + paint + '\nQ\n';
  }
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${n(width)} ${n(height)}] /Contents 4 0 R /Resources << >> >>`,
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}endstream`,
  ];
  let pdf = '%PDF-1.4\n';
  const offsets = [];
  objects.forEach((body, i) => { offsets.push(Buffer.byteLength(pdf)); pdf += `${i + 1} 0 obj\n${body}\nendobj\n`; });
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` + offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('');
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return pdf;
}

export function createGraphicsTools(server) {
  const fmt = (result, operation) => server.formatResponse(result, operation);
  const run = (script) => server.executeInDesignScript(script);
  const frameExpr = (args) => server.frameExpr(args);

  const FRAME_PROPS = {
    frameId: { type: 'number', description: 'Stable id of the text frame (preferred over frameIndex)' },
    frameIndex: { type: 'number', description: 'Index in page.textFrames (0 = frontmost)' },
    pageIndex: { type: 'number', description: 'Page index (0-based), used with frameIndex', default: 0 },
  };
  const PLACE_PROPS = {
    x: { type: 'number', description: 'X position in mm', default: 10 },
    y: { type: 'number', description: 'Y position in mm', default: 10 },
    width: { type: 'number', description: 'Frame width in mm' },
    height: { type: 'number', description: 'Frame height in mm' },
    pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
    fitOption: { type: 'string', enum: Object.keys(FITS), description: 'How the graphic fits the frame (with width and height): PROPORTIONALLY = fit inside, FILL_PROPORTIONALLY = crop to fill', default: 'PROPORTIONALLY' },
    embed: { type: 'boolean', description: 'Unlink and embed the graphic in the document', default: false },
  };

  const definitions = [
    {
      name: 'render_preview',
      description: 'Render one page (or one object) to PNG and return it directly as an image, so you can see the result without exporting to a folder. Options: pageIndex, objectId (renders just that object), resolution in dpi (default 72), includeBleed. The longest side is limited to 4000 px.',
      inputSchema: {
        type: 'object',
        properties: {
          pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
          objectId: { type: 'number', description: 'Render only this object instead of a page' },
          resolution: { type: 'number', description: 'Resolution in dpi (10-600)', default: 72 },
          includeBleed: { type: 'boolean', description: 'Include the document bleed (page previews)', default: false },
        },
      },
    },
    {
      name: 'measure_text',
      description: 'Measure a text frame without exporting: line count, the text of every line, the width of every line in mm, the frame width and the overflow state.',
      inputSchema: { type: 'object', properties: { ...FRAME_PROPS } },
    },
    {
      name: 'place_graphic',
      description: 'Place a vector graphic (.svg, .ai, .eps, .pdf) into a new frame (mm). Give width and height for a fixed frame, only one for the aspect ratio, or neither for the natural size. PDF/AI: pdfPage (default 1) and crop (pdf, art, trim, bleed, media, content_visible, content_all). Returns the frame id and whether the result is linked; embed: true unlinks and embeds it.',
      inputSchema: {
        type: 'object',
        properties: {
          filePath: { type: 'string', description: 'Path of the .svg, .ai, .eps or .pdf file' },
          ...PLACE_PROPS,
          pdfPage: { type: 'number', description: 'PDF / AI: page number to place (1-based)', default: 1 },
          crop: { type: 'string', enum: Object.keys(CROPS), description: 'PDF / AI: crop to', default: 'pdf' },
        },
        required: ['filePath'],
      },
    },
    {
      name: 'create_graphic_from_svg',
      description: 'Place SVG markup (a string) as a vector graphic: the server writes it to a temp .svg under $TMPDIR/indesign-mcp/svg/ (or to saveTo) and places it like place_graphic. Use it as the fallback when a native path cannot express something (text as outlines, complex filters). LIMITS: the result is NOT editable as paths in InDesign, and the RGB colours of the SVG are converted to CMYK only at output, through the CMYK profile of the document (accurate, but you cannot dictate exact CMYK numbers). For exact CMYK values use create_path_from_svg with swatches, or create_cmyk_pdf_shape. With embed: true the graphic is embedded and the temp file deleted; otherwise it stays linked to the file (temp files are removed after 3 days). External references, scripts, entities and external url() in the SVG are rejected.',
      inputSchema: {
        type: 'object',
        properties: {
          svg: { type: 'string', description: 'Complete SVG markup, starting with <svg ...>' },
          ...PLACE_PROPS,
          saveTo: { type: 'string', description: 'Optional: keep the .svg at this path (inside the allowed folders) instead of a temp file' },
        },
        required: ['svg'],
      },
    },
    {
      name: 'create_cmyk_pdf_shape',
      description: 'Generate a small vector PDF with EXACT CMYK fill/stroke values from SVG path data and place it (like place_graphic). Use it for logos or multi-colour shapes when you need exact brand CMYK numbers (the SVG route converts RGB through the profile) and native paths (create_path_from_svg with one swatch per colour) do not fit. The result is a linked or embedded graphic, not an editable path.',
      inputSchema: {
        type: 'object',
        properties: {
          shapes: {
            type: 'array',
            description: 'Shapes painted in order: { d: SVG path data, fill: [C,M,Y,K] 0-100, stroke: [C,M,Y,K], strokeWidth: pt in viewBox units }',
            items: { type: 'object', properties: { d: { type: 'string' }, fill: { type: 'array', items: { type: 'number' } }, stroke: { type: 'array', items: { type: 'number' } }, strokeWidth: { type: 'number' } }, required: ['d'] },
          },
          viewBox: { type: 'string', description: 'SVG viewBox "minX minY width height" of the path data (default: bounds of all shapes)' },
          ...PLACE_PROPS,
          embed: { type: 'boolean', description: 'Unlink and embed the generated PDF (recommended: the temp file is removed afterwards)', default: true },
          saveTo: { type: 'string', description: 'Optional: keep the generated .pdf at this path instead of a temp file' },
        },
        required: ['shapes'],
      },
    },
  ];

  const handlers = {};

  // ---- render_preview ---------------------------------------------------------------------
  const renderToPng = async (args) => {
    const dpi = has(args.resolution) ? jsNum(args.resolution, 'resolution') : 72;
    if (dpi < 10 || dpi > 600) throw new Error('resolution must be 10-600 dpi');
    const file = path.join(server.scriptDir, `preview_${Date.now()}_${rand()}.png`);
    const pageIdx = has(args.pageIndex) ? jsNum(args.pageIndex, 'pageIndex') : 0;
    const bleed = args.includeBleed === true;
    const script = `
      var doc = __requireDoc();
      var out = File(${jsStr(file)});
      var prefs = app.pngExportPreferences;
      prefs.exportResolution = ${dpi};
      prefs.useDocumentBleeds = ${bleed ? 'true' : 'false'};
      var label;
      ${has(args.objectId) ? `
        var item = __resolve(doc, ${jsNum(args.objectId, 'objectId')}, null, null);
        item.exportFile(ExportFormat.PNG_FORMAT, out, false);
        label = "object " + item.id + " (" + item.constructor.name + ")";
      ` : `
        var page = __page(doc, ${pageIdx});
        var dp = doc.documentPreferences;
        var wmm = dp.pageWidth + (${bleed ? 'true' : 'false'} ? dp.documentBleedInsideOrLeftOffset + dp.documentBleedOutsideOrRightOffset : 0);
        var hmm = dp.pageHeight + (${bleed ? 'true' : 'false'} ? dp.documentBleedTopOffset + dp.documentBleedBottomOffset : 0);
        var longest = Math.max(wmm, hmm) / 25.4 * ${dpi};
        if (longest > 4000) throw new Error("The preview would be " + Math.round(longest) + " px on its longest side (maximum 4000). Lower the resolution.");
        // pageString is only honoured in EXPORT_RANGE mode
        prefs.pngExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
        prefs.exportingSpread = false;
        prefs.pageString = "+" + (page.documentOffset + 1);   // absolute position: page names can repeat or be renumbered
        doc.exportFile(ExportFormat.PNG_FORMAT, out, false);
        label = "page " + (page.documentOffset + 1);
      `}
      if (!out.exists) throw new Error("The preview export did not create a file");
      label;
    `;
    let label;
    try {
      label = await run(script);
      const data = fs.readFileSync(file);
      const w = data.readUInt32BE(16), h = data.readUInt32BE(20);
      return { label, data, w, h, dpi };
    } finally {
      fs.rmSync(file, { force: true });
    }
  };
  server.renderToPng = renderToPng;

  handlers.render_preview = async (args) => {
    const { label, data, w, h, dpi } = await renderToPng(args);
    return {
      content: [
        { type: 'text', text: `Preview of ${label}: ${w} x ${h} px at ${dpi} dpi` },
        { type: 'image', data: data.toString('base64'), mimeType: 'image/png' },
      ],
    };
  };

  // ---- measure_text -----------------------------------------------------------------------
  handlers.measure_text = async (args) => {
    const script = `
      var doc = __requireDoc();
      var frame = ${frameExpr(args)};
      var b = frame.geometricBounds;
      var out = ["=== MEASURE TEXT: frame id=" + frame.id + " ==="];
      var inset = frame.textFramePreferences.insetSpacing;
      var insets = (inset instanceof Array) ? inset : [inset, inset, inset, inset];
      var textWidth = (b[3] - b[1]) - insets[1] - insets[3];
      out.push("frame width=" + __r(b[3] - b[1]) + " mm, text area width=" + __r(textWidth) + " mm (columns: " + frame.textFramePreferences.textColumnCount + ")");
      out.push(__overflowNote(frame));
      var longest = 0, count = frame.lines.length;
      for (var i = 0; i < count; i++) {
        var ln = frame.lines[i];
        var k = ln.characters.length - 1;
        // the visible end of the line: skip trailing spaces and line/paragraph break characters
        while (k > 0) {
          var c = ln.characters[k].contents;
          if (typeof c !== "string" || c === " " || c === "\\r" || c === "\\n" || c === "\\u00A0") k--; else break;
        }
        var w = ln.characters[k].endHorizontalOffset - ln.characters[0].horizontalOffset;
        if (w > longest) longest = w;
        var text = String(ln.contents).replace(/[\\r\\n]+$/, "");
        out.push("[" + i + "] width=" + __r(w) + " mm baseline=" + __r(ln.baseline) + " mm: " + JSON_STR(text));
      }
      out.push("longest line=" + __r(longest) + " mm");
      out.join("\\n");
    `;
    return fmt(await run(script), 'Measure Text');
  };

  // ---- place_graphic ----------------------------------------------------------------------
  // trusted: the path was produced by the server itself (temp file); user paths are validated first.
  const placeGraphic = async (args, trustedPath) => {
    const file = trustedPath || server.validateFilePath(args.filePath);
    const ext = path.extname(file).toLowerCase();
    if (!VECTOR_EXTENSIONS.includes(ext)) throw new Error(`Unsupported file type ${ext || '(none)'}: place_graphic handles ${VECTOR_EXTENSIONS.join(', ')} (use place_image for bitmaps)`);
    const fit = has(args.fitOption) ? args.fitOption : 'PROPORTIONALLY';
    if (!(fit in FITS)) throw new Error(`Invalid fitOption: ${fit}`);
    const crop = has(args.crop) ? args.crop : 'pdf';
    if (!CROPS[crop]) throw new Error(`Invalid crop: ${crop}`);
    const pdfPage = has(args.pdfPage) ? Math.floor(jsNum(args.pdfPage, 'pdfPage')) : 1;
    if (pdfPage < 1) throw new Error('pdfPage must be >= 1');
    const nx = has(args.x) ? jsNum(args.x, 'x') : 10, ny = has(args.y) ? jsNum(args.y, 'y') : 10;
    const w = has(args.width) ? jsNum(args.width, 'width') : null, h = has(args.height) ? jsNum(args.height, 'height') : null;
    const usesPdfPrefs = ext === '.pdf' || ext === '.ai';
    const embed = args.embed === true;

    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${has(args.pageIndex) ? jsNum(args.pageIndex, 'pageIndex') : 0});
      var file = File(${jsStr(file)});
      if (!file.exists) throw new Error("File not found: " + file.fsName);
      var rect = page.rectangles.add();
      var pp = app.pdfPlacePreferences, oldPage = null, oldCrop = null;
      try {
        ${usesPdfPrefs ? `
          oldPage = pp.pageNumber; oldCrop = pp.pdfCrop;
          pp.pageNumber = ${pdfPage};
          pp.pdfCrop = PDFCrop.${CROPS[crop]};
        ` : ''}
        var fw = ${w !== null ? w : (h !== null ? h : 50)}, fh = ${h !== null ? h : (w !== null ? w : 50)};
        rect.geometricBounds = [${ny}, ${nx}, ${ny} + fh, ${nx} + fw];
        rect.place(file);
        ${usesPdfPrefs ? `
          // InDesign silently falls back to page 1 when the file has fewer pages
          var placedPage = rect.graphics[0].pdfAttributes.pageNumber;
          if (placedPage !== ${pdfPage}) throw new Error("The file has no page ${pdfPage}: InDesign placed page " + placedPage + " instead (the file has fewer pages)");
        ` : ''}
        ${(w !== null) !== (h !== null) || (w === null && h === null) ? `
          rect.fit(FitOptions.FRAME_TO_CONTENT);
          var nb = rect.geometricBounds, iw = nb[3] - nb[1], ih = nb[2] - nb[0];
          ${w !== null ? `rect.geometricBounds = [${ny}, ${nx}, ${ny} + ${w} * ih / iw, ${nx} + ${w}]; rect.fit(FitOptions.CONTENT_TO_FRAME);`
            : h !== null ? `rect.geometricBounds = [${ny}, ${nx}, ${ny} + ${h}, ${nx} + ${h} * iw / ih]; rect.fit(FitOptions.CONTENT_TO_FRAME);`
            : `rect.geometricBounds = [${ny}, ${nx}, ${ny} + ih, ${nx} + iw]; rect.fit(FitOptions.CONTENT_TO_FRAME);`}
        ` : (FITS[fit] ? `rect.fit(${FITS[fit]});` : '')}
        ${embed ? `
          rect.graphics[0].itemLink.unlink();
        ` : ''}
      } catch (e) {
        try { rect.remove(); } catch (e2) {}
        throw e;
      } finally {
        ${usesPdfPrefs ? 'try { pp.pageNumber = oldPage; pp.pdfCrop = oldCrop; } catch (e3) {}' : ''}
      }
      var g = rect.graphics[0];
      var linked = false, status = "embedded";
      try {
        var lk = g.itemLink;
        // an embedded graphic keeps a link object whose status is LINK_EMBEDDED
        if (lk && lk.isValid && String(lk.status) !== "LINK_EMBEDDED") { linked = true; status = String(lk.status); }
      } catch (e4) {}
      var pdfInfo = "";
      ${usesPdfPrefs ? 'try { pdfInfo = " pdfPage=" + g.pdfAttributes.pageNumber; } catch (e5) {}' : ''}
      "Graphic placed: " + __describe(rect) + " type=" + g.imageTypeName + " linked=" + linked + " (" + status + ")" + pdfInfo;
    `;
    return run(script);
  };

  handlers.place_graphic = async (args) => fmt(await placeGraphic(args), 'Place Graphic');

  const tempFile = (subdir, ext) => {
    const dir = path.join(server.scriptDir, subdir);
    fs.mkdirSync(dir, { recursive: true });
    return path.join(dir, `${subdir}_${Date.now()}_${rand()}${ext}`);
  };
  const target = (args, subdir, ext) => (has(args.saveTo) ? server.validateFilePath(args.saveTo) : tempFile(subdir, ext));

  handlers.create_graphic_from_svg = async (args) => {
    validateSvgMarkup(args.svg);
    const file = target(args, 'svg', '.svg');
    if (has(args.saveTo) && path.extname(file).toLowerCase() !== '.svg') throw new Error('saveTo must end with .svg');
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, args.svg, 'utf8');
    let ok = false;
    try {
      const result = await placeGraphic(args, file);
      ok = true;
      return fmt(`${result}\nsource file: ${args.embed === true && !has(args.saveTo) ? '(deleted, graphic is embedded)' : file}\nNote: the graphic is not editable as paths; RGB colours are converted to the document colour space.`, 'Create Graphic From SVG');
    } finally {
      // an embedded graphic no longer needs its file; a failed placement leaves nothing behind
      if ((args.embed === true && ok && !has(args.saveTo)) || (!ok && !has(args.saveTo))) fs.rmSync(file, { force: true });
    }
  };

  handlers.create_cmyk_pdf_shape = async (args) => {
    if (!Array.isArray(args.shapes) || args.shapes.length === 0) throw new Error('shapes must list at least one shape');
    const cmyk = (v, label) => {
      if (!has(v)) return null;
      if (!Array.isArray(v) || v.length !== 4 || v.some((n) => !Number.isFinite(Number(n)) || n < 0 || n > 100)) throw new Error(`${label} must be [C,M,Y,K] with values 0-100`);
      return v.map(Number);
    };
    let viewBox;
    if (has(args.viewBox)) {
      viewBox = String(args.viewBox).trim().split(/[\s,]+/).map(Number);
      if (viewBox.length !== 4 || viewBox.some((n) => !Number.isFinite(n)) || viewBox[2] <= 0 || viewBox[3] <= 0) throw new Error('viewBox must be "minX minY width height"');
    }
    const parsed = args.shapes.map((s, i) => {
      if (typeof s.d !== 'string') throw new Error(`shapes[${i}].d (SVG path data) is required`);
      const fill = cmyk(s.fill, `shapes[${i}].fill`), stroke = cmyk(s.stroke, `shapes[${i}].stroke`);
      if (!fill && !stroke) throw new Error(`shapes[${i}] needs a fill and/or a stroke`);
      return { subpaths: svgPathToSubpaths(s.d).subpaths, fill, stroke, strokeWidth: has(s.strokeWidth) ? jsNum(s.strokeWidth, `shapes[${i}].strokeWidth`) : 1 };
    });
    const box = viewBox ? { x: viewBox[0], y: viewBox[1], width: viewBox[2], height: viewBox[3] } : subpathBounds(parsed.flatMap((s) => s.subpaths));
    if (!(box.width > 0 && box.height > 0)) throw new Error('The shapes have no width or height; pass a viewBox');
    // shift into the page origin: the PDF page is exactly the viewBox
    const shift = (p) => ({ x: p.x - box.x, y: p.y - box.y });
    const shapes = parsed.map((s) => ({
      ...s,
      subpaths: s.subpaths.map((sp) => ({ closed: sp.closed, points: sp.points.map((pt) => ({ ...shift(pt), left: shift(pt.left), right: shift(pt.right) })) })),
    }));
    const pdf = buildCmykPdf({ width: box.width, height: box.height, shapes });
    const file = target(args, 'pdf', '.pdf');
    if (has(args.saveTo) && path.extname(file).toLowerCase() !== '.pdf') throw new Error('saveTo must end with .pdf');
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, pdf, 'latin1');
    const embed = args.embed !== false;
    let ok = false;
    try {
      const result = await placeGraphic({ ...args, embed, crop: 'media' }, file);
      ok = true;
      return fmt(`${result}\nsource file: ${embed && !has(args.saveTo) ? '(deleted, graphic is embedded)' : file}`, 'Create CMYK PDF Shape');
    } finally {
      if ((embed && ok && !has(args.saveTo)) || (!ok && !has(args.saveTo))) fs.rmSync(file, { force: true });
    }
  };

  return { definitions, handlers };
}
