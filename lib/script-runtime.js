// Shared helpers for generating and running ExtendScript in InDesign.
//
// Design rules (see README, "How scripts run"):
//  * Every tool builds plain ExtendScript. It is never embedded in AppleScript source.
//    The script is written to a temp .jsx and started with
//    `do script (POSIX file "...") language javascript`, so quoting inside the
//    script can never collide with AppleScript syntax.
//  * Values from tool arguments are embedded only through jsStr / jsJson / jsNum / jsEnum.
//  * Errors are reported with numeric InDesign error codes. Nothing depends on the
//    (localised) text of InDesign's messages.

// Escape non-ASCII so the generated .jsx is pure ASCII (ExtendScript's file
// encoding is not reliable), and keep U+2028/2029 from ending a string literal.
function asciiEscape(str) {
  return str.replace(/[\u0080-￿]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
}

// Safe ExtendScript string literal for any value
export function jsStr(value) {
  return asciiEscape(JSON.stringify(String(value)));
}

// Safe ExtendScript literal for JSON-serialisable data (objects, arrays, numbers, ...)
export function jsJson(value) {
  return asciiEscape(JSON.stringify(value));
}

// Finite number, embedded as a bare literal
export function jsNum(value, label) {
  const n = Number(value);
  if (value === null || value === '' || !Number.isFinite(n)) {
    throw new Error(`Invalid ${label}: ${value}`);
  }
  return n;
}

// Identifier (e.g. an enum member name) that will be spliced into code
export function jsEnum(value, label) {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(String(value))) {
    throw new Error(`Invalid ${label}: ${value}`);
  }
  return String(value);
}

// Escape a value for use inside an AppleScript string literal
export function appleScriptStr(value) {
  return '"' + String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
}

// Turn an osascript failure into a short, locale-independent message.
// Uses the numeric error code (last "(-123)" in the text), never the wording.
export function describeOsascriptError(error) {
  const raw = String((error && (error.stderr || error.message)) || error || '');
  if (error && (error.killed || error.signal === 'SIGTERM' || /ETIMEDOUT/.test(raw))) {
    return 'InDesign did not respond in time (is a dialog open in InDesign?)';
  }
  const codeMatch = raw.match(/\((-?\d+)\)\s*$/m);
  const code = codeMatch ? Number(codeMatch[1]) : null;
  if (code === -600 || code === -10810 || code === -1728) {
    return `InDesign is not running or not reachable (AppleScript error ${code})`;
  }
  // Drop file paths, the "Command failed" prefix and the "file:line:col: execution error:" prefix
  let msg = raw
    .replace(/^Command failed:.*\n?/m, '')
    .replace(/\/[^\s"]+\.(scpt|jsx)(:\d+:\d+)?:?/g, '')
    .replace(/execution error:?/i, '')
    .trim();
  // Remove the localised “Adobe InDesign 2026” hat einen Fehler erhalten: prefix if present
  msg = msg.replace(/^[^:]*[”“"»][^:]*:\s*/, '');
  return msg || `AppleScript error${code !== null ? ' ' + code : ''}`;
}

// ExtendScript prepended to every script. Defines shared helpers and switches the
// active document to millimetres with a per-page ruler origin, so every tool works in
// the same coordinate system. The previous settings are restored afterwards.
export const PRELUDE = `
function __r(n) { return Math.round(n * 1000) / 1000; }
function __has(v) { return v !== null && v !== undefined; }

// Find any page item (also nested in groups) by its stable id
function __findById(doc, id) {
  try {
    var it = doc.pageItems.itemByID(id);
    // itemByID returns a generic PageItem; getElements() gives the concrete type (TextFrame, Rectangle, ...)
    if (it.isValid) return it.getElements()[0];
  } catch (e) {}
  var all = doc.allPageItems;
  for (var i = 0; i < all.length; i++) {
    if (all[i].id === id) return all[i];
  }
  return null;
}

// True when the document has a file on disk. fullName throws for a document that was never saved, and is the
// reliable test (doc.saved describes the save state, not whether a file exists).
function __hasFile(doc) {
  try { return doc.fullName !== null && doc.fullName !== undefined; } catch (e) { return false; }
}

function __requireDoc() {
  if (app.documents.length === 0) throw new Error("No document open. Please create or open a document first.");
  return app.activeDocument;
}

function __page(doc, pageIndex) {
  var p = __has(pageIndex) ? pageIndex : 0;
  if (p < 0 || p >= doc.pages.length) {
    throw new Error("Invalid page index: " + p + " (document has " + doc.pages.length + " pages, 0-based)");
  }
  return doc.pages[p];
}

// Resolve an object by id, or by page + index within a collection (default: all page items)
function __resolve(doc, id, pageIndex, index, coll) {
  if (__has(id)) {
    var found = __findById(doc, id);
    if (!found) throw new Error("No object with id " + id + " in the active document");
    return found;
  }
  if (!__has(index)) throw new Error("Specify the object id, or pageIndex plus index");
  var page = __page(doc, pageIndex);
  var items = coll ? page[coll] : page.allPageItems;
  if (index < 0 || index >= items.length) {
    throw new Error("Invalid index: " + index + " (page " + (page.documentOffset + 1) + " has " + items.length + (coll ? " " + coll : " page items") + ")");
  }
  return items[index];
}

function __textFrame(doc, id, pageIndex, frameIndex) {
  var f = __resolve(doc, id, pageIndex, frameIndex, "textFrames");
  if (f.constructor.name !== "TextFrame") throw new Error("Object " + f.id + " is a " + f.constructor.name + ", not a text frame");
  return f;
}

function __swatch(doc, name) {
  var s = doc.swatches.itemByName(name);
  if (!s.isValid) throw new Error("Swatch not found: " + name);
  return s;
}

function __pstyle(doc, name) {
  var s = doc.paragraphStyles.itemByName(name);
  if (!s.isValid) throw new Error("Paragraph style not found: " + name);
  return s;
}

function __ostyle(doc, name) {
  var s = doc.objectStyles.itemByName(name);
  if (!s.isValid) throw new Error("Object style not found: " + name);
  return s;
}

// GREP find/change on a story (keeps formatting). Returns the number of changes.
function __grep(story, find, change) {
  app.findGrepPreferences = NothingEnum.nothing;
  app.changeGrepPreferences = NothingEnum.nothing;
  app.findGrepPreferences.findWhat = find;
  app.changeGrepPreferences.changeTo = change;
  var n = 0;
  try {
    n = story.changeGrep().length;
  } finally {
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
  }
  return n;
}

function __cstyle(doc, name) {
  var s = doc.characterStyles.itemByName(name);
  if (!s.isValid) throw new Error("Character style not found: " + name);
  return s;
}

// Colour-managed conversion through InDesign's colour engine (the document's RGB and CMYK profiles):
// a temporary swatch is created, its colour space is switched, and the swatch is removed again.
function __convertColor(doc, values, from, to) {
  var c = doc.colors.add({ model: ColorModel.PROCESS, space: ColorSpace[from], colorValue: values });
  try {
    c.space = ColorSpace[to];
    var v = c.colorValue, out = [];
    for (var i = 0; i < v.length; i++) out.push(Math.round(v[i] * 10) / 10);
    return out;
  } finally {
    try { c.remove(); } catch (e) {}
  }
}
function __cmsCmyk(doc, rgb) { return __convertColor(doc, rgb, "RGB", "CMYK"); }
function __cmsRgb(doc, cmyk) { return __convertColor(doc, cmyk, "CMYK", "RGB"); }

// Set font family (+ optional style) on a text object or style. Returns the applied font name.
function __fontOk(f) {
  try { return f.isValid && f.status === FontStatus.INSTALLED; } catch (e) { return false; }
}

// Set font family (+ optional style) on a text object or style. Returns the applied font name.
function __applyFont(target, family, style) {
  var full = style ? family + "\\t" + style : family;
  var font = app.fonts.itemByName(full);
  if (!__fontOk(font)) {
    var byFamily = app.fonts.itemByName(family);
    // family-only lookup is a fallback for the default style only; a misspelled style must be an error
    if (__fontOk(byFamily) && (!style || style === "Regular")) {
      font = byFamily;
    } else {
      // "Helvetica Neue Bold" given as the family: try the last one or two words as the style
      font = null;
      var words = String(family).split(" ");
      for (var k = 1; k <= 2 && !font && words.length > k; k++) {
        var fam2 = words.slice(0, words.length - k).join(" "), sty2 = words.slice(words.length - k).join(" ");
        var cand = app.fonts.itemByName(fam2 + "\\t" + sty2);
        if (__fontOk(cand)) font = cand;
      }
      if (!font) throw new Error(__fontHelp(family, style));
    }
  }
  target.appliedFont = font;
  return font.name.replace("\\t", " ");
}

// Installed font families with their styles, built from bulk property arrays (fast even with 1000+ fonts)
function __fontFamilies() {
  var fam = app.fonts.everyItem().fontFamily, sty = app.fonts.everyItem().fontStyleName, st = app.fonts.everyItem().status;
  var map = {}, order = [];
  for (var i = 0; i < fam.length; i++) {
    if (String(st[i]) !== "INSTALLED") continue;
    if (!map[fam[i]]) { map[fam[i]] = []; order.push(fam[i]); }
    map[fam[i]].push(sty[i]);
  }
  order.sort();
  return { order: order, map: map };
}

// Error text for a font that is not installed, with suggestions
function __fontHelp(family, style) {
  var all = __fontFamilies(), fl = String(family).toLowerCase();
  if (all.map[family]) {
    return "Font family " + family + " has no style " + style + ". Available styles: " + all.map[family].join(", ");
  }
  var first = fl.split(" ")[0], hits = [], used = {};
  // best matches first: families starting with what was typed, then containing it, then sharing its first word
  for (var pass = 0; pass < 3 && hits.length < 5; pass++) {
    for (var i = 0; i < all.order.length && hits.length < 5; i++) {
      var f = all.order[i], lf = f.toLowerCase();
      if (used[f]) continue;
      var ok;
      // (plain ifs: ExtendScript mis-parses chained ternaries with ===)
      if (pass === 0) { ok = lf.indexOf(fl) === 0; }
      else if (pass === 1) { ok = lf.indexOf(fl) !== -1 || fl.indexOf(lf) !== -1; }
      else { ok = lf.indexOf(first) === 0; }
      if (ok) {
        used[f] = true;
        hits.push(f + " [" + all.map[f].slice(0, 8).join(", ") + (all.map[f].length > 8 ? ", ..." : "") + "]");
      }
    }
  }
  return "Font not installed: " + family + (style ? " (" + style + ")" : "") + (hits.length ? ". Did you mean: " + hits.join("; ") : ". Use list_fonts to see the installed families");
}

// Usage of every swatch name by objects (fill/stroke) and styles: name -> { fills, strokes, styles, ids }
function __usageMap(doc) {
  var map = {};
  function slot(n) { if (!map[n]) map[n] = { fills: 0, strokes: 0, styles: 0, ids: [] }; return map[n]; }
  var items = doc.allPageItems;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    try { var f = it.fillColor; if (f && f.isValid) { var a = slot(f.name); a.fills++; if (a.ids.length < 10) a.ids.push(it.id); } } catch (e1) {}
    try { var k = it.strokeColor; if (k && k.isValid) { var b = slot(k.name); b.strokes++; if (b.ids.length < 10 && b.ids[b.ids.length - 1] !== it.id) b.ids.push(it.id); } } catch (e2) {}
  }
  var colls = ["paragraphStyles", "characterStyles", "objectStyles"];
  for (var c = 0; c < colls.length; c++) {
    var coll = doc[colls[c]];
    for (var j = 0; j < coll.length; j++) {
      try { var sf = coll[j].fillColor; if (sf && sf.isValid) slot(sf.name).styles++; } catch (e3) {}
      try { var ss = coll[j].strokeColor; if (ss && ss.isValid) slot(ss.name).styles++; } catch (e4) {}
    }
  }
  return map;
}

function __describe(it) {
  var t = it.constructor.name;
  var s = t + " id=" + it.id;
  try { s += " page=" + (it.parentPage ? (it.parentPage.documentOffset + 1) : "pasteboard"); } catch (e) {}
  try {
    var b = it.geometricBounds;
    s += " x=" + __r(b[1]) + " y=" + __r(b[0]) + " w=" + __r(b[3] - b[1]) + " h=" + __r(b[2] - b[0]);
  } catch (e) {}
  try { s += " rotation=" + __r(it.rotationAngle); } catch (e) {}
  try { if (it.fillColor && it.fillColor.isValid) s += " fill=" + it.fillColor.name; } catch (e) {}
  try { if (t === "TextFrame") s += " overflows=" + it.overflows + " text=" + JSON_STR(it.contents.substring(0, 40)); } catch (e) {}
  try { if (it.allGraphics && it.allGraphics.length > 0) s += " image=" + it.allGraphics[0].itemLink.name; } catch (e) {}
  try { if (it.label) s += " label=" + it.label; } catch (e) {}
  return s;
}

// Minimal JSON string quoting (ExtendScript has no JSON object)
function JSON_STR(str) {
  return '"' + String(str).replace(/\\\\/g, "\\\\\\\\").replace(/"/g, '\\\\"').replace(/\\r/g, "\\\\r").replace(/\\n/g, "\\\\n") + '"';
}

// ---- vector paths -------------------------------------------------------------------------
// A path is described as a list of { x, y, lx, ly, rx, ry, type } (anchor, left/right handle, type).
function __readPath(path) {
  var out = [];
  for (var i = 0; i < path.pathPoints.length; i++) {
    var p = path.pathPoints[i];
    out.push({ x: p.anchor[0], y: p.anchor[1], lx: p.leftDirection[0], ly: p.leftDirection[1],
               rx: p.rightDirection[0], ry: p.rightDirection[1], type: String(p.pointType) });
  }
  return out;
}

// Write points to a path. Types are set before the handles because changing a point's type can move its handles.
function __writePath(path, pts, closed) {
  var ep = [];
  for (var i = 0; i < pts.length; i++) {
    ep.push([[pts[i].lx, pts[i].ly], [pts[i].x, pts[i].y], [pts[i].rx, pts[i].ry]]);
  }
  path.entirePath = ep;
  path.pathType = closed ? PathType.CLOSED_PATH : PathType.OPEN_PATH;
  for (var k = 0; k < pts.length; k++) {
    var pp = path.pathPoints[k];
    if (pts[k].type && pts[k].type !== "auto") pp.pointType = PointType[pts[k].type];
    pp.leftDirection = [pts[k].lx, pts[k].ly];
    pp.rightDirection = [pts[k].rx, pts[k].ry];
  }
}

function __fmtPoints(path) {
  var pts = __readPath(path), lines = [];
  for (var i = 0; i < pts.length; i++) {
    var p = pts[i];
    lines.push("  [" + i + "] x=" + __r(p.x) + " y=" + __r(p.y) + " left=(" + __r(p.lx) + "," + __r(p.ly) + ") right=(" + __r(p.rx) + "," + __r(p.ry) + ") type=" + p.type);
  }
  return lines.join("\\n");
}

function __pathCount(item) {
  var n = 0;
  try { for (var i = 0; i < item.paths.length; i++) n += item.paths[i].pathPoints.length; } catch (e) {}
  return n;
}

function __isClosed(path) { return String(path.pathType) === "CLOSED_PATH"; }

// "none" selects the [None] swatch (always the first swatch), any other name must exist
function __swatchOrNone(doc, name) {
  if (name === null || String(name).toLowerCase() === "none") return doc.swatches[0];
  return __swatch(doc, name);
}

// Shared style options for newly created shapes: { fill, stroke, strokeWeight, opacity, name, label }
function __styleShape(doc, item, o) {
  if (__has(o.fill)) item.fillColor = __swatchOrNone(doc, o.fill);
  if (__has(o.stroke)) item.strokeColor = __swatchOrNone(doc, o.stroke);
  if (__has(o.strokeWeight)) item.strokeWeight = o.strokeWeight;
  if (__has(o.opacity)) item.transparencySettings.blendingSettings.opacity = o.opacity;
  if (__has(o.name)) item.name = o.name;
  if (__has(o.label)) item.label = o.label;
}

// Stroke style by name; "solid" / "dashed" work in every UI language (the built-in styles keep their position)
function __strokeStyle(doc, name) {
  var n = String(name).toLowerCase(), list = doc.strokeStyles, s;
  if (n === "solid") { s = list.itemByName("Solid"); return s.isValid ? s : list[list.length - 1]; }
  if (n === "dashed") { s = list.itemByName("Dashed"); return s.isValid ? s : list[list.length - 2]; }
  s = list.itemByName(String(name));
  if (!s.isValid) throw new Error("Stroke style not found: " + name);
  return s;
}

function __describePath(item) {
  var s = __describe(item) + " points=" + __pathCount(item);
  try { s += " paths=" + item.paths.length + " closed=" + __isClosed(item.paths[0]); } catch (e) {}
  try { s += " opacity=" + __r(item.transparencySettings.blendingSettings.opacity); } catch (e) {}
  try { if (item.strokeColor && item.strokeColor.isValid) s += " stroke=" + item.strokeColor.name + "/" + __r(item.strokeWeight) + "pt"; } catch (e) {}
  return s;
}

function __overflowNote(frame) {
  var lines = "";
  try { lines = "lines=" + frame.lines.length + " "; } catch (e) {}
  try { return lines + (frame.overflows ? "overflows=true (text does not fit - enlarge the frame or reduce the text)" : "overflows=false"); } catch (e) { return ""; }
}
`;

// Wrapper that runs a tool script and writes its completion value to a result file.
export function buildWrapper(script, resultPath) {
  return `${PRELUDE}
var __result__;
var __doc0 = null, __hu, __vu, __ro, __ui = null;
// An error dialog would block InDesign (and every later tool call) until someone clicks OK:
// scripts must get an exception instead.
try {
  __ui = app.scriptPreferences.userInteractionLevel;
  app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
} catch (e) { __ui = null; }
try {
  if (app.documents.length > 0) {
    __doc0 = app.activeDocument;
    var __vp = __doc0.viewPreferences;
    __hu = __vp.horizontalMeasurementUnits; __vu = __vp.verticalMeasurementUnits; __ro = __vp.rulerOrigin;
    __vp.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    __vp.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    __vp.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  }
} catch (e) { __doc0 = null; }
try {
  __result__ = eval(${jsStr(script)});
} catch (error) {
  __result__ = "ERROR: " + error.message + " (Code: " + (error.number || "n/a") + ", Line: " + (error.line || "unknown") + ")";
}
try { if (__ui !== null) app.scriptPreferences.userInteractionLevel = __ui; } catch (e) {}
try {
  if (__doc0 && __doc0.isValid) {
    var __vp2 = __doc0.viewPreferences;
    __vp2.horizontalMeasurementUnits = __hu; __vp2.verticalMeasurementUnits = __vu; __vp2.rulerOrigin = __ro;
  }
} catch (e) {}
if (typeof __result__ !== "undefined" && __result__ !== null) {
  var __f__ = new File(${jsStr(resultPath)});
  __f__.encoding = "UTF-8";
  __f__.lineFeed = "Unix";
  __f__.open("w");
  __f__.write(String(__result__));
  __f__.close();
}
`;
}
