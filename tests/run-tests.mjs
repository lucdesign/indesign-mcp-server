#!/usr/bin/env node
// End-to-end tests: drives the MCP server over stdio against a real, running InDesign.
//
//   npm test                       (starts InDesign if needed)
//   INDESIGN_APP_NAME="Adobe InDesign 2026" npm test
//
// Every document the tests create is closed again; documents you already have open are
// never touched. Assertions only use our own messages and numeric values, never the
// (localised) wording of InDesign's messages, so the suite passes on any system language.

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { ToolListChangedNotificationSchema } from '@modelcontextprotocol/sdk/types.js';
import { execSync } from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';
import { deltaE } from '../lib/color.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP = process.env.INDESIGN_APP_NAME || 'Adobe InDesign 2026';
// Real-world sample graphics shipped in the repo (override with INDESIGN_TEST_AI / INDESIGN_TEST_EPS / INDESIGN_TEST_JPG)
const SAMPLES = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'sample-files');
for (const [key, name] of [['INDESIGN_TEST_AI', 'sample-1MB.ai'], ['INDESIGN_TEST_EPS', 'sample-1MB.eps'], ['INDESIGN_TEST_JPG', 'sample-1MB.jpg']]) {
  if (!process.env[key] && fs.existsSync(path.join(SAMPLES, name))) process.env[key] = path.join(SAMPLES, name);
}
const TMP = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'indesign-mcp-test-')));

// ---------------------------------------------------------------- harness
const results = [];
let client;
const createdDocs = [];

function assert(cond, message) {
  if (!cond) throw new Error(`Assertion failed: ${message}`);
}
function eq(actual, expected, message) {
  assert(actual === expected, `${message} (expected ${expected}, got ${actual})`);
}
function near(actual, expected, tol, message) {
  assert(Math.abs(actual - expected) <= tol, `${message} (expected ~${expected} +/-${tol}, got ${actual})`);
}
let flyer;                       // the main scratch document (created by the first document test)
let testDoc;                     // the document every test starts with (flyer unless a section switches)
async function test(name, fn) {
  if (process.env.TEST_FILTER && !new RegExp(process.env.TEST_FILTER, 'i').test(name)) return;   // e.g. TEST_FILTER="colour management"
  const started = Date.now();
  try {
    // Every test starts with one of our own documents active, never a document the user had open
    const start = testDoc && createdDocs.includes(testDoc) ? testDoc : flyer;
    if (start && createdDocs.includes(start)) await call('activate_document', { name: start });
    await fn();
    results.push({ name, ok: true });
    console.log(`  ok    ${name} (${Date.now() - started} ms)`);
  } catch (error) {
    results.push({ name, ok: false, error });
    console.log(`  FAIL  ${name}\n        ${String(error.message).split('\n')[0]}`);
  }
}

async function call(name, args = {}) {
  const r = await client.callTool({ name, arguments: args });
  return r.content.map((c) => c.text).join('\n');
}
// Full result content (text and image blocks)
async function callRaw(name, args = {}) {
  const r = await client.callTool({ name, arguments: args });
  return r.content;
}
// Expect the call to fail; returns the error message
async function fails(name, args, pattern) {
  try {
    await client.callTool({ name, arguments: args });
  } catch (error) {
    if (pattern) assert(pattern.test(error.message), `error for ${name} should match ${pattern}, got: ${error.message}`);
    return error.message;
  }
  throw new Error(`Assertion failed: ${name} should have failed`);
}
const idOf = (text) => Number(text.match(/\bid=(\d+)/)?.[1] ?? NaN);
const numAfter = (text, label) => Number(text.match(new RegExp(`${label}[:=]\\s*(-?[\\d.]+)`))?.[1] ?? NaN);

async function newDoc(args) {
  const out = await call('create_document', args);
  const name = out.match(/Document created: (.+?) - /)?.[1];
  assert(name, `create_document should report the document name: ${out}`);
  createdDocs.push(name);
  return name;
}
async function closeDoc(name) {
  await call('close_document', { name, confirmDestructive: true });
  const i = createdDocs.indexOf(name);
  if (i >= 0) createdDocs.splice(i, 1);
}

// Read-only inspection with our own tiny runner (test-only, does not use the server)
function inspect(js) {
  const dir = path.join(TMP, 'inspect');
  fs.mkdirSync(dir, { recursive: true });
  const jsx = path.join(dir, 'i.jsx');
  const res = path.join(dir, 'r.txt');
  fs.rmSync(res, { force: true });
  fs.writeFileSync(jsx, `var __r; try { __r = (function(){ ${js} })(); } catch (e) { __r = "ERR " + e.number + " " + e.message; }
var f = new File(${JSON.stringify(res)}); f.encoding = "UTF-8"; f.lineFeed = "Unix"; f.open("w"); f.write(String(__r)); f.close();`);
  execSync(`osascript -e 'tell application "${APP}" to do script (POSIX file "${jsx}") language javascript'`);
  return fs.readFileSync(res, 'utf8');
}

function ensureInDesign() {
  const running = () => {
    try { execSync(`pgrep -f "${APP}.app/Contents/MacOS"`, { stdio: 'ignore' }); return true; } catch { return false; }
  };
  if (!running()) {
    console.log(`Starting ${APP} ...`);
    execSync(`open -a "${APP}"`);
  }
  const deadline = Date.now() + 180000;
  for (;;) {
    try {
      execSync(`osascript -e 'tell application "${APP}" to get version'`, { stdio: 'ignore', timeout: 20000 });
      return;
    } catch {
      if (Date.now() > deadline) throw new Error(`${APP} did not become ready in 3 minutes`);
      execSync('sleep 3');
    }
  }
}

// Minimal PNG writer (RGB, gradient) for the place_image tests
function writePng(file, w, h) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type), data]);
    const c = Buffer.alloc(4); c.writeUInt32BE(crc(body));
    return Buffer.concat([len, body, c]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const o = y * (w * 3 + 1) + 1 + x * 3;
      raw[o] = (x * 255) / w; raw[o + 1] = (y * 255) / h; raw[o + 2] = 128;
    }
  }
  fs.writeFileSync(file, Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0)),
  ]));
}
const pngSize = (f) => { const b = fs.readFileSync(f); return [b.readUInt32BE(16), b.readUInt32BE(20)]; };
function jpegSize(f) {
  const b = fs.readFileSync(f);
  let i = 2;
  while (i < b.length) {
    if (b[i] !== 0xff) { i++; continue; }
    const m = b[i + 1];
    if (m >= 0xc0 && m <= 0xc3) return [b.readUInt16BE(i + 7), b.readUInt16BE(i + 5)];
    i += 2 + b.readUInt16BE(i + 2);
  }
  throw new Error('no JPEG SOF marker');
}
function pdfBoxes(f) {
  const text = fs.readFileSync(f, 'latin1');
  const box = (n) => (text.match(new RegExp(`/${n}\\[([^\\]]+)\\]`))?.[1] || '').trim().split(/\s+/).map(Number);
  const pages = (text.match(/\/Type\s*\/Page[^s]/g) || []).length;
  return { media: box('MediaBox'), trim: box('TrimBox'), pages };
}
const boxSize = (b) => [b[2] - b[0], b[3] - b[1]];
const mm = (pt) => (pt * 25.4) / 72;

// Minimal PNG decoder (8-bit, non-interlaced, gray/RGB/RGBA/palette) for pixel checks
function readPng(file) {
  const b = Buffer.isBuffer(file) ? file : fs.readFileSync(file);
  let off = 8, w = 0, h = 0, depth = 8, ctype = 2, palette = null;
  const idat = [];
  while (off < b.length) {
    const len = b.readUInt32BE(off), type = b.toString('ascii', off + 4, off + 8), data = b.subarray(off + 8, off + 8 + len);
    if (type === 'IHDR') { w = data.readUInt32BE(0); h = data.readUInt32BE(4); depth = data[8]; ctype = data[9]; }
    else if (type === 'PLTE') palette = data;
    else if (type === 'IDAT') idat.push(data);
    off += 12 + len;
  }
  assert(depth === 8, `PNG depth ${depth} not supported by the test decoder`);
  const ch = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ctype];
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = w * ch, out = Buffer.alloc(h * stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)], row = y * stride, prev = row - stride;
    for (let x = 0; x < stride; x++) {
      const v = raw[y * (stride + 1) + 1 + x];
      const a = x >= ch ? out[row + x - ch] : 0, up = y > 0 ? out[prev + x] : 0, c = x >= ch && y > 0 ? out[prev + x - ch] : 0;
      let r;
      if (f === 0) r = v; else if (f === 1) r = v + a; else if (f === 2) r = v + up; else if (f === 3) r = v + ((a + up) >> 1);
      else { const p = a + up - c, pa = Math.abs(p - a), pb = Math.abs(p - up), pc = Math.abs(p - c); r = v + (pa <= pb && pa <= pc ? a : pb <= pc ? up : c); }
      out[row + x] = r & 255;
    }
  }
  const px = (x, y) => {
    const i = y * stride + x * ch;
    if (ctype === 3) return [palette[out[i] * 3], palette[out[i] * 3 + 1], palette[out[i] * 3 + 2]];
    if (ctype === 0 || ctype === 4) return [out[i], out[i], out[i]];
    return [out[i], out[i + 1], out[i + 2]];
  };
  return { w, h, px };
}
// Bounding box of "ink" (pixels that are not white)
function inkBox(png) {
  let x0 = png.w, y0 = png.h, x1 = -1, y1 = -1;
  for (let y = 0; y < png.h; y++) for (let x = 0; x < png.w; x++) {
    const [r, g, b] = png.px(x, y);
    if (r < 200 || g < 200 || b < 200) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
  }
  return x1 < 0 ? null : { x0, y0, x1, y1, w: x1 - x0 + 1, h: y1 - y0 + 1 };
}

// Object geometry from get_object_info: { x, y, w, h }
const geo = (text) => {
  const m = text.match(/geometry: .*? x=(-?[\d.]+) y=(-?[\d.]+) w=(-?[\d.]+) h=(-?[\d.]+)/);
  assert(m, `no geometry in: ${text.slice(0, 200)}`);
  return { x: Number(m[1]), y: Number(m[2]), w: Number(m[3]), h: Number(m[4]) };
};
// Points from a "[i] x=.. y=.. left=(..,..) right=(..,..) type=.." listing
const readPoints = (text) => [...text.matchAll(/\[(\d+)\] x=(-?[\d.]+) y=(-?[\d.]+) left=\((-?[\d.]+),(-?[\d.]+)\) right=\((-?[\d.]+),(-?[\d.]+)\) type=(\w+)/g)]
  .map((m) => ({ x: +m[2], y: +m[3], lx: +m[4], ly: +m[5], rx: +m[6], ry: +m[7], type: m[8] }));
const cubic = (p0, p1, p2, p3, t) => { const u = 1 - t; return u ** 3 * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t ** 3 * p3; };
const sampleY = (a, b, n = 200) => Array.from({ length: n + 1 }, (_, i) => cubic(a.y, a.ry, b.ly, b.y, i / n));

const isRed = ([r, g, b]) => r > 200 && g < 80 && b < 80;
const isBlue = ([r, g, b]) => b > 100 && r < 90 && g < 90 && b > r * 2; // CMYK 100/100/0/0 converts to about RGB 43,36,131


// ---------------------------------------------------------------- tests
async function main() {
  ensureInDesign();
  const outDir = path.join(TMP, 'export');
  const png = path.join(TMP, 'test-image.png');
  writePng(png, 200, 100);

  const transport = new StdioClientTransport({
    command: 'node',
    args: [path.join(__dirname, '..', 'index.js')],
    // optional real-world graphics: INDESIGN_TEST_AI=/path/file.ai  INDESIGN_TEST_EPS=/path/file.eps (their folders are made readable for the run)
    env: { ...process.env, INDESIGN_ALLOWED_DIRS: [TMP, ...['INDESIGN_TEST_AI', 'INDESIGN_TEST_EPS', 'INDESIGN_TEST_JPG'].filter((k) => process.env[k]).map((k) => path.dirname(process.env[k]))].join(':'), INDESIGN_ALLOW_ARBITRARY_CODE: '' },
  });
  client = new Client({ name: 'indesign-mcp-tests', version: '1' }, { capabilities: {} });
  await client.connect(transport);

  const scriptDir = path.join(os.tmpdir(), 'indesign-mcp');
  const scriptFiles = () => (fs.existsSync(scriptDir) ? fs.readdirSync(scriptDir).filter((f) => /^(script|result|applescript)_/.test(f)) : []);

  console.log('\nServer');
  await test('the server announces its tool list (notifications/tools/list_changed) and declares the capability', async () => {
    assert(client.getServerCapabilities()?.tools?.listChanged === true, 'tools.listChanged capability');
    const got = new Promise((resolve) => client.setNotificationHandler(ToolListChangedNotificationSchema, () => resolve(true)));
    const second = new Client({ name: 'notify-check', version: '1' }, { capabilities: {} });
    const seen = new Promise((resolve) => second.setNotificationHandler(ToolListChangedNotificationSchema, () => resolve(true)));
    await second.connect(new StdioClientTransport({ command: 'node', args: [path.join(__dirname, '..', 'index.js')], env: { ...process.env, INDESIGN_ALLOWED_DIRS: TMP } }));
    const timeout = new Promise((r) => setTimeout(() => r(false), 5000));
    assert(await Promise.race([seen, timeout]), 'a freshly started server sends tools/list_changed after initialisation');
    await second.close();
    const { serverInfo } = { serverInfo: client.getServerVersion() };
    assert(/^\d+\.\d+\.\d+$/.test(serverInfo.version) && serverInfo.version !== '1.0.0', `server version is bumped: ${serverInfo.version}`);
  });

  await test('every tool is listed and the arbitrary-code tool stays locked', async () => {
    const { tools } = await client.listTools();
    const names = tools.map((t) => t.name);
    for (const n of ['list_page_items', 'delete_object', 'rotate_object', 'set_object_geometry', 'send_to_back', 'bring_to_front',
      'send_backward', 'bring_forward', 'set_object_opacity', 'set_object_overprint', 'set_text_frame_options', 'group_objects',
      'ungroup', 'apply_character_style_to_range', 'clear_overrides', 'list_open_documents', 'activate_document',
      'convert_rgb_to_cmyk', 'insert_markdown_text', 'export_pdf', 'export_images']) {
      assert(names.includes(n), `tool ${n} should be listed`);
    }
    await fails('execute_indesign_code', { code: '1' }, /disabled/i);
  });

  console.log('\nDocuments');
  await test('create_document A5 landscape, 2 pages, bleed and slug', async () => {
    const before = scriptFiles().length;
    flyer = await newDoc({ preset: 'A5', orientation: 'Landscape', pages: 2, bleed: 3, slug: 5, marginTop: 10, marginBottom: 10, marginLeft: 10, marginRight: 10 });
    eq(scriptFiles().length, before, 'temp scripts are cleaned up after successful calls');
    const info = await call('get_document_info');
    near(numAfter(info, 'Width'), 210, 0.001, 'width');
    near(numAfter(info, 'Height'), 148, 0.001, 'height');
    eq(numAfter(info, 'Pages'), 2, 'pages');
    assert(/Bleed top\/bottom\/inside\/outside: 3 \/ 3 \/ 3 \/ 3/.test(info), 'bleed reported');
    assert(/Slug top\/bottom\/inside\/outside: 5 \/ 5 \/ 5 \/ 5/.test(info), 'slug reported');
    assert(/Top: 10\b/.test(info) && /Left: 10\b/.test(info), 'margins reported');
    assert(/Facing Pages: false/.test(info), 'facing pages off');
    assert(/Intent: Print/.test(info) && /SWATCHES \(\d+\)/.test(info), 'colour intent and swatches reported');
  });

  await test('create_document A5 facing pages x5 (portrait 148 x 210)', async () => {
    const name = await newDoc({ preset: 'A5', facingPages: true, pages: 5 });
    const info = await call('get_document_info');
    eq(numAfter(info, 'Pages'), 5, 'pages');
    near(numAfter(info, 'Width'), 148, 0.001, 'width');
    near(numAfter(info, 'Height'), 210, 0.001, 'height');
    assert(/Facing Pages: true/.test(info), 'facing pages on');
    await closeDoc(name);
  });

  await test('Custom preset honours an explicit orientation and rounds sizes', async () => {
    let n = await newDoc({ preset: 'Custom', width: 100, height: 210, orientation: 'Landscape' });
    let info = await call('get_document_info');
    near(numAfter(info, 'Width'), 210, 0.001, 'landscape width'); near(numAfter(info, 'Height'), 100, 0.001, 'landscape height');
    await closeDoc(n);
    n = await newDoc({ preset: 'Custom', width: 100, height: 210 });
    info = await call('get_document_info');
    near(numAfter(info, 'Width'), 100, 0.001, 'as-given width'); near(numAfter(info, 'Height'), 210, 0.001, 'as-given height');
    await closeDoc(n);
    n = await newDoc({ preset: 'Custom', width: 210, height: 100, orientation: 'Portrait' });
    info = await call('get_document_info');
    near(numAfter(info, 'Width'), 100, 0.001, 'portrait width');
    await closeDoc(n);
    const out = await call('create_document', { preset: 'A5', orientation: 'Landscape' });
    assert(/\(210 x 148 mm\)/.test(out), `sizes are rounded: ${out}`);
    await closeDoc(out.match(/created: (.+?) - /)[1]);
    await fails('create_document', { preset: 'Custom', width: 100 }, /width and height/);
    await fails('create_document', { preset: 'Nope' }, /Unknown preset|Invalid parameter "preset"/);
  });

  await test('list_open_documents / activate_document / close_document', async () => {
    const second = await newDoc({ preset: 'A4' });
    let list = await call('list_open_documents');
    assert(new RegExp(`${second.replace(/[-]/g, '\\-')}[^\\n]*\\(ACTIVE\\)`).test(list), 'second doc is active');
    await call('activate_document', { name: flyer });
    list = await call('list_open_documents');
    assert(new RegExp(`${flyer.replace(/[-]/g, '\\-')}[^\\n]*\\(ACTIVE\\)`).test(list), 'flyer is active again');
    await fails('activate_document', { name: 'no-such-document.indd' }, /not found/i);
    await fails('close_document', { name: second }, /unsaved changes/i);
    await call('close_document', { name: second, confirmDestructive: true });
    createdDocs.splice(createdDocs.indexOf(second), 1);
    await fails('close_document', { name: second }, /not found/i);
  });

  console.log('\nColour');
  await test('swatches: CMYK default, RGB/hex converted for print, presets', async () => {
    let out = await call('create_color_swatch', { name: 'Brand', colorValues: [100, 50, 0, 10] });
    assert(/CMYK 100,50,0,10/.test(out), out);
    out = await call('create_color_swatch', { name: 'Orange', hex: '#FF6600' });
    assert(/CMYK/.test(out) && /converted to CMYK with the document profile/.test(out), out);
    out = await call('create_color_swatch', { name: 'RichBlack', preset: 'rich_black' });
    assert(/CMYK 60,40,40,100/.test(out), out);
    out = await call('create_color_swatch', { name: 'ScreenBlue', colorModel: 'RGB', colorValues: [0, 80, 255], keepRgb: true });
    assert(/RGB 0,80,255/.test(out), out);
    await fails('create_color_swatch', { name: 'Brand', colorValues: [0, 0, 0, 0] }, /already exists/);
    await fails('create_color_swatch', { name: 'Bad', colorValues: [1, 2] }, /colorValues/);
    const conv = await call('convert_rgb_to_cmyk', { rgb: [255, 0, 0] });
    assert(/CMYK 0,\d+(\.\d+)?,\d+(\.\d+)?,0 \(C,M,Y,K/.test(conv) && /Colour-managed/.test(conv), conv);
    assert(/Orange \[CMYK process: 0,70\.\d+,9\d\.\d+,0\]/.test(await call('get_document_info')), 'Orange stored as a colour-managed CMYK value');
  });


  await test('colour management: hex/RGB become CMYK through the document profile (no blue shift)', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    const intended = [0x2e, 0x7d, 0x32];
    const render = async (swatch, y) => {
      const id = idOf(await call('create_rectangle', { x: 10, y, width: 30, height: 20, fillColor: swatch }));
      const c = await callRaw('render_preview', { objectId: id, resolution: 36 });
      const png = readPng(Buffer.from(c.find((q) => q.type === 'image').data, 'base64'));
      return png.px(png.w >> 1, png.h >> 1);
    };
    // profile conversion (default): the numbers InDesign itself produces, and the colour on screen stays put
    let out = await call('create_color_swatch', { name: 'CmsGreen', hex: '#2e7d32' });
    const v = out.match(/CMYK ([\d.]+),([\d.]+),([\d.]+),([\d.]+)/).slice(1).map(Number);
    [72.7, 10.4, 89.5, 28.9].forEach((e, i) => near(v[i], e, 1.5, `profile conversion CMYK[${i}]`));
    assert(/PSO Coated v3|document profile/.test(out), out);
    const cms = await render('CmsGreen', 10);
    assert(deltaE(intended, cms) < 4, `the profile conversion renders as the intended green: ${cms} vs ${intended} (delta E ${deltaE(intended, cms).toFixed(1)})`);
    // simple formula: what used to happen, kept as an option: visibly shifted toward blue
    out = await call('create_color_swatch', { name: 'SimpleGreen', hex: '#2e7d32', conversion: 'simple' });
    const w = out.match(/CMYK ([\d.]+),([\d.]+),([\d.]+),([\d.]+)/).slice(1).map(Number);
    [63.2, 0, 60, 51].forEach((e, i) => near(w[i], e, 0.6, `simple conversion CMYK[${i}]`));
    assert(/simple formula/.test(out), out);
    const simple = await render('SimpleGreen', 50);
    assert(deltaE(intended, simple) > 12, `the simple formula is clearly off (delta E ${deltaE(intended, simple).toFixed(1)})`);
    assert(simple[2] - intended[2] > 15, `and its blue channel is too high: ${simple} vs ${intended}`);
    assert(deltaE(intended, cms) < deltaE(intended, simple) / 3, 'the profile conversion is at least 3 times closer');
    // update_color_swatch uses the same conversion
    out = await call('update_color_swatch', { name: 'CmsGreen', hex: '#00897b' });
    const u = out.match(/CMYK process ([\d.]+),([\d.]+),([\d.]+),([\d.]+)/).slice(1).map(Number);
    [77.1, 13.1, 49.3, 17.7].forEach((e, i) => near(u[i], e, 1.5, `updated swatch CMYK[${i}]`));
    // convert_rgb_to_cmyk reports the gamut
    let conv = await call('convert_rgb_to_cmyk', { hex: '#2e7d32' });
    assert(/CMYK 72\.7,10\.4,89\.5,28\.9/.test(conv) && /reproduced accurately/.test(conv) && /PSO Coated v3/.test(conv), conv);
    conv = await call('convert_rgb_to_cmyk', { hex: '#1976d2' });
    assert(/OUTSIDE the print gamut/.test(conv) && /looks like #[0-9A-F]{6}/.test(conv), `out-of-gamut colours are flagged: ${conv}`);
    conv = await call('convert_rgb_to_cmyk', { hex: '#2e7d32', conversion: 'simple' });
    assert(/CMYK 63\.2,0,60,51/.test(conv) && /hues drift/.test(conv), conv);
    await fails('convert_rgb_to_cmyk', { rgb: [1, 2] }, /Pass rgb/);
    await fails('create_color_swatch', { name: 'X', hex: '#123456', conversion: 'magic' }, /Invalid parameter "conversion"/);
    await closeDoc(name);
    testDoc = flyer;
    // without a document only the profile-free formula works
    if (/OPEN DOCUMENTS \(0\)/.test(await call('list_open_documents'))) await fails('convert_rgb_to_cmyk', { hex: '#2e7d32' }, /No document open/);
  });

  console.log('\nStyles');
  await test('paragraph and character styles', async () => {
    let out = await call('create_paragraph_style', {
      name: 'Title', fontFamily: 'Helvetica Neue', fontStyle: 'Bold', fontSize: 24, leading: 28, tracking: 10, kerning: 'Optical',
      spaceBeforePt: 6, spaceAfter: 2, firstLineIndent: 3, leftIndent: 2, hyphenation: false, keepWithNext: 1, keepLinesTogether: true, alignment: 'LEFT_ALIGN',
    });
    assert(/font=Helvetica Neue Bold/.test(out), `applied font is reported: ${out}`);
    assert(/size=24pt/.test(out) && /leading=28pt/.test(out) && /tracking=10/.test(out) && /firstLineIndent=3mm/.test(out) && /hyphenation=false/.test(out), out);
    near(numAfter(out, 'spaceBefore'), 2.117, 0.01, '6 pt in mm');
    out = await call('create_paragraph_style', { name: 'Body', baseStyle: 'Title', fontFamily: 'Helvetica Neue', fontStyle: 'Regular', fontSize: 10, leading: 13 });
    assert(/font=Helvetica Neue Regular/.test(out), out);
    out = await call('modify_paragraph_style', { styleName: 'Body', fontStyle: 'Medium', spaceAfterPt: 4 });
    assert(/fontStyle=Medium/.test(out) && /font=Helvetica Neue Medium/.test(out), out);
    await call('modify_paragraph_style', { styleName: 'Body', fontStyle: 'Regular', fontSize: 10 });
    await call('create_paragraph_style', { name: 'Heading 1', basedOn: undefined, fontFamily: 'Helvetica Neue', fontStyle: 'Bold', fontSize: 20 });
    await call('create_character_style', { name: 'Bold', fontFamily: 'Helvetica Neue', fontStyle: 'Bold' });
    await call('create_character_style', { name: 'Italic', fontFamily: 'Helvetica Neue', fontStyle: 'Italic' });
    await call('create_character_style', { name: 'Bold Italic', fontFamily: 'Helvetica Neue', fontStyle: 'Bold Italic' });
    await fails('create_paragraph_style', { name: 'Title' }, /already exists/);
    await fails('create_paragraph_style', { name: 'X', fontFamily: 'NoSuchFontXYZ' }, /Font not installed/);
    await fails('create_paragraph_style', { name: 'X', baseStyle: 'NoSuchStyle' }, /not found/);
    assert(!/\bX\b/.test(await call('list_styles', { styleType: 'paragraph' })), 'a failed create leaves no half-made style');
    await fails('modify_paragraph_style', { styleName: 'Nope', fontSize: 9 }, /not found/);
  });

  console.log('\nText frames');
  let headline;
  await test('create_text_frame: id, style wins over defaults, no local overrides', async () => {
    const out = await call('create_text_frame', { content: 'Headline', paragraphStyle: 'Title', x: 10, y: 10, width: 120, height: 20 });
    headline = idOf(out);
    assert(headline > 0, `frame id returned: ${out}`);
    assert(/overflows=false/.test(out), out);
    const list = await call('list_text_frames');
    assert(new RegExp(`id=${headline}\\b[^\\n]*style="Title" size=24pt font="Helvetica Neue Bold"`).test(list), `style values must not be overridden by defaults: ${list}`);
    // explicit parameters still apply
    const o2 = await call('create_text_frame', { content: 'Explicit', paragraphStyle: 'Body', fontSize: 33, x: 10, y: 40, width: 120, height: 20 });
    assert(new RegExp(`id=${idOf(o2)}\\b[^\\n]*size=33pt`).test(await call('list_text_frames')), 'explicit fontSize is applied');
    await fails('create_text_frame', { content: 'x', paragraphStyle: 'Nope' }, /not found/);
    await fails('create_text_frame', { content: 'x', fontFamily: 'NoSuchFontXYZ' }, /Font not installed/);
    await fails('create_text_frame', { content: 'x', pageIndex: 9 }, /Invalid page index/);
    const frames = numAfter(await call('get_document_info'), 'Text Frames');
    eq(frames, 2, 'failed creates leave no stray frames');
  });

  await test('line breaks: paragraph vs forced (real Shift+Enter, not a space)', async () => {
    const p = await call('create_text_frame', { content: 'A\nB', paragraphStyle: 'Body', y: 70, width: 100, height: 20 });
    eq(numAfter(p, 'paragraphs'), 2, '\\n is a paragraph break by default');
    eq(numAfter(p, 'lines'), 2, 'two paragraphs are two lines');
    const f = await call('create_text_frame', { content: 'A\nB', lineBreak: 'forced', paragraphStyle: 'Body', y: 70, width: 100, height: 20 });
    eq(numAfter(f, 'paragraphs'), 1, 'forced mode keeps one paragraph');
    eq(numAfter(f, 'lines'), 2, 'forced mode really breaks the line in a wide frame');
    const br = await call('create_text_frame', { content: 'A<br>B\nC', paragraphStyle: 'Body', y: 70, width: 100, height: 20 });
    eq(numAfter(br, 'paragraphs'), 2, '<br> is a forced break, \\n a paragraph break');
    eq(numAfter(br, 'lines'), 3, 'three lines: A / B / C');
    const shown = await call('get_text_content', { frameId: idOf(br), normalizeSpaces: false });
    assert(/A↵B/.test(shown), `a forced break is shown as an arrow, not a space: ${shown}`);
    assert(/A↵B/.test(await call('get_text_content', { frameId: idOf(br) })), 'also with normalizeSpaces');
    const spaced = await call('create_text_frame', { content: 'A B', paragraphStyle: 'Body', y: 70, width: 100, height: 20 });
    eq(numAfter(spaced, 'lines'), 1, 'a plain space does not break');
    const edited = await call('edit_text_frame', { frameId: idOf(spaced), content: 'X<br>Y', lineBreak: 'paragraph' });
    eq(numAfter(edited, 'lines'), 2, 'edit_text_frame also inserts real forced breaks');
    eq(numAfter(edited, 'paragraphs'), 1, 'one paragraph after <br>');
    const edited2 = await call('edit_text_frame', { frameId: idOf(spaced), content: 'X\nY', lineBreak: 'forced' });
    eq(numAfter(edited2, 'paragraphs'), 1, 'lineBreak forced in edit_text_frame');
    for (const o of [p, f, br, spaced]) await call('delete_object', { id: idOf(o) });
  });

  await test('forced line breaks are visible in an exported image', async () => {
    const name = await newDoc({ preset: 'A5', pages: 2 });
    await call('create_paragraph_style', { name: 'Big', fontFamily: 'Helvetica Neue', fontStyle: 'Bold', fontSize: 40, leading: 48, hyphenation: false });
    await call('create_text_frame', { content: 'AAA<br>BBB', paragraphStyle: 'Big', x: 10, y: 20, width: 120, height: 60 });
    await call('create_text_frame', { content: 'AAA BBB', paragraphStyle: 'Big', pageIndex: 1, x: 10, y: 20, width: 120, height: 60 });
    const dir = path.join(outDir, 'linebreak');
    await call('export_images', { folderPath: dir, resolution: 72, confirmDestructive: true });
    const base = name.replace(/\.indd$/i, '');
    const withBreak = inkBox(readPng(path.join(dir, `${base}_page1.png`)));
    const withSpace = inkBox(readPng(path.join(dir, `${base}_page2.png`)));
    assert(withBreak && withSpace, 'both pages have ink');
    assert(withBreak.h > withSpace.h * 1.6, `the text with <br> is about two lines high (${withBreak.h}px) vs one line (${withSpace.h}px)`);
    assert(withBreak.w < withSpace.w * 0.75, `the broken lines are narrower than the single line (${withBreak.w}px vs ${withSpace.w}px)`);
    await closeDoc(name);
  });

  await test('overflow is reported by create/edit/get_text_content/list', async () => {
    const out = await call('create_text_frame', { content: 'This text is far too long for a tiny frame. '.repeat(5), paragraphStyle: 'Body', x: 10, y: 100, width: 20, height: 6 });
    assert(/overflows=true/.test(out), out);
    const id = idOf(out);
    assert(/overflows=true/.test(await call('get_text_content', { frameId: id })), 'get_text_content reports overflow');
    assert(new RegExp(`id=${id}\\b[^\\n]*overflows=true`).test(await call('list_text_frames')), 'list_text_frames reports overflow');
    const fixed = await call('edit_text_frame', { frameId: id, content: 'Short' });
    assert(/overflows=false/.test(fixed), fixed);
    await call('delete_object', { id });
  });

  await test('edit_text_frame replaces the whole story', async () => {
    const out = await call('create_text_frame', { content: 'Alpha\n Beta', paragraphStyle: 'Body', y: 100, width: 60, height: 20 });
    const id = idOf(out);
    eq(numAfter(out, 'paragraphs'), 2, 'starts with two paragraphs');
    const edited = await call('edit_text_frame', { frameId: id, content: ' ' });
    eq(numAfter(edited, 'paragraphs'), 1, 'one paragraph after replacing');
    const text = await call('get_text_content', { frameId: id, normalizeSpaces: false });
    assert(/Original length: 1 characters/.test(text), text);
    const t = await call('get_text_content', { frameId: id });
    assert(/Paragraphs: 1/.test(t), t);
    await call('delete_object', { id });
  });

  await test('frame ids are stable while the list order changes', async () => {
    const a = idOf(await call('create_text_frame', { content: 'first', paragraphStyle: 'Body', y: 100, height: 8 }));
    const b = idOf(await call('create_text_frame', { content: 'second', paragraphStyle: 'Body', y: 110, height: 8 }));
    let list = await call('list_text_frames');
    assert(/frontmost/.test(list), 'ordering is documented in the output');
    await call('send_to_back', { id: b });
    const edit = await call('edit_text_frame', { frameId: a, content: 'first!' });
    assert(new RegExp(`id=${a}\\b`).test(edit), 'edit by id still hits the right frame after reordering');
    assert(/first!/.test(await call('get_text_content', { frameId: a })), 'content changed on the right frame');
    assert(/second/.test(await call('get_text_content', { frameId: b })), 'other frame untouched');
    await fails('edit_text_frame', { frameId: 999999, content: 'x' }, /No object with id/);
    await fails('edit_text_frame', { frameIndex: 99, content: 'x' }, /Invalid index/);
    await call('delete_object', { id: a }); await call('delete_object', { id: b });
  });

  console.log('\nMarkdown');
  await test('insert_markdown_text: umlauts, quotes, bold/italic/headers, errors', async () => {
    const frame = idOf(await call('create_text_frame', { content: 'x', paragraphStyle: 'Body', x: 10, y: 60, width: 120, height: 60 }));
    const md = '# Größe äöü ß\nUm **7.30 Uhr** beginnt *es* mit "Quotes" und \\\\ Backslash\n\n***beides*** ok, 2 * 3 und **offen';
    const out = await call('insert_markdown_text', { frameId: frame, markdownText: md, bodyStyle: 'Body' });
    assert(/3 paragraphs, 3 character runs/.test(out), out);
    const text = await call('get_text_content', { frameId: frame, normalizeSpaces: false });
    assert(text.includes('Größe äöü ß') && text.includes('Um 7.30 Uhr beginnt es'), `umlauts and text survive: ${text}`);
    assert(text.includes('2 * 3') && text.includes('**offen'), 'unmatched markers stay literal');
    const runs = inspect(`
      var f = __f(); function __f(){ return app.activeDocument.pageItems.itemByID(${frame}); }
      var s = f.parentStory, o = [], last = null, seg = "";
      for (var k = 0; k < s.characters.length; k++) {
        var st = s.characters[k].appliedCharacterStyle.name;
        if (st !== last) { if (seg) o.push(last + "=" + seg); seg = ""; last = st; }
        seg += (typeof s.characters[k].contents === "string" && s.characters[k].contents !== "\\r") ? s.characters[k].contents : "|";
      }
      o.push(last + "=" + seg);
      o.push("P0=" + s.paragraphs[0].appliedParagraphStyle.name + " P1=" + s.paragraphs[1].appliedParagraphStyle.name);
      return o.join("\\n");`);
    assert(/Bold=7\.30 Uhr/.test(runs) && /Italic=es/.test(runs) && /Bold Italic=beides/.test(runs), `character styles applied: ${runs}`);
    assert(/P0=Heading 1 P1=Body/.test(runs), `paragraph styles applied: ${runs}`);
    // append
    const appended = await call('insert_markdown_text', { frameId: frame, markdownText: 'Angehängt', replaceContent: false });
    assert(/1 paragraphs/.test(appended), appended);
    assert((await call('get_text_content', { frameId: frame, normalizeSpaces: false })).includes('Angehängt'), 'appended');
    // missing styles: clear error, nothing changed
    const before = await call('get_text_content', { frameId: frame, normalizeSpaces: false });
    await fails('insert_markdown_text', { frameId: frame, markdownText: '# Hi', styleMap: { h1: 'No Such Style' } }, /Missing styles.*No Such Style/);
    await fails('insert_markdown_text', { frameId: frame, markdownText: '**x**', styleMap: { bold: 'No Such Style' } }, /Missing styles/);
    eq(await call('get_text_content', { frameId: frame, normalizeSpaces: false }), before, 'a failed insert changes nothing');
    await fails('insert_markdown_text', { frameId: frame, markdownText: '   \n\n' }, /no text/);
    await call('delete_object', { id: frame });
  });

  console.log('\nCharacter styles, overrides');
  await test('apply_character_style_to_range and clear_overrides', async () => {
    const id = idOf(await call('create_text_frame', { content: 'Um 7.30 Uhr beginnt es', paragraphStyle: 'Body', x: 10, y: 60, width: 100, height: 20 }));
    let out = await call('apply_character_style_to_range', { frameId: id, styleName: 'Bold', matchText: '7.30 Uhr' });
    assert(/1 range/.test(out), out);
    out = await call('apply_character_style_to_range', { frameId: id, styleName: 'Italic', startIndex: 0, endIndex: 1 });
    assert(/1 range/.test(out), out);
    const applied = inspect(`var s = app.activeDocument.pageItems.itemByID(${id}).parentStory; return s.characters[3].appliedCharacterStyle.name + "," + s.characters[0].appliedCharacterStyle.name + "," + s.characters[15].appliedCharacterStyle.name;`);
    eq(applied, 'Bold,Italic,[None]', 'styles landed on the right characters');
    await fails('apply_character_style_to_range', { frameId: id, styleName: 'Bold', matchText: 'not there' }, /not found/);
    await fails('apply_character_style_to_range', { frameId: id, styleName: 'Bold', startIndex: 5, endIndex: 500 }, /Invalid range/);
    await fails('apply_character_style_to_range', { frameId: id, styleName: 'NoSuch', matchText: 'Um' }, /not found/);
    await call('edit_text_frame', { frameId: id, fontSize: 30 });
    assert(new RegExp(`id=${id}\\b[^\\n]*size=30pt`).test(await call('list_text_frames')), 'local override applied');
    await call('clear_overrides', { frameId: id });
    assert(new RegExp(`id=${id}\\b[^\\n]*size=10pt`).test(await call('list_text_frames')), 'override removed, style size back');
    await call('delete_object', { id });
  });

  console.log('\nObjects');
  let rect;
  await test('rectangle: id, rotate, geometry, opacity, overprint', async () => {
    const out = await call('create_rectangle', { x: 20, y: 30, width: 40, height: 20, fillColor: 'Brand' });
    rect = idOf(out);
    assert(rect > 0 && /fill=Brand/.test(out), out);
    await fails('create_rectangle', { x: 0, y: 0, width: 5, height: 5, fillColor: 'NoSuchSwatch' }, /Swatch not found/);
    let r = await call('rotate_object', { id: rect, angle: 30 });
    near(numAfter(r, 'rotation'), 30, 0.01, 'rotated 30 (counter-clockwise positive)');
    r = await call('rotate_object', { id: rect, angle: 15 });
    near(numAfter(r, 'rotation'), 45, 0.01, 'relative rotation adds');
    r = await call('rotate_object', { id: rect, angle: 10, absolute: true, reference: 'top-left' });
    near(numAfter(r, 'rotation'), 10, 0.01, 'absolute rotation sets');
    r = await call('set_object_geometry', { id: rect, x: 25, y: 35, width: 50, height: 10, rotation: 0 });
    assert(/x=25 y=35 w=50 h=10 rotation=0/.test(r), r);
    r = await call('set_object_geometry', { id: rect, x: 30 });
    assert(/x=30 y=35 w=50 h=10/.test(r), `partial update keeps the rest: ${r}`);
    r = await call('set_object_geometry', { id: rect, rotation: 20 });
    near(numAfter(r, 'rotation'), 20, 0.01, 'geometry can set rotation');
    r = await call('set_object_geometry', { id: rect, x: 12, rotation: 0 });
    assert(/x=12 y=35 w=50 h=10 rotation=0/.test(r), r);
    r = await call('set_object_geometry', { id: rect, shear: 15 });
    assert(/shear=15/.test(r), r);
    await call('set_object_geometry', { id: rect, shear: 0 });
    await call('set_object_overprint', { id: rect, fill: true });
    eq(inspect(`return String(app.activeDocument.pageItems.itemByID(${rect}).overprintFill);`), 'true', 'overprint fill');
    await call('set_object_opacity', { id: rect, opacity: 40, blendMode: 'MULTIPLY' });
    eq(inspect(`var b = app.activeDocument.pageItems.itemByID(${rect}).transparencySettings.blendingSettings; return Math.round(b.opacity) + "," + (b.blendMode === BlendMode.MULTIPLY);`), '40,true', 'opacity and blend mode set');
    await call('set_object_opacity', { id: rect, opacity: 70, target: 'fill' });
    await fails('set_object_opacity', { id: rect, opacity: 150 }, /0-100/);
    await fails('set_object_opacity', { id: rect }, /opacity and\/or blendMode/);
    await fails('set_object_overprint', { id: rect, stroke: true }, /Overprint is not available/);
    await fails('rotate_object', { id: 424242, angle: 5 }, /No object with id/);
    await fails('rotate_object', { id: rect, angle: 5, reference: 'nowhere' }, /Invalid reference|Invalid parameter "reference"/);
    await fails('set_object_geometry', { id: rect }, /Nothing to change/);
  });

  await test('z-order, groups, list_page_items, delete', async () => {
    const back = rect;
    const front = idOf(await call('create_ellipse', { x: 30, y: 30, width: 20, height: 20, fillColor: 'Orange' }));
    const order = () => inspect(`var it = app.activeDocument.pages[0].allPageItems, o = []; for (var i = 0; i < it.length; i++) o.push(it[i].id); return o.join(",");`).split(',').map(Number);
    assert(order().indexOf(front) < order().indexOf(back), 'ellipse (newest) is in front of the rectangle (lower index = in front)');
    await call('send_to_back', { id: front });
    assert(order().indexOf(front) > order().indexOf(back), 'send_to_back moved it behind');
    await call('bring_to_front', { id: front });
    assert(order().indexOf(front) < order().indexOf(back), 'bring_to_front');
    await call('send_backward', { id: front });
    await call('bring_forward', { id: front });
    const list = await call('list_page_items', { pageIndex: 0 });
    assert(new RegExp(`Rectangle id=${back}\\b`).test(list) && new RegExp(`Oval id=${front}\\b`).test(list), list);
    assert(/TextFrame id=\d+/.test(list) && /fill=Brand/.test(list), 'text frames and fills are listed');
    const onlyOvals = await call('list_page_items', { pageIndex: 0, type: 'Oval' });
    assert(!/Rectangle id=/.test(onlyOvals) && /Oval id=/.test(onlyOvals), 'type filter');

    const g = await call('group_objects', { ids: [back, front] });
    const gid = idOf(g);
    assert(/Group id=/.test(g), g);
    assert(new RegExp(`Rectangle id=${back}\\b[^\\n]*group=${gid}`).test(await call('list_page_items', {})), 'children reference their group');
    const u = await call('ungroup', { id: gid });
    assert(u.includes(String(back)) && u.includes(String(front)), `ungroup returns the released ids: ${u}`);
    await fails('ungroup', { id: back }, /not a group/);
    await fails('group_objects', { ids: [back] }, /at least 2/);

    const del = await call('delete_object', { id: front });
    assert(/Deleted: Oval/.test(del), del);
    assert(!new RegExp(`id=${front}\\b`).test(await call('list_page_items', {})), 'deleted object is gone');
    await fails('delete_object', { id: front }, /No object with id/);
  });

  await test('set_text_frame_options and page-relative coordinates on page 2', async () => {
    const id = idOf(await call('create_text_frame', { content: 'Options', paragraphStyle: 'Body', pageIndex: 1, x: 5, y: 5, width: 60, height: 30 }));
    const list = await call('list_page_items', { pageIndex: 1 });
    assert(new RegExp(`TextFrame id=${id}\\b[^\\n]*page=2 x=5 y=5 w=60 h=30`).test(list), `coordinates are relative to the page: ${list}`);
    await call('set_text_frame_options', { id, inset: 3, insetTop: 1, verticalAlignment: 'CENTER', columns: 2, columnGutter: 4, autosize: 'OFF', textWrap: 'BOUNDING_BOX', textWrapOffset: 2 });
    const v = inspect(`var t = app.activeDocument.pageItems.itemByID(${id}), p = t.textFramePreferences;
      var ins = p.insetSpacing; return [Math.round(ins[0]*100)/100, Math.round(ins[1]*100)/100, p.verticalJustification === VerticalJustification.CENTER_ALIGN, p.textColumnCount, Math.round(p.textColumnGutter*100)/100, t.textWrapPreferences.textWrapMode === TextWrapModes.BOUNDING_BOX_TEXT_WRAP].join(",");`);
    eq(v, '1,3,true,2,4,true', 'frame options applied (top inset 1, others 3)');
    await call('set_text_frame_options', { id, autosize: 'HEIGHT_ONLY' });
    await fails('set_text_frame_options', { id }, /No options/);
    await fails('set_text_frame_options', { id, autosize: 'BOGUS' }, /Invalid autosize|Invalid parameter "autosize"/);
    await call('delete_object', { id });
  });

  console.log('\nImages');
  await test('place_image: id, sizing from aspect ratio, crop, offsets', async () => {
    let out = await call('place_image', { imagePath: png, x: 10, y: 10, width: 60, height: 60, fitOption: 'FILL_PROPORTIONALLY' });
    const a = idOf(out);
    assert(a > 0 && /w=60 h=60/.test(out) && /image=test-image\.png/.test(out), out);
    out = await call('place_image', { imagePath: png, x: 80, y: 10, width: 60 });
    assert(/w=60 h=30/.test(out), `height follows the 2:1 aspect ratio: ${out}`);
    const b = idOf(out);
    out = await call('place_image', { imagePath: png, x: 80, y: 50, height: 20 });
    assert(/w=40 h=20/.test(out), out);
    const c = idOf(out);
    out = await call('place_image', { imagePath: png, x: 10, y: 80, width: 40, height: 40, fitOption: 'FILL_PROPORTIONALLY', contentOffsetX: -5, contentScale: 200 });
    const d = idOf(out);
    assert(d > 0, out);
    out = await call('place_image', { imagePath: png, x: 120, y: 80 });
    assert(idOf(out) > 0, out);
    {   // natural size: the picture must fill the moved frame (gradient: dark top-left, bright bottom-right)
      const c = await callRaw('render_preview', { objectId: idOf(out), resolution: 72 });
      const shot = readPng(Buffer.from(c.find((q) => q.type === 'image').data, 'base64'));
      const tl = shot.px(2, 2), br = shot.px(shot.w - 3, shot.h - 3);
      assert(tl[0] < 60 && tl[1] < 60 && br[0] > 190 && br[1] > 190, `natural-size image fills its frame: top-left ${tl}, bottom-right ${br}`);
    }
    await call('delete_object', { id: idOf(out) });
    await fails('place_image', { imagePath: path.join(TMP, 'missing.png') }, /not found/);
    await fails('place_image', { imagePath: '/etc/hosts' }, /Access denied/);
    await fails('place_image', { imagePath: png, fitOption: 'BOGUS' }, /Invalid fitOption|Invalid parameter "fitOption"/);
    assert(numAfter(await call('get_document_info'), 'Rectangles \\(incl\\. image frames\\)') >= 5, 'image frames exist');
    for (const id of [a, b, c, d]) await call('delete_object', { id });
  });

  console.log('\nExport');
  await test('export_pdf: presets, page range, bleed, slug', async () => {
    // content so the pages are not empty
    await call('create_text_frame', { content: 'PDF test', paragraphStyle: 'Title', x: 10, y: 10, width: 100, height: 20 });
    await call('create_text_frame', { content: 'Page two', paragraphStyle: 'Title', pageIndex: 1, x: 10, y: 10, width: 100, height: 20 });
    await fails('export_pdf', { filePath: path.join(outDir, 'x.pdf') }, /confirm/i);
    for (const preset of ['Print', 'Web', 'HighQualityPrint', 'PressQuality']) {
      const file = path.join(outDir, `${preset}.pdf`);
      const out = await call('export_pdf', { filePath: file, preset, confirmDestructive: true });
      assert(/PDF exported/.test(out) && fs.existsSync(file), `${preset}: ${out}`);
      const { media, trim, pages } = pdfBoxes(file);
      eq(pages, 2, `${preset}: page count`);
      near(mm(boxSize(media)[0]), 210, 0.5, `${preset}: media width = trim (no bleed by default)`);
      near(mm(boxSize(trim)[1]), 148, 0.5, `${preset}: trim height`);
    }
    const one = path.join(outDir, 'page2.pdf');
    await call('export_pdf', { filePath: one, pageRange: '2', confirmDestructive: true });
    eq(pdfBoxes(one).pages, 1, 'pageRange "2" exports one page');
    const bleed = path.join(outDir, 'bleed.pdf');
    await call('export_pdf', { filePath: bleed, includeBleed: true, confirmDestructive: true });
    const bb = pdfBoxes(bleed);
    near(mm(boxSize(bb.media)[0]), 216, 0.5, 'includeBleed: media = trim + 2 x 3 mm');
    near(mm(boxSize(bb.trim)[0]), 210, 0.5, 'includeBleed: trim box unchanged');
    const slug = path.join(outDir, 'slug.pdf');
    await call('export_pdf', { filePath: slug, includeBleed: true, includeSlug: true, confirmDestructive: true });
    const sb = pdfBoxes(slug);
    assert(mm(boxSize(sb.media)[0]) > mm(boxSize(bb.media)[0]) + 2, 'includeSlug: the media box grows to the slug area');
    near(mm(boxSize(sb.trim)[0]), 210, 0.5, 'includeSlug: trim box unchanged');
    await fails('export_pdf', { filePath: path.join(outDir, 'n.pdf'), preset: 'No Such Preset', confirmDestructive: true }, /PDF preset not found.*Available/);
    await fails('export_pdf', { filePath: path.join(outDir, 'n.pdf'), pageRange: 'abc', confirmDestructive: true }, /Invalid pageRange/);
    await fails('export_pdf', { filePath: '/etc/x.pdf', confirmDestructive: true }, /Access denied/);
    eq(inspect('var n = []; for (var i = 0; i < app.pdfExportPresets.length; i++) n.push(app.pdfExportPresets[i].name); return n.join("|").indexOf("__indesign_mcp_tmp__");'), '-1', 'temporary preset is removed');
  });

  await test('export_images: PNG/JPEG at 72, 100, 300 dpi, page range, bleed, folder creation', async () => {
    const base = flyer.replace(/\.indd$/i, '');
    const expected = (dpi, mmW) => (mmW / 25.4) * dpi;
    for (const format of ['PNG', 'JPEG']) {
      for (const dpi of [72, 100, 300]) {
        const dir = path.join(outDir, 'new', 'nested', `${format}-${dpi}`); // does not exist yet
        const out = await call('export_images', { folderPath: dir, format, resolution: dpi, confirmDestructive: true });
        assert(/Exported 2 page\(s\)/.test(out), out);
        const ext = format === 'PNG' ? 'png' : 'jpg';
        const f1 = path.join(dir, `${base}_page1.${ext}`);
        assert(fs.existsSync(f1) && fs.existsSync(path.join(dir, `${base}_page2.${ext}`)), `${format} ${dpi}: both pages written`);
        eq(fs.readdirSync(dir).sort().join(','), `${base}_page1.${ext},${base}_page2.${ext}`, `${format} ${dpi}: exactly one file per page, no duplicates`);
        const [w, h] = format === 'PNG' ? pngSize(f1) : jpegSize(f1);
        near(w, expected(dpi, 210), 2, `${format} ${dpi} dpi width`);
        near(h, expected(dpi, 148), 2, `${format} ${dpi} dpi height`);
      }
    }
    const dir = path.join(outDir, 'range');
    const out = await call('export_images', { folderPath: dir, pageRange: '2', includeBleed: true, resolution: 100, confirmDestructive: true });
    assert(/Exported 1 page/.test(out), out);
    eq(fs.readdirSync(dir).join(','), `${base}_page2.png`, 'a page range writes only that page');
    const [w] = pngSize(path.join(dir, `${base}_page2.png`));
    near(w, expected(100, 216), 2, 'includeBleed adds the bleed');
    // facing pages, several pages: every page exactly once, in both formats
    const facing = await newDoc({ preset: 'A5', facingPages: true, pages: 4 });
    for (const format of ['PNG', 'JPEG']) {
      const fdir = path.join(outDir, `facing-${format}`);
      await call('export_images', { folderPath: fdir, format, resolution: 72, confirmDestructive: true });
      const fbase = facing.replace(/\.indd$/i, '');
      const ext = format === 'PNG' ? 'png' : 'jpg';
      eq(fs.readdirSync(fdir).sort().join(','), [1, 2, 3, 4].map((n) => `${fbase}_page${n}.${ext}`).join(','), `${format}: 4 pages give 4 files`);
    }
    await closeDoc(facing);
    await call('activate_document', { name: flyer });
    await fails('export_images', { folderPath: dir, format: 'GIF', confirmDestructive: true }, /Invalid format|Invalid parameter "format"/);
    await fails('export_images', { folderPath: dir, resolution: 0, confirmDestructive: true }, /Invalid resolution/);
    await fails('export_images', { folderPath: dir, pageRange: '9', confirmDestructive: true }, /does not exist/);
  });


  await test('export_images: every page file contains that page (two pages with different content)', async () => {
    const name = await newDoc({ preset: 'A5', pages: 2 });
    await call('create_color_swatch', { name: 'PureRed', colorValues: [0, 100, 100, 0] });
    await call('create_color_swatch', { name: 'PureBlue', colorValues: [100, 100, 0, 0] });
    await call('create_rectangle', { x: 0, y: 0, width: 148, height: 210, fillColor: 'PureRed', pageIndex: 0 });
    await call('create_rectangle', { x: 0, y: 0, width: 148, height: 210, fillColor: 'PureBlue', pageIndex: 1 });
    const base = name.replace(/\.indd$/i, '');
    const center = (file) => { const p = readPng(file); return p.px(p.w >> 1, p.h >> 1); };
    for (const format of ['PNG', 'JPEG']) {
      const ext = format === 'PNG' ? 'png' : 'jpg';
      for (const range of ['all', '1', '2']) {
        const dir = path.join(outDir, `content-${format}-${range}`);
        const out = await call('export_images', { folderPath: dir, format, resolution: 36, pageRange: range, confirmDestructive: true });
        const expectedPages = range === 'all' ? [1, 2] : [Number(range)];
        assert(out.includes(`(pages: ${expectedPages.join(', ')})`), `the result lists the exported pages: ${out}`);
        eq(fs.readdirSync(dir).sort().join(','), expectedPages.map((n) => `${base}_page${n}.${ext}`).join(','), `${format} ${range}: files`);
        if (format === 'PNG') {
          for (const n of expectedPages) {
            const px = center(path.join(dir, `${base}_page${n}.png`));
            assert(n === 1 ? isRed(px) : isBlue(px), `${format} ${range}: ${base}_page${n} shows page ${n} (centre pixel ${px}, expected ${n === 1 ? 'red' : 'blue'})`);
          }
        }
      }
    }
    await closeDoc(name);
  });

  await test('save_document (Save As) and exports never close or replace the document', async () => {
    const name = await newDoc({ preset: 'A5' });
    await call('create_text_frame', { content: 'keep me open' });
    const target = path.join(TMP, 'saved-as.indd');
    const out = await call('save_document', { filePath: target, confirmDestructive: true });
    assert(/now named saved-as\.indd/.test(out) && /document\(s\) open/.test(out), `save reports the rename: ${out}`);
    createdDocs[createdDocs.indexOf(name)] = 'saved-as.indd';
    await call('export_pdf', { filePath: path.join(outDir, 'after-save.pdf'), confirmDestructive: true });
    await call('export_images', { folderPath: path.join(outDir, 'after-save'), confirmDestructive: true });
    await call('export_epub', { filePath: path.join(outDir, 'after-save.epub'), confirmDestructive: true });
    const list = await call('list_open_documents');
    assert(/saved-as\.indd \(ACTIVE\)/.test(list), `the saved document is still open and active: ${list}`);
    assert(!list.includes(name), 'and it is the same document under its new name (no extra Untitled document)');
    await call('save_document', {}); // plain Save on a saved document
    await call('save_document', { filePath: target, confirmDestructive: true }); // Save As onto itself
    assert(/saved-as\.indd \(ACTIVE\)/.test(await call('list_open_documents')), 'still open after saving onto its own path');
    await closeDoc('saved-as.indd');
  });


  console.log('\nOther tools');
  await test('pages, layers, object styles, find/replace, preflight and the remaining tools still run', async () => {
    let out = await call('add_page');
    assert(/Total pages: 3/.test(out), out);
    await call('duplicate_page', { pageIndex: 0 });
    await call('navigate_to_page', { pageIndex: 1 });
    await fails('delete_page', { pageIndex: 3 }, /confirm/i);
    out = await call('delete_page', { pageIndex: 3, confirmDestructive: true });
    assert(/deleted/i.test(out), out);
    await call('create_character_style', { name: 'Sm"art', fontSize: 9 });
    await call('modify_character_style', { styleName: 'Sm"art', tracking: 5 });
    await call('create_object_style', { name: 'OS1', fillColor: 'Brand' });
    await call('modify_object_style', { styleName: 'OS1', strokeColor: 'Brand', strokeWidth: 2 });
    await call('apply_object_style', { styleName: 'OS1', objectIndex: 0, pageIndex: 0 });
    await call('apply_color', { objectIndex: 0, swatchName: 'Brand', property: 'stroke' });
    await call('create_layer', { name: 'L2' });
    await call('set_active_layer', { layerName: 'L2' });
    assert(/L2[^\n]*\[ACTIVE\]/.test(await call('list_layers')), 'active layer');
    const t = idOf(await call('create_text_frame', { content: 'Say "hi" and a\\b', paragraphStyle: 'Body', y: 120 }));
    out = await call('find_replace_text', { findText: 'Say "hi"', replaceText: 'Said', caseSensitive: true });
    assert(/Replaced 1 instance/.test(out), out);
    assert(/Said and a\\b/.test(await call('get_text_content', { frameId: t })), 'replacement applied, backslash intact');
    await call('zoom_to_page', { pageIndex: 0 });
    await call('view_document');
    await call('list_styles');
    await call('list_color_swatches');
    out = await call('preflight_document');
    assert(/Preflight with profile/.test(out), out);
    await call('analyze_text_problems', { frameIndex: 0 });
    assert(/EMBEDDED OBJECTS ANALYSIS/.test(await call('analyze_embedded_objects', { frameId: t })), 'analyze_embedded_objects accepts frameId');
    await fails('analyze_embedded_objects', { frameId: 999999 }, /No object with id/);
    await call('delete_object', { id: t });
  });


  console.log('\nFixed legacy tools');
  if (flyer) await call('activate_document', { name: flyer });
  await test('tables: create_table and populate_table', async () => {
    let out = await call('create_table', { x: 10, y: 10, width: 100, height: 40, rows: 4, columns: 3, headerRows: 1, footerRows: 1 });
    const frame = idOf(out);
    assert(/rows=4 columns=3 \(header 1, footer 1\)/.test(out), out);
    out = await call('populate_table', { tableIndex: 0, data: [['H1', 'H2', 'H3'], ['a', null, 3]] });
    assert(/6 cell\(s\) written/.test(out), out);
    const cells = inspect(`var t = app.activeDocument.pageItems.itemByID(${frame}).tables[0]; return [t.rows[0].cells[0].contents, t.rows[1].cells[0].contents, "[" + t.rows[1].cells[1].contents + "]", t.rows[1].cells[2].contents].join("|");`);
    eq(cells, 'H1|a|[]|3', 'cells hold the data (null becomes empty)');
    await call('populate_table', { tableIndex: 0, data: [['body']], includeHeaders: false });
    eq(inspect(`var t = app.activeDocument.pageItems.itemByID(${frame}).tables[0]; return t.rows[0].cells[0].contents + "|" + t.rows[1].cells[0].contents;`), 'H1|body', 'includeHeaders=false starts below the header row');
    await fails('populate_table', { tableIndex: 5, data: [['x']] }, /Table index 5 not found/);
    await fails('create_table', { x: 0, y: 0, width: 50, height: 20, rows: 1, columns: 2, headerRows: 1 }, /at least/);
    await call('delete_object', { id: frame });
  });

  await test('typography fixes keep the formatting (GREP, not story rewrites)', async () => {
    const id = idOf(await call('create_text_frame', { content: 'Am 5. 3. 2024 sagte er "Hallo" -- und  "Welt" - ok  \nIch bin da', paragraphStyle: 'Body', x: 10, y: 60, width: 120, height: 40 }));
    await call('apply_character_style_to_range', { frameId: id, styleName: 'Bold', matchText: 'sagte' });
    const out = await call('fix_typography_in_selection', { frameId: id });
    assert(/date/i.test(out) && /opening quote/.test(out) && /em dash/.test(out) && /multiple space/.test(out), out);
    const text = await call('get_text_content', { frameId: id, normalizeSpaces: false });
    assert(text.includes('5. 3. 2024'), `date uses thin spaces: ${text}`);
    assert(text.includes('„Hallo“') && text.includes('„Welt“'), `German quotes: ${text}`);
    assert(text.includes('—') && text.includes(' – ') && !/ {2}/.test(text.replace(/\s+$/, '')), `dashes and spaces: ${text}`);
    eq(inspect(`var s = app.activeDocument.pageItems.itemByID(${id}).parentStory; var i = s.contents.indexOf("sagte"); return s.characters[i].appliedCharacterStyle.name;`), 'Bold', 'the character style survived the fixes');
    const again = await call('fix_typography_in_selection', { frameId: id });
    assert(/No typography issues found/.test(again), `second run finds nothing: ${again}`);
    assert(/No typography issues|Total/.test(await call('find_typography_issues', { frameId: id })), 'find_typography_issues runs');
    await call('delete_object', { id });
  });

  await test('clean_imported_text keeps words and character styles', async () => {
    const id = idOf(await call('create_text_frame', { content: 'Ich  bin da 1990-2000   \r\r\r- Punkt eins\rKapitel 3: Titel\rIV. Roemisch\rVier Freunde', paragraphStyle: 'Body', x: 10, y: 60, width: 120, height: 60 }));
    await call('apply_character_style_to_range', { frameId: id, styleName: 'Italic', matchText: 'Freunde' });
    await call('edit_text_frame', { frameId: id, fontSize: 40 }); // a local override
    const out = await call('clean_imported_text', { frameId: id });
    assert(/chapter number/.test(out) && /roman numeral/.test(out) && /bullet/.test(out), out);
    const text = await call('get_text_content', { frameId: id, normalizeSpaces: false });
    assert(text.includes('Ich bin da 1990–2000') && text.includes('Punkt eins') && text.includes('Titel') && text.includes('Roemisch') && text.includes('Vier Freunde'), `words such as "Ich" and "Vier" stay intact: ${text}`);
    assert(!/Kapitel|IV\.|- Punkt/.test(text), text);
    eq(inspect(`var s = app.activeDocument.pageItems.itemByID(${id}).parentStory; var i = s.contents.indexOf("Freunde"); return s.characters[i].appliedCharacterStyle.name + "," + Math.round(s.characters[0].pointSize);`), 'Italic,10', 'character style kept, local size override cleared');
    assert(/No issues found/.test(await call('clean_imported_text', { frameId: id, fixFormatting: false })), 'second run is clean');
    await call('analyze_text_problems', { frameId: id });
    await call('delete_object', { id });
  });

  await test('character and object styles: validation, zero values, cleanup', async () => {
    await fails('create_character_style', { name: 'Bold' }, /already exists/);
    await fails('create_character_style', { name: 'Ghost', fontFamily: 'NoSuchFontXYZ' }, /Font not installed/);
    await fails('create_character_style', { name: 'Ghost', textColor: 'NoSuchSwatch' }, /Swatch not found/);
    await fails('create_character_style', { name: 'Ghost', baseStyle: 'NoSuchStyle' }, /not found/);
    assert(!/Ghost/.test(await call('list_styles', { styleType: 'character' })), 'failed creates leave no half-made style');
    await call('create_character_style', { name: 'Tracked', tracking: 25, textColor: 'Brand' });
    let out = await call('modify_character_style', { styleName: 'Tracked', tracking: 0 });
    assert(/tracking=0/.test(out), `zero is applied, not skipped: ${out}`);
    await fails('modify_character_style', { styleName: 'Tracked' }, /No properties/);
    await fails('modify_character_style', { styleName: 'Nope', tracking: 1 }, /not found/);
    out = await call('modify_object_style', { styleName: 'OS1', transparency: 0 });
    assert(/opacity=100%/.test(out), out);
    await fails('create_object_style', { name: 'OS1' }, /already exists/);
    await fails('create_object_style', { name: 'Ghost2', fillColor: 'NoSuchSwatch' }, /Swatch not found/);
    const rect = idOf(await call('create_rectangle', { x: 10, y: 130, width: 10, height: 10 }));
    out = await call('apply_object_style', { styleName: 'OS1', objectId: rect });
    assert(/applied to 1 object/.test(out), out);
    await fails('apply_object_style', { styleName: 'OS1', objectId: 987654 }, /No object with id/);
    await fails('apply_object_style', { styleName: 'Nope', objectId: rect }, /not found/);
    out = await call('apply_color', { objectId: rect, swatchName: 'Orange', property: 'stroke' });
    assert(/applied to stroke/.test(out), out);
    await fails('apply_color', { objectId: rect, swatchName: 'NoSuchSwatch' }, /Swatch not found/);
    await fails('apply_color', { objectIndex: 999, swatchName: 'Orange' }, /Invalid index/);
    await call('delete_object', { id: rect });
  });

  await test('layers, zoom, find/replace scopes report clear errors', async () => {
    await call('create_layer', { name: 'Colour "L"', color: 'RED' });
    await fails('create_layer', { name: 'Colour "L"' }, /already exists/);
    await fails('create_layer', { name: 'Ghost', color: 'NO_SUCH_COLOUR' }, /Unknown layer colour/);
    assert(!/Ghost/.test(await call('list_layers')), 'failed create leaves no layer');
    await call('set_active_layer', { layerName: 'Colour "L"' });
    await fails('set_active_layer', { layerName: 'NoSuchLayer' }, /not found/);
    await fails('zoom_to_page', { pageIndex: 99 }, /Invalid page index/);
    await fails('zoom_to_page', { fitOption: 'ZOOM_TO_SELECTION' }, /Invalid parameter "fitOption"/);
    await fails('find_replace_text', { findText: 'a', replaceText: 'b', scope: 'story' }, /needs a selection/);
    await fails('find_replace_text', { findText: 'a', replaceText: 'b', scope: 'selection' }, /needs something selected/);
  });


  await test('duplicate_page copies every object once (groups stay intact)', async () => {
    const name = await newDoc({ preset: 'A5' });
    const r = idOf(await call('create_rectangle', { x: 10, y: 10, width: 20, height: 20 }));
    const e = idOf(await call('create_ellipse', { x: 40, y: 10, width: 20, height: 20 }));
    await call('group_objects', { ids: [r, e] });
    const count = async (pg) => Number((await call('list_page_items', { pageIndex: pg })).match(/PAGE ITEMS ON PAGE \d+ \((\d+)\)/)[1]);
    const source = await count(0);
    const out = await call('duplicate_page', { pageIndex: 0 });
    assert(/New page position: 2/.test(out), out);
    eq(await count(1), source, 'the copy has as many objects as the source');
    const end = await call('duplicate_page', { pageIndex: 0, position: 'end' });
    assert(/New page position: 3/.test(end) && /Total pages: 3/.test(end), end);
    await fails('duplicate_page', { pageIndex: 0, position: 'sideways' }, /Invalid parameter "position"/);
    await call('delete_page', { pageIndex: 2, confirmDestructive: true });
    await call('delete_page', { pageIndex: 1, confirmDestructive: true });
    await fails('delete_page', { pageIndex: 0, confirmDestructive: true }, /Cannot delete the last page/);
    await closeDoc(name);
    await call('activate_document', { name: flyer });
  });

  await test('export_epub', async () => {
    const epub = path.join(outDir, 'sub', 'book.epub');
    const out = await call('export_epub', { filePath: epub, version: 'EPUB2', imageFormat: 'PNG', confirmDestructive: true });
    assert(/EPUB exported/.test(out) && fs.existsSync(epub) && fs.statSync(epub).size > 100, out);
    await fails('export_epub', { filePath: epub }, /confirm/i);
  });

  await test('package_document (needs a saved document) and save_document', async () => {
    const name = await newDoc({ preset: 'A5' });
    await call('create_text_frame', { content: 'Package me', y: 20 });
    await fails('package_document', { folderPath: path.join(TMP, 'pkg'), confirmDestructive: true }, /never been saved/);
    await call('save_document', { filePath: path.join(TMP, 'saved-doc.indd'), confirmDestructive: true });
    const saved = (await call('list_open_documents')).match(/\] (\S+\.indd) \(ACTIVE\)/)?.[1];
    assert(saved === 'saved-doc.indd', `the document was renamed by saving: ${saved}`);
    createdDocs[createdDocs.indexOf(name)] = saved;
    const pkg = await call('package_document', { folderPath: path.join(TMP, 'pkg'), confirmDestructive: true });
    assert(/Document packaged/.test(pkg) && fs.readdirSync(path.join(TMP, 'pkg')).length > 0, pkg);
    await closeDoc(saved);
    await call('activate_document', { name: flyer });
  });

  await test('data_merge: PDF, INDD, record ranges and errors', async () => {
    const csv = path.join(TMP, 'data.csv');
    fs.writeFileSync(csv, 'name,city\nAnna,Wien\nBerta,Graz\nCarl,Linz\n');
    const name = await newDoc({ preset: 'A5' });
    // placeholders are set up here (test-only); the tool does the merge
    inspect(`var d = app.activeDocument; var tf = d.pages[0].textFrames.add(); tf.geometricBounds = [10, 10, 50, 100]; tf.contents = "Hello ";
      var dm = d.dataMergeProperties; dm.selectDataSource(File(${JSON.stringify(csv)})); var st = tf.parentStory;
      d.dataMergeTextPlaceholders.add(st, st.insertionPoints[-1].index, dm.dataMergeFields[0]); return "ok";`);
    const out = path.join(TMP, 'merged');
    await fails('data_merge', { dataSourcePath: csv, outputFolder: out }, /confirm/i);
    let r = await call('data_merge', { dataSourcePath: csv, outputFolder: out, confirmDestructive: true });
    const pdf = path.join(out, `${name.replace(/\.indd$/i, '')}_merged.pdf`);
    assert(/Data merge completed/.test(r) && fs.existsSync(pdf), r);
    eq(pdfBoxes(pdf).pages, 3, 'one page per record');
    r = await call('data_merge', { dataSourcePath: csv, outputFolder: out, recordRange: '1-2', confirmDestructive: true });
    eq(pdfBoxes(pdf).pages, 2, 'record range 1-2 gives two pages');
    r = await call('data_merge', { dataSourcePath: csv, outputFolder: out, recordRange: '3', fileFormat: 'BOTH', confirmDestructive: true });
    assert(fs.existsSync(path.join(out, `${name.replace(/\.indd$/i, '')}_merged.indd`)), `INDD written: ${r}`);
    eq(pdfBoxes(pdf).pages, 1, 'a single record gives one page');
    assert(new RegExp(name.replace(/[-.]/g, '\\$&') + '[^\\n]*\\(ACTIVE\\)').test(await call('list_open_documents')), 'the original document is active again');
    await fails('data_merge', { dataSourcePath: path.join(TMP, 'missing.csv'), outputFolder: out, confirmDestructive: true }, /not found/);
    await fails('data_merge', { dataSourcePath: csv, outputFolder: out, recordRange: 'first', confirmDestructive: true }, /Invalid recordRange/);
    await closeDoc(name);
    await call('activate_document', { name: flyer });
  });

  await test('paths: symlinks cannot lead outside the allowed directories', async () => {
    const link = path.join(TMP, 'sneaky.png');
    fs.symlinkSync('/etc/hosts', link);
    await fails('place_image', { imagePath: link }, /Access denied/);
    const dirLink = path.join(TMP, 'sneaky-dir');
    fs.symlinkSync('/etc', dirLink);
    await fails('export_pdf', { filePath: path.join(dirLink, 'x.pdf'), confirmDestructive: true }, /Access denied/);
  });


  console.log('\nVector shapes (Part A)');
  let vec;             // document for the vector tests
  const VEC_SVG = 'M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z';
  await test('acceptance: the curved flyer band from SVG path data', async () => {
    vec = await newDoc({ preset: 'A5', orientation: 'Portrait', bleed: 3, facingPages: true, pages: 2 });
    testDoc = vec;
    await call('create_color_swatch', { name: 'BandGreen', colorValues: [85, 10, 100, 20] });
    const out = await call('create_path_from_svg', { d: VEC_SVG, viewBox: '0 0 216 50', x: -3, y: 55, width: 216, height: 48, fill: 'BandGreen', opacity: 85, label: 'band' });
    const band = idOf(out);
    assert(band > 0 && /Polygon/.test(out), out);
    assert(/points=5 paths=1 closed=true/.test(out), `5 anchors, closed: ${out}`);
    assert(/fill=BandGreen/.test(out) && /opacity=85/.test(out) && /label=band/.test(out), `fill swatch, 85% opacity, label: ${out}`);
    const pts = readPoints(await call('edit_path_points', { id: band, action: 'read' }));
    eq(pts.length, 5, 'five points read back');
    // expected geometry: the SVG scaled by 216/216 in x and 48/50 in y, then moved to (-3, 55)
    const X = (x) => -3 + x, Y = (y) => 55 + y * 0.96;
    const expected = [[0, 8], [120, 8], [216, 6], [216, 50], [0, 50]];
    expected.forEach(([x, y], i) => { near(pts[i].x, X(x), 0.01, `anchor ${i} x`); near(pts[i].y, Y(y), 0.01, `anchor ${i} y`); });
    near(pts[0].rx, X(40), 0.01, 'handle 0 right x'); near(pts[0].ry, Y(0), 0.01, 'handle 0 right y');
    near(pts[1].lx, X(80), 0.01, 'handle 1 left x'); near(pts[1].ly, Y(14), 0.01, 'handle 1 left y');
    near(pts[1].rx, X(160), 0.01, 'S command reflects the handle: right x'); near(pts[1].ry, Y(2), 0.01, 'S reflection: right y');
    near(pts[2].lx, X(190), 0.01, 'handle 2 left x'); near(pts[2].ly, Y(0), 0.01, 'handle 2 left y');
    // the top edge is a gentle S-curve: sample it and compare with the straight chord between the end anchors
    const top = [...sampleY(pts[0], pts[1]), ...sampleY(pts[1], pts[2])];
    const chordY = (i, n) => pts[0].y + ((pts[2].y - pts[0].y) * i) / n;
    const dev = Math.max(...top.map((y, i) => Math.abs(y - chordY(i, top.length - 1))));
    assert(dev > 0.5 && dev < 8, `the edge deviates from a straight line by ${dev.toFixed(2)} mm (a gentle curve)`);
    const info = geo(await call('get_object_info', { id: band }));
    near(info.y, Math.min(...top), 0.05, 'the path bounds start at the top of the sampled curve (real curve geometry)');
    near(info.x, -3, 0.01, 'left edge reaches into the bleed');
    near(info.w, 216, 0.01, 'width 216');
    near(info.y + info.h, Y(50), 0.01, 'bottom edge');
    const full = await call('get_object_info', { id: band });
    assert(/fill: BandGreen \[CMYK 85,10,100,20\]/.test(full) && /opacity=85/.test(full), `object info shows swatch values and opacity: ${full}`);
    await call('delete_object', { id: band });
  });

  await test('create_path -> edit_path_points read: the same points come back (within 0.01 mm)', async () => {
    const points = [
      { x: 10, y: 10 },
      { x: 60.123, y: 12.5, leftDirection: { x: 45.5, y: 5.25 }, rightDirection: { x: 70.75, y: 20 }, pointType: 'corner' },
      { x: 80, y: 60, leftDirection: { x: 90, y: 50 }, rightDirection: { x: 70, y: 70 }, pointType: 'smooth' },
      { x: 30.333, y: 70.667 },
    ];
    const out = await call('create_path', { points, closed: true, fill: 'BandGreen', stroke: 'Black', strokeWeight: 2, opacity: 60, label: 'roundtrip' });
    const id = idOf(out);
    assert(/points=4 paths=1 closed=true/.test(out) && /opacity=60/.test(out) && /stroke=Black\/2pt/.test(out), out);
    const back = readPoints(await call('edit_path_points', { id }));
    eq(back.length, 4, 'four points');
    points.forEach((p, i) => {
      near(back[i].x, p.x, 0.01, `point ${i} x`); near(back[i].y, p.y, 0.01, `point ${i} y`);
      near(back[i].lx, p.leftDirection?.x ?? p.x, 0.01, `point ${i} left x`); near(back[i].ly, p.leftDirection?.y ?? p.y, 0.01, `point ${i} left y`);
      near(back[i].rx, p.rightDirection?.x ?? p.x, 0.01, `point ${i} right x`); near(back[i].ry, p.rightDirection?.y ?? p.y, 0.01, `point ${i} right y`);
    });
    eq(back[1].type, 'CORNER', 'explicit corner type kept'); eq(back[2].type, 'SMOOTH', 'explicit smooth type kept');
    // open path, and a page-relative position on the second page of a facing-pages spread
    const open = await call('create_path', { points: [{ x: 5, y: 5 }, { x: 50, y: 5 }, { x: 50, y: 40 }], pageIndex: 1 });
    assert(/points=3 paths=1 closed=false/.test(open) && /page=2/.test(open), open);
    const g = geo(await call('get_object_info', { id: idOf(open) }));
    near(g.x, 5, 0.01, 'page-relative x on page 2'); near(g.y, 5, 0.01, 'page-relative y');
    await fails('create_path', { points: [{ x: 1, y: 1 }] }, /at least 2/);
    await fails('create_path', { points: [{ x: 1, y: 1 }, { x: 'a', y: 2 }] }, /Invalid points\[1\]\.x/);
    await fails('create_path', { points: [{ x: 1, y: 1 }, { x: 2, y: 2, pointType: 'wobbly' }] }, /Invalid pointType/);
    await fails('create_path', { points: [{ x: 1, y: 1 }, { x: 2, y: 2 }], fill: 'NoSuchSwatch' }, /Swatch not found/);
    assert(!/label=|NoSuch/.test((await call('list_page_items', { pageIndex: 0 })).split('\n').filter((l) => /label=undefined/.test(l)).join('')), 'no stray object after a failed create');
    await call('delete_object', { id }); await call('delete_object', { id: idOf(open) });
  });

  await test('edit_path_points: add, move, delete, set_type, reverse, set_closed', async () => {
    const id = idOf(await call('create_path', { points: [{ x: 10, y: 10 }, { x: 50, y: 10 }, { x: 50, y: 50 }, { x: 10, y: 50 }], closed: true }));
    let r = readPoints(await call('edit_path_points', { id, action: 'add', pointIndex: 2, point: { x: 60, y: 30, leftDirection: { x: 60, y: 20 }, rightDirection: { x: 60, y: 40 }, pointType: 'smooth' } }));
    eq(r.length, 5, 'a point was added'); near(r[2].x, 60, 0.01, 'inserted at index 2'); near(r[3].x, 50, 0.01, 'the old point 2 moved to index 3');
    r = readPoints(await call('edit_path_points', { id, action: 'move', pointIndex: 2, x: 65, y: 31 }));
    near(r[2].x, 65, 0.01, 'anchor moved x'); near(r[2].y, 31, 0.01, 'anchor moved y'); near(r[2].lx, 65, 0.01, 'handles moved with the anchor (60 + 5)'); near(r[2].ly, 21, 0.01, 'left handle y (20 + 1)');
    r = readPoints(await call('edit_path_points', { id, action: 'move', pointIndex: 0, x: 12, y: 8, moveHandles: false }));
    near(r[0].x, 12, 0.01, 'move without handles');
    r = readPoints(await call('edit_path_points', { id, action: 'move', pointIndex: 1, rightDirection: { x: 55, y: 12 }, leftDirection: { x: 45, y: 8 } }));
    near(r[1].rx, 55, 0.01, 'right handle set'); near(r[1].lx, 45, 0.01, 'left handle set'); near(r[1].x, 50, 0.01, 'anchor untouched');
    r = readPoints(await call('edit_path_points', { id, action: 'set_type', pointIndex: 1, pointType: 'corner' }));
    eq(r[1].type, 'CORNER', 'type changed');
    r = readPoints(await call('edit_path_points', { id, action: 'set_type', pointIndex: 3, pointType: 'plain' }));
    eq(r[3].type, 'PLAIN', 'plain type removes the handles'); near(r[3].lx, r[3].x, 0.01, 'no handle left');
    const before = readPoints(await call('edit_path_points', { id }));
    const rev = readPoints(await call('edit_path_points', { id, action: 'reverse' }));
    eq(rev.length, before.length, 'same number of points');
    near(rev[0].x, before[before.length - 1].x, 0.01, 'order reversed'); near(rev[0].lx, before[before.length - 1].rx, 0.01, 'handles swapped');
    r = readPoints(await call('edit_path_points', { id, action: 'delete', pointIndex: 2 }));
    eq(r.length, 4, 'a point was deleted');
    let out = await call('edit_path_points', { id, action: 'set_closed', closed: false });
    assert(/closed=false/.test(out), out);
    out = await call('edit_path_points', { id, action: 'set_closed', closed: true });
    assert(/closed=true/.test(out), out);
    await fails('edit_path_points', { id, action: 'delete', pointIndex: 99 }, /Invalid pointIndex/);
    await fails('edit_path_points', { id, action: 'move', pointIndex: 0 }, /needs x\/y/);
    await fails('edit_path_points', { id, action: 'add' }, /needs point/);
    await fails('edit_path_points', { id, pathIndex: 3 }, /Invalid pathIndex/);
    await fails('edit_path_points', { id, action: 'set_type', pointIndex: 0, pointType: 'wobbly' }, /Invalid (parameter|pointType)/);
    await call('edit_path_points', { id, action: 'delete', pointIndex: 0 });
    await fails('edit_path_points', { id, action: 'delete', pointIndex: 0 }, /needs at least 3/);
    const text = idOf(await call('create_text_frame', { content: 'x', y: 150 }));
    await fails('edit_path_points', { id: text }, /has no editable path/);
    await call('delete_object', { id: text }); await call('delete_object', { id });
  });

  await test('compound paths from SVG data and edit of a rectangle\'s points', async () => {
    const out = await call('create_path_from_svg', { d: 'M0 0 H10 V10 H0 Z M2 2 V8 H8 V2 Z', x: 100, y: 100, width: 40, height: 40, fill: 'BandGreen' });
    assert(/paths=2 closed=true/.test(out) && /points=8/.test(out), `compound path with two sub-paths: ${out}`);
    const id = idOf(out);
    const p1 = readPoints(await call('edit_path_points', { id, pathIndex: 1 }));
    eq(p1.length, 4, 'the hole has 4 points');
    await fails('create_path_from_svg', { d: 'M 0 0 X' }, /Invalid SVG path data/);
    await fails('create_path_from_svg', { d: 'M0 0 L10 10', viewBox: '0 0 0 5' }, /viewBox/);
    const rect = idOf(await call('create_rectangle', { x: 100, y: 150, width: 30, height: 20 }));
    const grown = readPoints(await call('edit_path_points', { id: rect, action: 'add', pointIndex: 2, point: { x: 130, y: 190 } }));
    eq(grown.length, 5, 'rectangles are editable paths too');
    await call('delete_object', { id }); await call('delete_object', { id: rect });
  });

  await test('create_line and create_polygon', async () => {
    let out = await call('create_line', { x1: 10, y1: 10, x2: 100, y2: 10, weight: 3, stroke: 'BandGreen', endCap: 'round', arrowEnd: 'triangle', dash: [6, 3], opacity: 50, label: 'ln' });
    const line = idOf(out);
    assert(/GraphicLine/.test(out) && /points=2 paths=1 closed=false/.test(out) && /stroke=BandGreen\/3pt/.test(out), out);
    eq(inspect(`var l = app.activeDocument.pageItems.itemByID(${line}); return [String(l.rightLineEnd), String(l.leftLineEnd), String(l.endCap), l.strokeType.name, l.strokeDashAndGap.join("/")].join(",");`), 'TRIANGLE_ARROW_HEAD,NONE,ROUND_END_CAP,Dashed,6/3', 'arrowhead, cap and dash applied');
    out = await call('create_line', { x1: 10, y1: 30, x2: 100, y2: 30, controlPoint1: { x: 40, y: 10 }, controlPoint2: { x: 70, y: 50 }, dashStyle: 'dotted' .replace('d', 'D') });
    const curve = readPoints(await call('edit_path_points', { id: idOf(out) }));
    near(curve[0].rx, 40, 0.01, 'curved line handle 1 x'); near(curve[0].ry, 10, 0.01, 'handle 1 y'); near(curve[1].lx, 70, 0.01, 'handle 2 x'); near(curve[1].ly, 50, 0.01, 'handle 2 y');
    assert(/Dotted/.test(inspect(`return app.activeDocument.pageItems.itemByID(${idOf(out)}).strokeType.name;`)), 'named stroke style applied');
    await fails('create_line', { x1: 0, y1: 0, x2: 1, y2: 1, arrowEnd: 'rocket' }, /Invalid parameter "arrowEnd"/);
    await fails('create_line', { x1: 0, y1: 0, x2: 1, y2: 1, dashStyle: 'Nonexistent Style' }, /Stroke style not found/);
    for (const [args, points, label] of [
      [{ x: 50, y: 100, radius: 20, sides: 6 }, 6, 'hexagon'],
      [{ x: 50, y: 100, radius: 20, sides: 5, innerRadiusRatio: 0.4 }, 10, 'star'],
      [{ x: 50, y: 100, radius: 20, sides: 6, cornerRadius: 3 }, 12, 'rounded hexagon'],
    ]) {
      const o = await call('create_polygon', { ...args, fill: 'BandGreen' });
      assert(new RegExp(`points=${points} paths=1 closed=true`).test(o), `${label}: ${o}`);
      const g = geo(await call('get_object_info', { id: idOf(o) }));
      near(g.x + g.w / 2, 50, 1.2, `${label} is centred`);
      assert(g.w <= 40.01 && g.h <= 40.01 && g.w > 30, `${label} fits its circle: ${g.w} x ${g.h}`);
      await call('delete_object', { id: idOf(o) });
    }
    await fails('create_polygon', { x: 0, y: 0, radius: 5, sides: 2 }, /sides must be/);
    await fails('create_polygon', { x: 0, y: 0, radius: 5, innerRadiusRatio: 1.5 }, /between 0 and 1/);
    await fails('create_polygon', { x: 0, y: 0, radius: 0 }, /radius must be/);
    await call('delete_object', { id: line }); await call('delete_object', { id: idOf(out) });
  });

  await test('convert_shape and set_corner_options', async () => {
    const rect = idOf(await call('create_rectangle', { x: 20, y: 120, width: 60, height: 40 }));
    let out = await call('set_corner_options', { id: rect, option: 'rounded', radius: 5 });
    assert(/topLeft=ROUNDED_CORNER\/5mm/.test(out) && /bottomRight=ROUNDED_CORNER\/5mm/.test(out), out);
    out = await call('set_corner_options', { id: rect, corners: { top_left: { option: 'bevel', radius: 3 }, bottom_right: { option: 'inverse_rounded', radius: 4 } } });
    assert(/topLeft=BEVEL_CORNER\/3mm/.test(out) && /bottomRight=INVERSE_ROUNDED_CORNER\/4mm/.test(out) && /topRight=ROUNDED_CORNER\/5mm/.test(out), `per corner settings: ${out}`);
    await call('set_corner_options', { id: rect, option: 'none' });
    out = await call('convert_shape', { id: rect, to: 'polygon', sides: 6, insetPercent: 0, cornerRadius: 0 });
    assert(/Converted to polygon/.test(out) && /points=6/.test(out), out);
    out = await call('convert_shape', { id: rect, to: 'triangle' });
    assert(/points=3/.test(out), out);
    out = await call('convert_shape', { id: rect, to: 'oval' });
    assert(/points=4/.test(out), out);
    await fails('set_corner_options', { id: rect }, /Pass option/);
    await fails('set_corner_options', { id: rect, corners: { middle: { option: 'rounded' } } }, /Unknown corner/);
    await fails('convert_shape', { id: rect, to: 'blob' }, /Invalid parameter "to"/);
    await call('delete_object', { id: rect });
  });

  await test('pathfinder: subtract, union, intersect, exclude_overlap, minus_back', async () => {
    const fresh = async () => {
      const a = idOf(await call('create_rectangle', { x: 10, y: 10, width: 60, height: 40 }));
      const b = idOf(await call('create_ellipse', { x: 60, y: 20, width: 30, height: 30 }));
      return [a, b];
    };
    let [a, b] = await fresh();
    let out = await call('pathfinder', { operation: 'subtract', ids: [a, b] });
    const sub = idOf(out.match(/Result: (.*)/)[1]);
    assert(/Rectangle/.test(out) && /points=7 /.test(out), `a rectangle with a circular notch has 7 points: ${out}`);
    const g = geo(await call('get_object_info', { id: sub }));
    near(g.x, 10, 0.01, 'result keeps the rectangle bounds'); near(g.w, 60, 0.01, 'width 60');
    await fails('get_object_info', { id: b }, /No object with id/);
    await call('delete_object', { id: sub });

    [a, b] = await fresh();
    out = await call('pathfinder', { operation: 'union', ids: [a, b] });
    assert(/points=9 /.test(out), `union: ${out}`);
    const u = geo(await call('get_object_info', { id: idOf(out.match(/Result: (.*)/)[1]) }));
    near(u.x + u.w, 90, 0.01, 'union reaches the far side of the circle');
    await call('delete_object', { id: idOf(out.match(/Result: (.*)/)[1]) });

    [a, b] = await fresh();
    out = await call('pathfinder', { operation: 'intersect', ids: [a, b] });
    const i = geo(await call('get_object_info', { id: idOf(out.match(/Result: (.*)/)[1]) }));
    near(i.x, 60, 0.01, 'intersection starts at the circle'); near(i.x + i.w, 70, 0.01, 'and ends at the rectangle edge');
    await call('delete_object', { id: idOf(out.match(/Result: (.*)/)[1]) });

    [a, b] = await fresh();
    out = await call('pathfinder', { operation: 'exclude_overlap', ids: [a, b] });
    assert(/paths=2/.test(out), `exclude overlap leaves two sub-paths: ${out}`);
    await call('delete_object', { id: idOf(out.match(/Result: (.*)/)[1]) });

    [a, b] = await fresh();   // b (the circle) is in front
    out = await call('pathfinder', { operation: 'minus_back', ids: [a, b] });
    const mb = geo(await call('get_object_info', { id: idOf(out.match(/Result: (.*)/)[1]) }));
    near(mb.x, 70, 0.01, 'minus_back keeps the frontmost shape (the circle) minus the rectangle behind it'); near(mb.x + mb.w, 90, 0.01, 'right edge of the circle');
    await call('delete_object', { id: idOf(out.match(/Result: (.*)/)[1]) });

    [a, b] = await fresh();
    const far = idOf(await call('create_ellipse', { x: 120, y: 120, width: 10, height: 10 }));
    await fails('pathfinder', { operation: 'intersect', ids: [a, far] }, /empty/);
    await fails('pathfinder', { operation: 'subtract', ids: [a] }, /at least 2/);
    await fails('pathfinder', { operation: 'subtract', ids: [a, 999999] }, /No object with id/);
    const t = idOf(await call('create_text_frame', { content: 'x', y: 170 }));
    await fails('pathfinder', { operation: 'union', ids: [a, t] }, /no path|cannot be used/);
    for (const id of [a, b, far, t]) { try { await call('delete_object', { id }); } catch { /* consumed */ } }
  });

  await test('set_object_fill / set_object_stroke change objects in place (stacking order unchanged)', async () => {
    await call('create_color_swatch', { name: 'Accent', colorValues: [0, 80, 100, 0] });
    const ids = [];
    for (let k = 0; k < 3; k++) ids.push(idOf(await call('create_rectangle', { x: 10 + k * 20, y: 10 + k * 5, width: 40, height: 20, fillColor: 'BandGreen' })));
    const order = () => inspect(`var it = app.activeDocument.pages[0].allPageItems, o = []; for (var i = 0; i < it.length; i++) o.push(it[i].id); return o.join(",");`);
    const before = order();
    let out = await call('set_object_fill', { id: ids[1], swatch: 'Accent', tint: 50, opacity: 70 });
    assert(/fill=Accent\/50/.test(out), out);
    eq(order(), before, 'the stacking order is unchanged after changing the fill');
    const info = await call('get_object_info', { id: ids[1] });
    assert(/fill: Accent \[CMYK 0,80,100,0\] tint=50/.test(info), info);
    assert(/fillOpacity=70/.test(info), `fill opacity: ${info}`);
    out = await call('set_object_fill', { ids: [ids[0], ids[2]], swatch: 'Accent' });
    eq(out.split('\n').filter((l) => /fill=Accent/.test(l)).length, 2, 'several objects at once');
    await call('set_object_fill', { id: ids[0], swatch: 'none' });
    assert(/fill: None/.test(await call('get_object_info', { id: ids[0] })), 'fill none');
    out = await call('set_object_stroke', { id: ids[1], swatch: 'Black', weight: 4, tint: 80, alignment: 'inside', join: 'round', cap: 'round', miterLimit: 6, dash: [8, 4], opacity: 40 });
    assert(/stroke=Black\/4pt/.test(out), out);
    eq(order(), before, 'the stacking order is unchanged after changing the stroke');
    eq(inspect(`var r = app.activeDocument.pageItems.itemByID(${ids[1]}); return [String(r.strokeAlignment), String(r.endJoin), String(r.endCap), r.miterLimit, r.strokeType.name, r.strokeDashAndGap.join("/"), Math.round(r.strokeTint), Math.round(r.strokeTransparencySettings.blendingSettings.opacity)].join(",");`), 'INSIDE_ALIGNMENT,ROUND_END_JOIN,ROUND_END_CAP,6,Dashed,8/4,80,40', 'all stroke properties applied');
    await call('set_object_stroke', { id: ids[1], dashStyle: 'solid' });
    assert(/Solid/.test(inspect(`return app.activeDocument.pageItems.itemByID(${ids[1]}).strokeType.name;`)), 'back to solid');
    const tf = idOf(await call('create_text_frame', { content: 'framed', y: 150, width: 60, height: 20 }));
    await call('set_object_fill', { id: tf, swatch: 'Accent' });
    await call('set_object_stroke', { id: tf, swatch: 'Black', weight: 1 });
    assert(/fill: Accent/.test(await call('get_object_info', { id: tf })), 'text frames work too');
    await fails('set_object_fill', { id: ids[0] }, /Pass swatch/);
    await fails('set_object_fill', { id: ids[0], swatch: 'NoSuchSwatch' }, /Swatch not found/);
    await fails('set_object_fill', { id: ids[0], swatch: 'Accent', gradient: 'G' }, /either swatch or gradient/);
    await fails('set_object_fill', { id: ids[0], tint: 150 }, /0-100/);
    await fails('set_object_stroke', { id: ids[0] }, /at least one/);
    await fails('set_object_stroke', { id: ids[0], alignment: 'sideways' }, /Invalid parameter "alignment"/);
    await fails('set_object_stroke', { id: ids[0], dashStyle: 'Nope' }, /Stroke style not found/);
    await fails('set_object_fill', { id: 999999, swatch: 'Accent' }, /No object with id/);
    for (const id of [...ids, tf]) await call('delete_object', { id });
  });

  await test('gradients: create_gradient_swatch, apply with angle, gradient feather', async () => {
    await call('create_color_swatch', { name: 'GradEnd', colorValues: [0, 100, 0, 0] });
    let out = await call('create_gradient_swatch', { name: 'Fade3', type: 'linear', stops: [{ swatch: 'BandGreen', position: 0 }, { swatch: 'Accent', position: 40, midpoint: 30 }, { swatch: 'GradEnd', position: 100, midpoint: 60 }] });
    assert(/3 stops/.test(out) && /0% BandGreen/.test(out) && /40% Accent midpoint=30/.test(out) && /100% GradEnd midpoint=60/.test(out), out);
    const rect = idOf(await call('create_rectangle', { x: 10, y: 10, width: 80, height: 30 }));
    out = await call('set_object_fill', { id: rect, gradient: 'Fade3', gradientAngle: 45 });
    assert(/fill=Fade3/.test(out), out);
    const info = await call('get_object_info', { id: rect });
    assert(/fill: Fade3/.test(info) && /angle=45/.test(info), info);
    out = await call('create_gradient_swatch', { name: 'Glow', type: 'radial', stops: [{ swatch: 'Accent', position: 100 }, { swatch: 'BandGreen', position: 0 }] });
    assert(/radial, 2 stops/.test(out) && /0% BandGreen/.test(out), `stops are sorted by position: ${out}`);
    await fails('create_gradient_swatch', { name: 'Fade3', stops: [{ swatch: 'Accent', position: 0 }, { swatch: 'GradEnd', position: 100 }] }, /already exists/);
    await fails('create_gradient_swatch', { name: 'X1', stops: [{ swatch: 'Accent', position: 10 }, { swatch: 'GradEnd', position: 100 }] }, /first stop must be at position 0/);
    await fails('create_gradient_swatch', { name: 'X2', stops: [{ swatch: 'Accent', position: 0 }, { swatch: 'NoSuchSwatch', position: 100 }] }, /Swatch not found/);
    await fails('create_gradient_swatch', { name: 'X3', stops: [{ swatch: 'Accent', position: 0 }] }, /at least 2/);
    assert(!/X2/.test(inspect('var a = []; for (var i = 0; i < app.activeDocument.gradients.length; i++) a.push(app.activeDocument.gradients[i].name); return a.join("|");')), 'a failed gradient leaves nothing behind');
    await fails('set_object_fill', { id: rect, gradient: 'NoSuchGradient' }, /Gradient not found/);
    out = await call('set_gradient_feather', { id: rect, type: 'linear', angle: 90, startOpacity: 100, endOpacity: 10 });
    assert(/Gradient feather set/.test(out), out);
    eq(inspect(`var f = app.activeDocument.pageItems.itemByID(${rect}).transparencySettings.gradientFeatherSettings; return [f.applied, Math.round(f.angle), Math.round(f.opacityGradientStops[1].opacity)].join(",");`), 'true,90,10', 'gradient feather applied');
    assert(/gradientFeather/.test(await call('get_object_info', { id: rect })), 'get_object_info lists the effect');
    await call('set_gradient_feather', { id: rect, remove: true });
    assert(/effects: none/.test(await call('get_object_info', { id: rect })), 'and it can be removed');
    await call('delete_object', { id: rect });
  });

  await test('place_image_in_shape: an image clipped by an editable shape', async () => {
    const star = idOf(await call('create_polygon', { x: 60, y: 60, radius: 30, sides: 5, innerRadiusRatio: 0.5 }));
    const before = geo(await call('get_object_info', { id: star }));
    const out = await call('place_image_in_shape', { shapeId: star, imagePath: png, fitOption: 'FILL_PROPORTIONALLY', contentScale: 150, contentOffsetX: -2 });
    assert(new RegExp(`id=${star}\\b`).test(out) && /image=test-image\.png/.test(out), out);
    const after = geo(await call('get_object_info', { id: star }));
    near(after.w, before.w, 0.01, 'the shape keeps its size (it is the mask)'); near(after.h, before.h, 0.01, 'height unchanged');
    assert(/image: test-image\.png \(NORMAL/.test(await call('get_object_info', { id: star })), 'linked image is reported');
    await fails('place_image_in_shape', { shapeId: star, imagePath: path.join(TMP, 'missing.png') }, /not found/);
    await fails('place_image_in_shape', { shapeId: star, imagePath: '/etc/hosts' }, /Access denied/);
    const t = idOf(await call('create_text_frame', { content: 'x', y: 150 }));
    await fails('place_image_in_shape', { shapeId: t, imagePath: png }, /cannot be used as a mask/);
    await fails('place_image_in_shape', { shapeId: 999999, imagePath: png }, /No object with id/);
    await call('delete_object', { id: star }); await call('delete_object', { id: t });
  });

  await test('duplicate_object, align_objects, distribute_objects', async () => {
    const r = idOf(await call('create_rectangle', { x: 10, y: 10, width: 20, height: 10, fillColor: 'BandGreen' }));
    let out = await call('duplicate_object', { id: r, count: 3, offsetX: 30, offsetY: 5 });
    const copies = out.match(/ids: ([\d, ]+)/)[1].split(',').map((v) => Number(v.trim()));
    eq(copies.length, 3, 'three copies');
    const pos = [];
    for (const c of copies) pos.push(geo(await call('get_object_info', { id: c })));
    near(pos[0].x, 40, 0.01, 'first copy x'); near(pos[2].x, 100, 0.01, 'third copy x (offset applies to the previous copy)'); near(pos[2].y, 25, 0.01, 'third copy y');
    out = await call('duplicate_object', { id: r, toPageIndex: 1 });
    assert(/page=2/.test(out), out);
    const onPage2 = idOf(out.match(/ids: (\d+)/)[1] ? `id=${out.match(/ids: (\d+)/)[1]}` : '');
    await fails('duplicate_object', { id: r, count: 500 }, /count must be/);

    // align
    await call('align_objects', { ids: copies, alignment: 'top', relativeTo: 'selection' });
    const tops = [];
    for (const c of copies) tops.push(geo(await call('get_object_info', { id: c })).y);
    assert(tops.every((y) => Math.abs(y - tops[0]) < 0.01), `aligned tops: ${tops}`);
    await call('align_objects', { ids: [copies[0]], alignment: 'horizontal_center', relativeTo: 'page' });
    const c0 = geo(await call('get_object_info', { id: copies[0] }));
    near(c0.x + c0.w / 2, 74, 0.05, 'centred on the A5 page (148 mm)');
    await call('align_objects', { ids: [copies[1], copies[2]], alignment: 'left', relativeTo: 'key_object', keyObjectId: copies[2] });
    near(geo(await call('get_object_info', { id: copies[1] })).x, geo(await call('get_object_info', { id: copies[2] })).x, 0.01, 'aligned to the key object');
    await fails('align_objects', { ids: [copies[0]], alignment: 'left' }, /at least 2/);
    await fails('align_objects', { ids: copies, alignment: 'left', relativeTo: 'key_object' }, /keyObjectId/);

    // distribute with fixed spacing
    const d = [];
    for (const x of [10, 30, 100]) d.push(idOf(await call('create_rectangle', { x, y: 150, width: 10, height: 10 })));
    await call('distribute_objects', { ids: d, distribution: 'horizontal_space', spacing: 5 });
    const xs = [];
    for (const id of d) xs.push(geo(await call('get_object_info', { id })));
    const sorted = xs.sort((a, b) => a.x - b.x);
    near(sorted[1].x - (sorted[0].x + sorted[0].w), 5, 0.01, 'gap 1 is 5 mm'); near(sorted[2].x - (sorted[1].x + sorted[1].w), 5, 0.01, 'gap 2 is 5 mm');
    await call('distribute_objects', { ids: d, distribution: 'vertical_centers' });
    await fails('distribute_objects', { ids: d, distribution: 'left_edges', spacing: 3 }, /spacing only applies/);
    await fails('distribute_objects', { ids: [d[0]], distribution: 'left_edges' }, /at least 2/);
    for (const id of [r, ...copies, ...d]) await call('delete_object', { id });
  });

  await test('get_object_info: path, group, effects, styles', async () => {
    const a = idOf(await call('create_rectangle', { x: 10, y: 10, width: 20, height: 20, fillColor: 'BandGreen' }));
    const b = idOf(await call('create_ellipse', { x: 40, y: 10, width: 20, height: 20 }));
    const grp = idOf(await call('group_objects', { ids: [a, b] }));
    const info = await call('get_object_info', { id: a });
    assert(new RegExp(`group=${grp} \\(members: 2\\)`).test(info), `group membership: ${info}`);
    assert(/objectStyle=/.test(info) && /layer=/.test(info) && /paths=1 points=4/.test(info) && /path 0 \(closed\)/.test(info), info);
    const g = await call('get_object_info', { id: grp });
    assert(/Group/.test(g) && new RegExp(`members: ${a}, ${b}|members: ${b}, ${a}`).test(g), `group lists its members: ${g}`);
    await call('ungroup', { id: grp });
    await call('delete_object', { id: a }); await call('delete_object', { id: b });
    await fails('get_object_info', { id: 424242 }, /No object with id/);
  });

  await test('vector tools never disturb the unit settings of the active document', async () => {
    const units = () => inspect('var v = app.activeDocument.viewPreferences; return [String(v.horizontalMeasurementUnits), String(v.verticalMeasurementUnits), String(v.rulerOrigin)].join(",");');
    inspect('var v = app.activeDocument.viewPreferences; v.horizontalMeasurementUnits = MeasurementUnits.POINTS; v.verticalMeasurementUnits = MeasurementUnits.POINTS; v.rulerOrigin = RulerOrigin.SPREAD_ORIGIN; return "ok";');
    const before = units();
    eq(before, 'POINTS,POINTS,SPREAD_ORIGIN', 'test set-up');
    const id = idOf(await call('create_path', { points: [{ x: 10, y: 10 }, { x: 40, y: 10 }, { x: 40, y: 30 }], closed: true }));
    const g = geo(await call('get_object_info', { id }));
    near(g.x, 10, 0.01, 'the tool still works in millimetres');
    eq(units(), before, 'points and spread origin are restored after the script');
    inspect('var v = app.activeDocument.viewPreferences; v.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS; v.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS; v.rulerOrigin = RulerOrigin.PAGE_ORIGIN; return "ok";');
    await call('delete_object', { id });
    await closeDoc(vec);
    testDoc = flyer;
  });



  console.log('\nSeeing and measuring (Part D)');
  await test('render_preview returns the page as an image block (two pages, different content)', async () => {
    const name = await newDoc({ preset: 'A5', pages: 2, bleed: 3 });
    testDoc = name;
    await call('create_color_swatch', { name: 'PvRed', colorValues: [0, 100, 100, 0] });
    await call('create_color_swatch', { name: 'PvBlue', colorValues: [100, 100, 0, 0] });
    await call('create_rectangle', { x: -3, y: -3, width: 154, height: 216, fillColor: 'PvRed', pageIndex: 0 });
    await call('create_rectangle', { x: -3, y: -3, width: 154, height: 216, fillColor: 'PvBlue', pageIndex: 1 });
    const small = idOf(await call('create_rectangle', { x: 20, y: 20, width: 30, height: 20, fillColor: 'PvBlue', pageIndex: 0 }));
    const shot = async (args) => {
      const content = await callRaw('render_preview', args);
      const img = content.find((c) => c.type === 'image');
      assert(img && img.mimeType === 'image/png' && img.data.length > 100, `an image block is returned: ${JSON.stringify(content.map((c) => c.type))}`);
      const text = content.find((c) => c.type === 'text').text;
      const png = readPng(Buffer.from(img.data, 'base64'));
      return { png, text };
    };
    let { png: p1, text } = await shot({ pageIndex: 0 });
    near(p1.w, 148 / 25.4 * 72, 2, 'width at 72 dpi'); near(p1.h, 210 / 25.4 * 72, 2, 'height at 72 dpi');
    assert(/page 1: \d+ x \d+ px at 72 dpi/.test(text), text);
    assert(isRed(p1.px(5, 5)), `page 1 corner is red: ${p1.px(5, 5)}`);
    const { png: p2 } = await shot({ pageIndex: 1 });
    assert(isBlue(p2.px(p2.w >> 1, p2.h >> 1)), `page 2 is blue: ${p2.px(p2.w >> 1, p2.h >> 1)}`);
    const { png: hi } = await shot({ pageIndex: 0, resolution: 144 });
    near(hi.w, p1.w * 2, 3, '144 dpi is twice as wide');
    const { png: bl } = await shot({ pageIndex: 0, includeBleed: true });
    near(bl.w, (148 + 6) / 25.4 * 72, 2, 'includeBleed adds 2 x 3 mm');
    const { png: obj, text: objText } = await shot({ objectId: small });
    near(obj.w, 30 / 25.4 * 72, 3, 'object render width'); near(obj.h, 20 / 25.4 * 72, 3, 'object render height');
    assert(/object \d+ \(Rectangle\)/.test(objText) && isBlue(obj.px(obj.w >> 1, obj.h >> 1)), objText);
    assert(!fs.readdirSync(path.join(os.tmpdir(), 'indesign-mcp')).some((f) => f.startsWith('preview_')), 'preview temp files are removed');
    await fails('render_preview', { pageIndex: 9 }, /Invalid page index/);
    await fails('render_preview', { resolution: 5 }, /resolution must be/);
    await fails('render_preview', { objectId: 987654 }, /No object with id/);
    const big = await newDoc({ preset: 'A3' });
    await fails('render_preview', { resolution: 600 }, /Lower the resolution/);
    await closeDoc(big);
    await call('activate_document', { name });
    // view_document also shows the page
    const view = await callRaw('view_document');
    assert(view.some((c) => c.type === 'image') && /DOCUMENT VIEW/.test(view[0].text), 'view_document returns text and an image');
    await closeDoc(name);
    testDoc = flyer;
  });

  await test('measure_text: lines, widths, overflow', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    await call('create_paragraph_style', { name: 'Meas', fontFamily: 'Helvetica Neue', fontStyle: 'Regular', fontSize: 12, leading: 15, hyphenation: false });
    const wrap = idOf(await call('create_text_frame', { content: 'Hello world this text wraps across several lines in the small frame', paragraphStyle: 'Meas', x: 10, y: 10, width: 50, height: 40 }));
    let out = await call('measure_text', { frameId: wrap });
    const lines = [...out.matchAll(/\[(\d+)\] width=([\d.]+) mm baseline=[\d.]+ mm: "(.*)"/g)].map((m) => ({ w: Number(m[2]), t: m[3] }));
    assert(lines.length >= 3, `the text wraps into several lines: ${out}`);
    for (const l of lines) assert(l.w > 5 && l.w <= 50.01, `each line fits the 50 mm frame (${l.w} mm: ${l.t})`);
    assert(/Hello world/.test(lines[0].t), `line texts: ${lines[0].t}`);
    assert(new RegExp(`longest line=${Math.max(...lines.map((l) => l.w))} mm`).test(out), 'longest line reported');
    assert(/frame width=50 mm/.test(out) && /overflows=false/.test(out), out);
    eq(lines.map((l) => l.t).join(' ').replace(/\s+/g, ' ').trim(), 'Hello world this text wraps across several lines in the small frame', 'all lines together are the whole text');
    // width scales with the font size
    const one = idOf(await call('create_text_frame', { content: 'Measure me', paragraphStyle: 'Meas', x: 10, y: 60, width: 120, height: 20 }));
    const w12 = Number((await call('measure_text', { frameId: one })).match(/\[0\] width=([\d.]+)/)[1]);
    await call('edit_text_frame', { frameId: one, fontSize: 24 });
    const w24 = Number((await call('measure_text', { frameId: one })).match(/\[0\] width=([\d.]+)/)[1]);
    near(w24 / w12, 2, 0.05, 'double font size gives double width');
    // forced breaks
    const brk = idOf(await call('create_text_frame', { content: 'AAA<br>BBBBBB', paragraphStyle: 'Meas', x: 10, y: 90, width: 100, height: 30 }));
    out = await call('measure_text', { frameId: brk });
    const bw = [...out.matchAll(/\[\d\] width=([\d.]+)/g)].map((m) => Number(m[1]));
    eq(bw.length, 2, 'a forced break makes two lines'); assert(bw[1] > bw[0] * 1.5, `BBBBBB is wider than AAA: ${bw}`);
    // overflow
    const tiny = idOf(await call('create_text_frame', { content: 'This is far too much text for such a tiny frame. '.repeat(3), paragraphStyle: 'Meas', x: 10, y: 130, width: 30, height: 8 }));
    out = await call('measure_text', { frameId: tiny });
    assert(/overflows=true/.test(out), out);
    await fails('measure_text', { frameId: 999999 }, /No object with id/);
    await fails('measure_text', { frameIndex: 99 }, /Invalid index/);
    await closeDoc(name);
    testDoc = flyer;
  });

  console.log('\nVector graphics placement (Part B)');
  const SVG_MARKUP = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50" viewBox="0 0 100 50"><rect width="100" height="50" fill="#2e7d32"/><circle cx="30" cy="25" r="15" fill="#ffffff"/></svg>';
  await test('place_graphic: SVG (linked and embedded) and a two-page PDF', async () => {
    const name = await newDoc({ preset: 'A5', pages: 2 });
    testDoc = name;
    const svgFile = path.join(TMP, 'logo.svg');
    fs.writeFileSync(svgFile, SVG_MARKUP);
    let out = await call('place_graphic', { filePath: svgFile, x: 10, y: 10, width: 60 });
    const svgFrame = idOf(out);
    assert(/w=60 h=30/.test(out) && /type=SVG/.test(out) && /linked=true \(NORMAL\)/.test(out), `aspect ratio kept, linked: ${out}`);
    out = await call('place_graphic', { filePath: svgFile, x: 10, y: 50, width: 40, height: 40, fitOption: 'FILL_PROPORTIONALLY' });
    assert(/w=40 h=40/.test(out) && /linked=true/.test(out), out);
    out = await call('place_graphic', { filePath: svgFile, x: 10, y: 100, height: 20, embed: true });
    assert(/w=40 h=20/.test(out) && /linked=false \(embedded\)/.test(out), `embedded: ${out}`);
    assert(/LINK_EMBEDDED/.test(await call('get_object_info', { id: idOf(out) })), 'get_object_info reports the embedded state');
    out = await call('place_graphic', { filePath: svgFile, x: 80, y: 10 });
    assert(/x=80 y=10 w=35\.278/.test(out) && /linked=true/.test(out), `natural size (100 px = 35.278 mm) at the requested position: ${out}`);
    {   // the content must sit inside the frame (a moved frame with stale content shows blank or shifted art)
      const c = await callRaw('render_preview', { objectId: idOf(out), resolution: 144 });
      const png = readPng(Buffer.from(c.find((q) => q.type === 'image').data, 'base64'));
      const green = ([r, g, b]) => r < 90 && g > 100 && b < 90;
      for (const [x, y] of [[2, 2], [png.w - 3, 2], [2, png.h - 3], [png.w - 3, png.h - 3]]) assert(green(png.px(x, y)), `natural-size SVG fills its frame, corner (${x},${y}) is ${png.px(x, y)}`);
    }
    // a two-page PDF with different content per page; choose the page and check by rendering the frame
    await call('create_color_swatch', { name: 'PdfRed', colorValues: [0, 100, 100, 0] });
    await call('create_color_swatch', { name: 'PdfBlue', colorValues: [100, 100, 0, 0] });
    await call('create_rectangle', { x: 0, y: 0, width: 148, height: 210, fillColor: 'PdfRed', pageIndex: 0 });
    await call('create_rectangle', { x: 0, y: 0, width: 148, height: 210, fillColor: 'PdfBlue', pageIndex: 1 });
    const pdf = path.join(TMP, 'two-pages.pdf');
    await call('export_pdf', { filePath: pdf, confirmDestructive: true });
    const clean = async () => { for (const p of [0, 1]) for (const it of (await call('list_page_items', { pageIndex: p, type: 'Rectangle' })).matchAll(/Rectangle id=(\d+)[^\n]*image=/g)) await call('delete_object', { id: Number(it[1]) }); };
    const centre = async (id) => { const c = await callRaw('render_preview', { objectId: id }); const png = readPng(Buffer.from(c.find((x) => x.type === 'image').data, 'base64')); return png.px(png.w >> 1, png.h >> 1); };
    await clean();
    const target = await newDoc({ preset: 'A5' });
    testDoc = target;
    const p1 = await call('place_graphic', { filePath: pdf, x: 10, y: 10, width: 60, pdfPage: 1 });
    assert(/pdfPage=1/.test(p1) && /linked=true/.test(p1) && /type=(Adobe )?PDF/i.test(p1), p1);
    const p2 = await call('place_graphic', { filePath: pdf, x: 80, y: 10, width: 60, pdfPage: 2, crop: 'trim' });
    assert(/pdfPage=2/.test(p2), p2);
    assert(isRed(await centre(idOf(p1))), 'PDF page 1 is red');
    assert(isBlue(await centre(idOf(p2))), 'PDF page 2 is blue (the page number is honoured)');
    const emb = await call('place_graphic', { filePath: pdf, x: 10, y: 100, width: 40, pdfPage: 2, embed: true });
    assert(/linked=false \(embedded\)/.test(emb), emb);
    await closeDoc(target);
    testDoc = name;
    await fails('place_graphic', { filePath: path.join(TMP, 'notes.txt') }, /Unsupported file type \.txt/);
    await fails('place_graphic', { filePath: path.join(TMP, 'missing.svg') }, /File not found/);
    await fails('place_graphic', { filePath: '/etc/hosts.svg' }, /Access denied/);
    await fails('place_graphic', { filePath: svgFile, crop: 'sideways' }, /Invalid parameter "crop"/);
    await fails('place_graphic', { filePath: svgFile, fitOption: 'NOPE' }, /Invalid parameter "fitOption"/);
    fs.writeFileSync(path.join(TMP, 'broken.svg'), 'this is not svg at all');
    const before = numAfter(await call('get_document_info'), 'Rectangles \\(incl\\. image frames\\)');
    await fails('place_graphic', { filePath: path.join(TMP, 'broken.svg') }, /./);
    eq(numAfter(await call('get_document_info'), 'Rectangles \\(incl\\. image frames\\)'), before, 'a failed placement leaves no empty frame behind');
    assert(svgFrame > 0, 'frame id returned');
    await closeDoc(name);
    testDoc = flyer;
  });

  for (const [envName, ext, label] of [['INDESIGN_TEST_AI', 'ai', 'Illustrator'], ['INDESIGN_TEST_EPS', 'eps', 'EPS']]) {
    const file = process.env[envName];
    if (!file || !fs.existsSync(file)) {
      console.log(`  skip  place_graphic: real ${label} file (set ${envName}=/path/to/file.${ext} to test it)`);
      continue;
    }
    await test(`place_graphic: a real .${ext} file (${path.basename(file)}): linked, embedded, crop and frame geometry`, async () => {
      const name = await newDoc({ preset: 'A4' });
      testDoc = name;
      const before = fs.statSync(file).mtimeMs;
      let out = await call('place_graphic', { filePath: file, x: 10, y: 10, width: 60 });
      assert(/w=60 /.test(out) && /linked=true \(NORMAL\)/.test(out) && new RegExp(`image=${path.basename(file).replace(/[.]/g, '\\.')}`).test(out), out);
      const h = numAfter(out, 'h');
      assert(h > 5, `the height follows the aspect ratio: ${out}`);
      out = await call('place_graphic', { filePath: file, x: 10, y: 120, width: 60, embed: true });
      assert(/linked=false \(embedded\)/.test(out), out);
      const c = await callRaw('render_preview', { objectId: idOf(out), resolution: 72 });
      const shot = readPng(Buffer.from(c.find((q) => q.type === 'image').data, 'base64'));
      assert(inkBox(shot), 'the placed artwork renders (it is not blank)');
      if (ext === 'ai') {
        const crops = {};
        for (const crop of ['pdf', 'art']) crops[crop] = numAfter(await call('place_graphic', { filePath: file, x: 80, y: 10, width: 60, crop, embed: true }), 'h');
        assert(crops.pdf > 0 && crops.art > 0, `crop modes give a frame: ${JSON.stringify(crops)}`);
        await fails('place_graphic', { filePath: file, pdfPage: 2 }, /has no page 2/);
      }
      eq(fs.statSync(file).mtimeMs, before, 'the source file was only read, never modified');
      await closeDoc(name);
      testDoc = flyer;
    });
  }

  if (process.env.INDESIGN_TEST_JPG && process.env.INDESIGN_TEST_AI && process.env.INDESIGN_TEST_EPS) {
    await test('sample files: JPEG and EPS are the same photo, the AI sample is a landscape illustration', async () => {
      const name = await newDoc({ preset: 'A4' });
      testDoc = name;
      const jpg = await call('place_image', { imagePath: process.env.INDESIGN_TEST_JPG, x: 10, y: 10, width: 60 });
      const ai = await call('place_graphic', { filePath: process.env.INDESIGN_TEST_AI, x: 80, y: 10, width: 60, embed: true });
      const eps = await call('place_graphic', { filePath: process.env.INDESIGN_TEST_EPS, x: 150, y: 10, width: 60, embed: true });
      near(numAfter(jpg, 'h'), 60 * 2400 / 1351, 0.6, 'the JPEG (1351 x 2400 px) keeps its aspect ratio');
      near(numAfter(eps, 'h'), 60 * 2400 / 1351, 0.6, 'the EPS (440 x 782 pt) has the same aspect ratio');
      near(numAfter(ai, 'h'), 40, 0.6, 'the AI illustration is 3:2 landscape');
      assert(/linked=false \(embedded\)/.test(ai) && /linked=false \(embedded\)/.test(eps), 'ai and eps are embedded');
      const mean = async (o) => {
        const c = await callRaw('render_preview', { objectId: idOf(o), resolution: 36 });
        const png = readPng(Buffer.from(c.find((q) => q.type === 'image').data, 'base64'));
        let r = 0, g = 0, b = 0, n = 0;
        for (let y = 0; y < png.h; y += 3) for (let x = 0; x < png.w; x += 3) { const p = png.px(x, y); r += p[0]; g += p[1]; b += p[2]; n++; }
        return [r / n, g / n, b / n];
      };
      const [mj, ma, me] = [await mean(jpg), await mean(ai), await mean(eps)];
      for (let k = 0; k < 3; k++) assert(Math.abs(me[k] - mj[k]) < 30, `the EPS renders like the JPEG: mean colour ${me.map(Math.round)} vs ${mj.map(Math.round)}`);
      assert(ma.some((v) => v < 235) && Math.abs(ma[0] - mj[0]) + Math.abs(ma[2] - mj[2]) > 20, `the AI shows its own, different picture: ${ma.map(Math.round)}`);
      await closeDoc(name);
      testDoc = flyer;
    });
  }

  await test('create_graphic_from_svg: linked temp file, embed, saveTo, and rejected markup', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    const svgDir = path.join(os.tmpdir(), 'indesign-mcp', 'svg');
    let out = await call('create_graphic_from_svg', { svg: SVG_MARKUP, x: 10, y: 10, width: 60 });
    assert(/w=60 h=30/.test(out) && /linked=true/.test(out) && /not editable as paths/.test(out), out);
    const linkedFile = out.match(/source file: (\S+\.svg)/)[1];
    assert(linkedFile.startsWith(svgDir) && fs.existsSync(linkedFile), `linked temp file exists under $TMPDIR/indesign-mcp/svg: ${linkedFile}`);
    const before = fs.readdirSync(svgDir).length;
    out = await call('create_graphic_from_svg', { svg: SVG_MARKUP, x: 10, y: 60, width: 40, embed: true });
    assert(/linked=false \(embedded\)/.test(out) && /deleted, graphic is embedded/.test(out), out);
    eq(fs.readdirSync(svgDir).length, before, 'the temp file of an embedded graphic is deleted');
    const keep = path.join(TMP, 'kept', 'logo.svg');
    out = await call('create_graphic_from_svg', { svg: SVG_MARKUP, x: 10, y: 100, width: 30, saveTo: keep, embed: true });
    assert(fs.existsSync(keep) && fs.readFileSync(keep, 'utf8') === SVG_MARKUP, 'saveTo keeps the SVG file');
    // markup that could make InDesign read files or fetch URLs is refused, and nothing is placed
    const count = async () => Number((await call('get_document_info')).match(/Rectangles \(incl\. image frames\): (\d+)/)[1]);
    const n0 = await count();
    for (const [svg, re] of [
      ['<svg xmlns="http://www.w3.org/2000/svg"><image href="file:///etc/passwd"/></svg>', /external resource/],
      ['<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><image xlink:href="http://example.com/x.png"/></svg>', /external resource/],
      ['<!DOCTYPE svg [<!ENTITY x SYSTEM "file:///etc/passwd">]><svg xmlns="http://www.w3.org/2000/svg"/>', /DOCTYPE/],
      ['<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>', /<script>/],
      ['<svg xmlns="http://www.w3.org/2000/svg" onload="x()"/>', /event handler/],
      ['<svg xmlns="http://www.w3.org/2000/svg"><rect style="fill:url(http://evil/x)"/></svg>', /external url/],
      ['<svg xmlns="http://www.w3.org/2000/svg"><style>@import "http://evil/x.css";</style></svg>', /@import/],
      ['not svg', /must be SVG markup/],
    ]) await fails('create_graphic_from_svg', { svg }, re);
    eq(await count(), n0, 'rejected SVGs place nothing');
    await fails('create_graphic_from_svg', { svg: SVG_MARKUP, saveTo: path.join(TMP, 'x.png') }, /must end with \.svg/);
    await fails('create_graphic_from_svg', { svg: SVG_MARKUP, saveTo: '/etc/logo.svg' }, /Access denied/);
    // a placeholder that InDesign cannot read: no temp file is left behind
    const files0 = fs.readdirSync(svgDir).length;
    await fails('create_graphic_from_svg', { svg: '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="5"' });
    eq(fs.readdirSync(svgDir).length, files0, 'a failed placement removes its temp file');
    await closeDoc(name);
    testDoc = flyer;
  });

  await test('create_cmyk_pdf_shape: exact CMYK vector shapes', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    const pdfDir = path.join(os.tmpdir(), 'indesign-mcp', 'pdf');
    const pdfBefore = fs.existsSync(pdfDir) ? fs.readdirSync(pdfDir).length : 0;
    const out = await call('create_cmyk_pdf_shape', {
      shapes: [
        { d: 'M0 0 H100 V100 H0 Z', fill: [100, 0, 100, 0] },
        { d: 'M50 20 A30 30 0 1 1 49.9 20 Z', fill: [0, 100, 100, 0], stroke: [0, 0, 0, 100], strokeWidth: 3 },
      ],
      viewBox: '0 0 100 100', x: 20, y: 20, width: 50,
    });
    const id = idOf(out);
    assert(/w=50 h=50/.test(out) && /linked=false \(embedded\)/.test(out) && /type=(Adobe )?PDF/i.test(out), out);
    // the exact CMYK numbers survive: export as PDF/X-4 (no colour conversion) and read the page content stream
    const x4 = path.join(outDir, 'cmyk-shape-x4.pdf');
    await call('export_pdf', { filePath: x4, preset: 'PDFX4', confirmDestructive: true });
    const raw = fs.readFileSync(x4);
    const streams = [];
    const latin = raw.toString('latin1');
    for (const m of latin.matchAll(/stream\r?\n/g)) {
      if (latin.slice(m.index - 3, m.index) === 'end') continue;
      const start = m.index + m[0].length, end = latin.indexOf('endstream', start);
      try { streams.push(zlib.inflateSync(raw.subarray(start, end)).toString('latin1')); } catch { /* not a flate stream */ }
    }
    const content = streams.join('\n');
    assert(/(^|\s)1 0 1 0 k\b/.test(content), 'the square keeps its exact CMYK fill 100/0/100/0 (as "1 0 1 0 k")');
    assert(/(^|\s)0 1 1 0 k\b/.test(content) && /(^|\s)0 0 0 1 K\b/.test(content), 'the circle keeps 0/100/100/0 and its stroke 0/0/0/100');
    assert(!/DeviceRGB/.test(raw.toString('latin1')), 'no RGB colour space in the exported file');
    const c = await callRaw('render_preview', { objectId: id, resolution: 144 });
    const png = readPng(Buffer.from(c.find((x) => x.type === 'image').data, 'base64'));
    const corner = png.px(Math.floor(png.w * 0.05), Math.floor(png.h * 0.05));
    assert(corner[1] > 120 && corner[0] < 70 && corner[2] < 120, `the square is green: ${corner}`);
    assert(isRed(png.px(png.w >> 1, png.h >> 1)), `the circle is red: ${png.px(png.w >> 1, png.h >> 1)}`);
    await fails('create_cmyk_pdf_shape', { shapes: [] }, /at least one shape/);
    await fails('create_cmyk_pdf_shape', { shapes: [{ d: 'M0 0 L10 10' }] }, /needs a fill and\/or a stroke/);
    await fails('create_cmyk_pdf_shape', { shapes: [{ d: 'M0 0 L10 10 L0 10 Z', fill: [1, 2, 3] }] }, /\[C,M,Y,K\]/);
    await fails('create_cmyk_pdf_shape', { shapes: [{ d: 'M0 0 L10 10 L0 10 Z', fill: [0, 0, 0, 200] }] }, /0-100/);
    await fails('create_cmyk_pdf_shape', { shapes: [{ d: 'M0 0 X', fill: [0, 0, 0, 100] }] }, /Invalid SVG path data/);
    eq(fs.readdirSync(pdfDir).length, pdfBefore, 'embedded shapes and failed calls leave no temp files behind');
    await closeDoc(name);
    testDoc = flyer;
  });


  console.log('\nSwatches, fonts, styles, PDF (Part C)');
  await test('C4: swatches can be updated, renamed, deleted and listed with values and usage', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    await call('create_color_swatch', { name: 'Brand2', colorValues: [100, 50, 0, 10] });
    await fails('create_color_swatch', { name: 'Brand2', colorValues: [0, 0, 0, 0] }, /already exists.*update/);
    let out = await call('create_color_swatch', { name: 'Brand2', colorValues: [90, 40, 0, 5], update: true });
    assert(/updated: CMYK 90,40,0,5/.test(out), `create with update: true changes the values: ${out}`);
    const rect = idOf(await call('create_rectangle', { x: 10, y: 10, width: 30, height: 20, fillColor: 'Brand2' }));
    const other = idOf(await call('create_rectangle', { x: 50, y: 10, width: 30, height: 20, fillColor: 'Brand2' }));
    out = await call('update_color_swatch', { name: 'Brand2', colorValues: [10, 20, 30, 40] });
    assert(/updated: CMYK process 10,20,30,40/.test(out) && /Used by 2 fill\(s\)/.test(out), out);
    assert(/fill: Brand2 \[CMYK 10,20,30,40\]/.test(await call('get_object_info', { id: rect })), 'objects follow the changed swatch');
    out = await call('update_color_swatch', { name: 'Brand2', hex: '#FF6600' });
    assert(/CMYK process 0,70\.\d+,9\d\.\d+,0/.test(out) && /converted to CMYK with the document profile/.test(out), out);
    out = await call('update_color_swatch', { name: 'Brand2', spotColor: true, newName: 'Brand2 Spot' });
    assert(/Swatch Brand2 Spot updated: CMYK spot/.test(out), out);
    assert(/fill: Brand2 Spot/.test(await call('get_object_info', { id: other })), 'objects follow the rename');
    await fails('update_color_swatch', { name: 'Nope', colorValues: [0, 0, 0, 0] }, /Swatch not found/);
    await fails('update_color_swatch', { name: 'Brand2 Spot' }, /Nothing to change/);
    await fails('update_color_swatch', { name: 'Brand2 Spot', colorValues: [1, 2] }, /colorValues must be/);
    await call('create_color_swatch', { name: 'Other', colorValues: [0, 0, 0, 50] });
    await fails('update_color_swatch', { name: 'Brand2 Spot', newName: 'Other' }, /already exists/);
    // list with values and usage
    await call('create_gradient_swatch', { name: 'GradC4', stops: [{ swatch: 'Brand2 Spot', position: 0 }, { swatch: 'Other', position: 100 }] });
    const list = await call('list_color_swatches');
    assert(/• Brand2 Spot {2}\(Color, CMYK spot \[0, 70\.\d, 97\.\d, 0\]\)/.test(list), `values are listed (the colour-managed CMYK of #FF6600): ${list}`);
    assert(new RegExp(`Brand2 Spot[^\\n]*used: 2 fill\\(s\\), 0 stroke\\(s\\), 0 style\\(s\\), ids (${rect},${other}|${other},${rect})`).test(list), `usage with ids: ${list}`);
    assert(/Other {2}\(Color, CMYK process \[0, 0, 0, 50\]\) {2}used: not used/.test(list), 'an unused swatch says so');
    assert(/GradC4 {2}\(Gradient, linear: 0% Brand2 Spot \| 100% Other\)/.test(list), `gradients list their stops: ${list}`);
    assert(/\(built-in/.test(list), 'built-in swatches are marked');
    assert(!/used:/.test(await call('list_color_swatches', { includeUsage: false })), 'includeUsage false skips the scan');
    // delete
    await fails('delete_color_swatch', { name: 'Brand2 Spot' }, /used by 2 fill\(s\).*replaceWith/);
    out = await call('delete_color_swatch', { name: 'Brand2 Spot', replaceWith: 'Other' });
    assert(/deleted \(2 use\(s\) replaced\)/.test(out), out);
    assert(/fill: Other/.test(await call('get_object_info', { id: rect })), 'the replacement swatch took its place');
    await fails('delete_color_swatch', { name: 'Brand2 Spot' }, /Swatch not found/);
    await fails('delete_color_swatch', { name: 'Other', replaceWith: 'Other' }, /different swatch/);
    await call('set_object_fill', { ids: [rect, other], swatch: 'none' });
    out = await call('delete_color_swatch', { name: 'Other' });
    assert(/deleted\./.test(out), `an unused swatch needs no replacement: ${out}`);
    await closeDoc(name);
    testDoc = flyer;
  });

  await test('C5: paragraph style typography, list_fonts and font suggestions', async () => {
    const name = await newDoc({ preset: 'A5' });
    testDoc = name;
    let out = await call('create_paragraph_style', {
      name: 'Typo', fontFamily: 'Helvetica Neue', fontStyle: 'Bold', fontSize: 11, hyphenation: false, composer: 'single-line', capitalization: 'all_caps',
      ligatures: false, openType: { figureStyle: 'proportional_oldstyle', discretionaryLigatures: true, slashedZero: true },
      hyphenateWordsLongerThan: 8, hyphenateAfterFirst: 3, hyphenateBeforeLast: 4, hyphenateLadderLimit: 2, hyphenateCapitalizedWords: false,
      keepWithNext: 1, keepWithPrevious: true, keepFirstLines: 2, keepLastLines: 2, tracking: 25, kerning: 'Optical', leftIndent: 4, firstLineIndent: 2,
    });
    assert(/hyphenation=false/.test(out) && /composer=Adobe Single-line Composer/.test(out) && /capitalization=ALL_CAPS/.test(out) && /ligatures=false/.test(out), out);
    eq(inspect(`var p = app.activeDocument.paragraphStyles.itemByName("Typo"); return [p.hyphenateWordsLongerThan, p.hyphenateAfterFirst, p.hyphenateBeforeLast, p.hyphenateLadderLimit, p.hyphenateCapitalizedWords, String(p.otfFigureStyle), p.otfDiscretionaryLigature, p.otfSlashedZero, p.keepWithPrevious, p.keepFirstLines, p.keepLastLines, p.tracking, String(p.kerningMethod)].join(",");`),
      '8,3,4,2,false,PROPORTIONAL_OLDSTYLE,true,true,true,2,2,25,Optical', 'every property really landed on the style');
    out = await call('modify_paragraph_style', { styleName: 'Typo', composer: 'paragraph', capitalization: 'normal', hyphenation: true, ligatures: true });
    assert(/composer=Adobe Paragraph Composer/.test(out) && /capitalization=NORMAL/.test(out) && /hyphenation=true/.test(out), out);
    await fails('create_paragraph_style', { name: 'T2', composer: 'magic' }, /Invalid parameter "composer"/);
    await fails('create_paragraph_style', { name: 'T2', capitalization: 'shouting' }, /Invalid parameter "capitalization"/);
    await fails('create_paragraph_style', { name: 'T2', openType: { sparkles: true } }, /Unknown openType feature: sparkles/);
    await fails('create_paragraph_style', { name: 'T2', openType: { swash: 'yes' } }, /must be true or false/);
    await fails('create_paragraph_style', { name: 'T2', openType: { figureStyle: 'roman' } }, /Invalid openType\.figureStyle/);
    assert(!/T2/.test(await call('list_styles', { styleType: 'paragraph' })), 'failed creates leave no style');

    // fonts
    const all = await call('list_fonts', { limit: 5 });
    assert(/=== FONTS: \d+ of \d+ installed families ===/.test(all) && /\.\.\. \d+ more/.test(all), all);
    const hel = await call('list_fonts', { search: 'helvetica neue' });
    assert(/Helvetica Neue: .*\bBold\b/.test(hel), `families with their styles: ${hel}`);
    const one = await call('list_fonts', { family: 'Helvetica Neue' });
    assert(/^Helvetica Neue: /m.test(one) && (one.match(/^.+: /gm) || []).length >= 2, one);
    const missing = await fails('list_fonts', { family: 'Helvetica Neu' }, /Font not installed: Helvetica Neu/);
    assert(/Did you mean: Helvetica Neue \[/.test(missing), `suggestions in the error: ${missing}`);
    const noStyle = await fails('create_paragraph_style', { name: 'T3', fontFamily: 'Helvetica Neue', fontStyle: 'Boldx' }, /has no style Boldx/);
    assert(/Available styles: .*Bold/.test(noStyle), `the available styles are listed: ${noStyle}`);
    const near = await fails('create_paragraph_style', { name: 'T3', fontFamily: 'Helvetica Neu', fontStyle: 'Bold' }, /Font not installed: Helvetica Neu \(Bold\)/);
    assert(/Did you mean: Helvetica Neue/.test(near), near);
    await fails('create_text_frame', { content: 'x', fontFamily: 'Zzzz Unknown' }, /Font not installed: Zzzz Unknown.*Use list_fonts/);
    // "Family Style" written as one family name is understood
    out = await call('create_paragraph_style', { name: 'T4', fontFamily: 'Helvetica Neue Bold' });
    assert(/font=Helvetica Neue Bold/.test(out) && /fontStyle=Bold/.test(out), `"Helvetica Neue Bold" works as family + style: ${out}`);
    await fails('list_fonts', { limit: 0 }, /limit must be/);
    await closeDoc(name);
    testDoc = flyer;
  });

  await test('C7: export_pdf reports the effective preset; PDF/X presets, bleed, slug and page range', async () => {
    const name = await newDoc({ preset: 'A5', pages: 3, bleed: 3, slug: 5 });
    testDoc = name;
    await call('create_text_frame', { content: 'PDF/X test' });
    const dir = path.join(outDir, 'c7');
    const text = (f) => fs.readFileSync(f).toString('latin1');
    let out = await call('export_pdf', { filePath: path.join(dir, 'web.pdf'), preset: 'Web', confirmDestructive: true });
    assert(/requested preset Web, effective preset \[Smallest File Size\]/.test(out), `the alias and the effective preset are both reported: ${out}`);
    out = await call('export_pdf', { filePath: path.join(dir, 'hq.pdf'), preset: 'Print', confirmDestructive: true });
    assert(/requested preset Print, effective preset \[High Quality Print\]/.test(out), out);
    for (const [preset, effective, marker] of [['PDFX1a', '[PDF/X-1a:2001]', /PDF\/X-1a:2001/], ['PDFX3', '[PDF/X-3:2002]', /PDF\/X-3:2002/], ['PDFX4', '[PDF/X-4:2008]', /PDF\/X-4/]]) {
      const file = path.join(dir, `${preset}.pdf`);
      out = await call('export_pdf', { filePath: file, preset, confirmDestructive: true });
      assert(out.includes(`effective preset ${effective}`), `${preset}: ${out}`);
      assert(marker.test(text(file)), `${preset}: the file declares ${marker}`);
      eq(pdfBoxes(file).pages, 3, `${preset}: page count`);
    }
    // page ranges in several notations
    for (const [range, pages] of [['1', 1], ['2-3', 2], ['1,3', 2], ['all', 3]]) {
      const file = path.join(dir, `range-${range.replace(/\W/g, '_')}.pdf`);
      await call('export_pdf', { filePath: file, pageRange: range, confirmDestructive: true });
      eq(pdfBoxes(file).pages, pages, `pageRange ${range}`);
    }
    // bleed and slug enlarge the media box, the trim box stays
    const plain = pdfBoxes(path.join(dir, 'hq.pdf'));
    const bleed = path.join(dir, 'bleed.pdf'), slug = path.join(dir, 'slug.pdf');
    await call('export_pdf', { filePath: bleed, includeBleed: true, confirmDestructive: true });
    await call('export_pdf', { filePath: slug, includeBleed: true, includeSlug: true, confirmDestructive: true });
    near(mm(boxSize(plain.media)[0]), 148, 0.5, 'plain media width = trim');
    near(mm(boxSize(pdfBoxes(bleed).media)[0]), 154, 0.5, 'bleed adds 2 x 3 mm');
    assert(mm(boxSize(pdfBoxes(slug).media)[0]) > mm(boxSize(pdfBoxes(bleed).media)[0]) + 2, 'slug adds more');
    near(mm(boxSize(pdfBoxes(slug).trim)[0]), 148, 0.5, 'trim box unchanged');
    await closeDoc(name);
    testDoc = flyer;
  });

  console.log('\nPage numbering and saving');
  const isGreen = ([r, g, b]) => g > 100 && r < 110 && b < 120;
  const pageNames = () => inspect('var n = []; for (var i = 0; i < app.activeDocument.pages.length; i++) n.push(app.activeDocument.pages[i].name); return n.join(",");');

  // A 3-page document: red / green / blue full-page rectangles and a "PAGE n" text on each page
  async function colourDoc(docArgs, numbering) {
    const name = await newDoc({ preset: 'A5', pages: 3, ...docArgs });
    testDoc = name;
    if (numbering) inspect(numbering);
    await call('create_color_swatch', { name: 'NpRed', colorValues: [0, 100, 100, 0] });
    await call('create_color_swatch', { name: 'NpGreen', colorValues: [85, 10, 100, 20] });
    await call('create_color_swatch', { name: 'NpBlue', colorValues: [100, 90, 0, 0] });
    for (const [i, sw] of ['NpRed', 'NpGreen', 'NpBlue'].entries()) {
      await call('create_rectangle', { x: 0, y: 0, width: 148, height: 210, fillColor: sw, pageIndex: i });
      await call('create_text_frame', { content: `PAGE ${i + 1}`, pageIndex: i, x: 10, y: 10, width: 100, height: 20, fontSize: 24, textColor: 'Paper' });
    }
    return name;
  }
  const colourOf = (file) => { const p = readPng(file); const px = p.px(p.w >> 1, p.h >> 1); return isRed(px) ? 1 : isGreen(px) ? 2 : isBlue(px) ? 3 : 0; };

  for (const [label, docArgs, numbering, expectNames] of [
    ['plain numbering', {}, null, '1,2,3'],
    ['numbering starts at 5', {}, 'var s = app.activeDocument.sections[0]; s.continueNumbering = false; s.pageNumberStart = 5; return "ok";', '5,6,7'],
    ['sections with repeated page names (1,1,2)', {}, 'var d = app.activeDocument; d.sections[0].continueNumbering = false; d.sections[0].pageNumberStart = 1; d.sections.add(d.pages[1], { continueNumbering: false, pageNumberStart: 1 }); return "ok";', '1,1,2'],
    ['facing pages, numbering starts at 5', { facingPages: true }, 'var s = app.activeDocument.sections[0]; s.continueNumbering = false; s.pageNumberStart = 5; return "ok";', '5,6,7'],
  ]) {
    await test(`export_images: every page file shows its own page (${label})`, async () => {
      const name = await colourDoc(docArgs, numbering);
      eq(pageNames(), expectNames, `the page names of this document are ${expectNames} (so names cannot identify positions)`);
      const base = name.replace(/\.indd$/i, '');
      let n = 0;
      for (const [range, pages] of [['all', [1, 2, 3]], ['1', [1]], ['2', [2]], ['3', [3]], ['2-3', [2, 3]], ['1,3', [1, 3]]]) {
        const dir = path.join(outDir, `num-${label.replace(/\W+/g, '_')}-${range.replace(/\W/g, '_')}`);
        const out = await call('export_images', { folderPath: dir, resolution: 18, pageRange: range, confirmDestructive: true });
        assert(out.includes(`(pages: ${pages.join(', ')})`), `the result lists the pages: ${out}`);
        eq(fs.readdirSync(dir).sort().join(','), pages.map((p) => `${base}_page${p}.png`).join(','), `${range}: one file per requested page, named by position`);
        for (const p of pages) eq(colourOf(path.join(dir, `${base}_page${p}.png`)), p, `${range}: the file ${base}_page${p}.png shows page ${p} (1 = red, 2 = green, 3 = blue)`);
        n++;
      }
      const jdir = path.join(outDir, `num-${label.replace(/\W+/g, '_')}-jpeg`);
      await call('export_images', { folderPath: jdir, format: 'JPEG', resolution: 18, confirmDestructive: true });
      eq(fs.readdirSync(jdir).length, 3, 'JPEG: three files');
      // the preview and the PDF use the same absolute page mapping
      for (const idx of [0, 1, 2]) {
        const c = await callRaw('render_preview', { pageIndex: idx, resolution: 18 });
        eq(colourOf(Buffer.from(c.find((q) => q.type === 'image').data, 'base64')), idx + 1, `render_preview of page index ${idx} shows page ${idx + 1}`);
      }
      for (const [range, count] of [['2', 1], ['2-3', 2], ['1,3', 2], ['all', 3]]) {
        const pdf = path.join(outDir, `num-${label.replace(/\W+/g, '_')}-${range.replace(/\W/g, '_')}.pdf`);
        await call('export_pdf', { filePath: pdf, pageRange: range, confirmDestructive: true });
        eq(pdfBoxes(pdf).pages, count, `export_pdf pageRange ${range} gives ${count} page(s)`);
      }
      await fails('export_images', { folderPath: path.join(outDir, 'num-bad'), pageRange: '4', confirmDestructive: true }, /does not exist/);
      await closeDoc(name);
      testDoc = flyer;
    });
  }

  await test('a failing page mapping raises an error instead of an InDesign dialog', async () => {
    // InDesign is told never to open blocking dialogs, so a bad page range can not freeze all later tool calls
    const name = await newDoc({ preset: 'A5', pages: 2 });
    testDoc = name;
    const dlg = inspect('var before = app.scriptPreferences.userInteractionLevel; return String(before);');
    eq(dlg, 'INTERACT_WITH_ALL', 'the interaction level is back to normal after every tool call');
    await fails('export_images', { folderPath: path.join(outDir, 'nodialog'), pageRange: '9', confirmDestructive: true }, /does not exist/);
    assert(/OPEN DOCUMENTS/.test(await call('list_open_documents')), 'InDesign still answers');
    await closeDoc(name);
    testDoc = flyer;
  });

  await test('save_document: Save As, close, reopen, edit, plain save persists; same-path and never-saved cases', async () => {
    const dir = path.join(TMP, 'save-cases');
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, 'reopen.indd');
    const text = async (id) => (await call('get_text_content', { frameId: id, normalizeSpaces: false })).match(/TEXT:\n(.*)/s)[1].trim();
    const scratch = await newDoc({ preset: 'A5' });
    testDoc = scratch;
    const frame = idOf(await call('create_text_frame', { content: 'version 1', x: 10, y: 10, width: 100, height: 20 }));

    // a document that was never saved
    let e = await fails('save_document', {}, /has never been saved: pass filePath/);
    assert(!/Security|already open/.test(e), e);
    assert(/path=unsaved/.test(await call('list_open_documents')), 'listed as unsaved');

    // Save As to a new name (no confirmation needed for a new file)
    let out = await call('save_document', { filePath: file });
    assert(/Save As: the document .* is now named reopen\.indd/.test(out), out);
    createdDocs[createdDocs.indexOf(scratch)] = 'reopen.indd';
    testDoc = 'reopen.indd';
    assert(/reopen\.indd \(ACTIVE\)[^\n]*modified=false path=[^\n]*reopen\.indd/.test(await call('list_open_documents')), 'the saved document shows its path');

    // close, reopen, edit, plain save (the reported failing case)
    await call('close_document', { name: 'reopen.indd' });
    out = await call('open_document', { filePath: file });
    assert(/Document opened: reopen\.indd/.test(out), out);
    const reopened = idOf(await call('list_text_frames'));
    eq(await text(reopened), 'version 1', 'the file has version 1');
    await call('edit_text_frame', { frameId: reopened, content: 'version 2' });
    let list = await call('list_open_documents');
    assert(/reopen\.indd \(ACTIVE\)[^\n]*modified=true path=[^\n]*reopen\.indd/.test(list), `a modified document that has a file still shows its path: ${list}`);
    assert(/File Path: [^\n]*reopen\.indd/.test(await call('get_document_info')), 'get_document_info shows the file path of a modified document');
    out = await call('save_document', {});
    assert(/^Save Document: Document saved: reopen\.indd \(/.test(out), out);
    await call('close_document', { name: 'reopen.indd' });               // unmodified now: closes without confirmation
    await call('open_document', { filePath: file });
    eq(await text(idOf(await call('list_text_frames'))), 'version 2', 'the edit was persisted (verified by closing and reopening)');

    // same-path save: a plain save, no confirmation, no "already open" failure
    const again = idOf(await call('list_text_frames'));
    await call('edit_text_frame', { frameId: again, content: 'version 3' });
    out = await call('save_document', { filePath: file });
    assert(/plain save/.test(out) && /Document saved: reopen\.indd/.test(out), out);
    await call('save_document', { filePath: file, confirmDestructive: true });   // also fine with the flag
    await call('close_document', { name: 'reopen.indd' });
    await call('open_document', { filePath: file });
    eq(await text(idOf(await call('list_text_frames'))), 'version 3', 'the same-path save was persisted');

    // Save As to another name; overwriting a different existing file needs confirmation
    const second = path.join(dir, 'second.indd');
    await call('save_document', { filePath: second });
    assert(/second\.indd \(ACTIVE\)/.test(await call('list_open_documents')), 'renamed by Save As');
    createdDocs[createdDocs.indexOf('reopen.indd')] = 'second.indd';
    testDoc = 'second.indd';
    // the file reopen.indd exists (version 3) and is not open now
    e = await fails('save_document', { filePath: file }, /already exists.*confirmDestructive/);
    assert(fs.existsSync(file), 'nothing was overwritten');
    await call('save_document', { filePath: file, confirmDestructive: true });
    assert(/reopen\.indd \(ACTIVE\)/.test(await call('list_open_documents')), 'overwriting with confirmation works');
    createdDocs[createdDocs.indexOf('second.indd')] = 'reopen.indd';
    testDoc = 'reopen.indd';

    // saving to the file of ANOTHER open document is refused with a clear message
    await call('open_document', { filePath: second });
    const other = await newDoc({ preset: 'A5' });
    testDoc = other;
    e = await fails('save_document', { filePath: second, confirmDestructive: true }, /already open in the document second\.indd/);
    await call('close_document', { name: 'second.indd' });
    createdDocs.splice(createdDocs.indexOf('second.indd'), 1);
    await closeDoc(other);

    // a named document can be saved without being active; package works on a modified saved document
    await call('activate_document', { name: 'reopen.indd' });
    testDoc = 'reopen.indd';
    await call('edit_text_frame', { frameIndex: 0, content: 'version 4' });
    const pk = await call('package_document', { folderPath: path.join(dir, 'pkg'), confirmDestructive: true });
    assert(/Document packaged/.test(pk), `package_document works on a saved document with unsaved changes: ${pk}`);
    out = await call('save_document', { name: 'reopen.indd' });
    assert(/Document saved: reopen\.indd/.test(out), out);
    await fails('save_document', { name: 'No Such.indd' }, /Document not found/);
    await fails('save_document', { filePath: path.join(dir, 'x.txt') }, /must end with \.indd/);
    await fails('save_document', { filePath: '/etc/x.indd' }, /Access denied/);
    await closeDoc('reopen.indd');
    testDoc = flyer;
  });


  console.log('\nErrors');
  await test('errors are readable: no osascript paths, numeric codes for InDesign errors', async () => {
    const msg = await fails('create_text_frame', { content: 'x', paragraphStyle: 'Nope' });
    assert(!/osascript|\.scpt|Command failed/.test(msg), `no raw osascript output: ${msg}`);
    // an InDesign-level failure (the frame is far too small to lay out) still yields a short, path-free message
    const kept = scriptFiles().length;
    const e2 = await fails('create_paragraph_style', { name: 'Huge', fontSize: 100000 });
    assert(!/osascript|\.scpt|Command failed/.test(e2), `no raw osascript output: ${e2}`);
    assert(/Code: \d{2,}/.test(e2), `an InDesign error carries its numeric code: ${e2}`);
    assert(/debug script: /.test(e2), 'InDesign-level failures point at the kept script');
    assert(scriptFiles().length > kept, 'the failing script is kept for debugging');
    assert(!/Huge/.test(await call('list_styles', { styleType: 'paragraph' })), 'the half-made style was removed');
  });

  await test('arguments are validated against the tool schema before any script runs', async () => {
    await fails('create_text_frame', { content: 'x', x: '5"; app.quit(); "' }, /Invalid parameter "x".*a number/);
    await fails('create_text_frame', { content: 'x', alignment: 'LEFT_ALIGN); evil(' }, /Invalid parameter "alignment"/);
    await fails('list_text_frames', { pageIndex: {} }, /Invalid parameter "pageIndex"/);
    await fails('create_paragraph_style', {}, /Missing required parameter "name"/);
    await fails('group_objects', { ids: ['1; evil()', 2] }, /Invalid parameter "ids\[0\]"/);
    const ok = await call('list_page_items', { pageIndex: '0' }); // numeric strings are accepted
    assert(/PAGE ITEMS ON PAGE 1/.test(ok), ok);
  });

  await test('"No document open" is an error for every document tool (only when nothing else is open)', async () => {
    const listing = await call('list_open_documents');
    const others = listing.match(/OPEN DOCUMENTS \((\d+)\)/)?.[1];
    if (Number(others) > 1 || createdDocs.length !== Number(others)) {
      console.log(`        (skipped: other documents are open - never closed by the tests. Open: ${listing.replace(/\n/g, ' | ')}; created by the tests: ${createdDocs.join(', ')})`);
      return;
    }
    await closeDoc(flyer);
    for (const [tool, args] of [
      ['create_paragraph_style', { name: 'X' }], ['create_text_frame', { content: 'x' }], ['apply_paragraph_style', { styleName: 'X', frameIndex: 0 }],
      ['edit_text_frame', { frameIndex: 0, content: 'x' }], ['list_text_frames', {}], ['create_rectangle', { x: 0, y: 0, width: 1, height: 1 }],
      ['list_page_items', {}], ['delete_object', { id: 1 }], ['get_document_info', {}], ['create_color_swatch', { name: 'X', colorValues: [0, 0, 0, 0] }],
      ['export_pdf', { filePath: path.join(TMP, 'a.pdf'), confirmDestructive: true }],
      ['create_path', { points: [{ x: 0, y: 0 }, { x: 10, y: 10 }] }], ['create_path_from_svg', { d: 'M0 0 L10 10 L0 10 Z' }],
      ['edit_path_points', { id: 1 }], ['create_line', { x1: 0, y1: 0, x2: 5, y2: 5 }], ['create_polygon', { x: 10, y: 10, radius: 5 }],
      ['convert_shape', { id: 1, to: 'oval' }], ['set_corner_options', { id: 1, option: 'rounded' }], ['pathfinder', { operation: 'union', ids: [1, 2] }],
      ['set_object_fill', { id: 1, swatch: 'none' }], ['set_object_stroke', { id: 1, weight: 1 }], ['create_gradient_swatch', { name: 'G', stops: [{ swatch: 'Black', position: 0 }, { swatch: 'Black', position: 100 }] }],
      ['set_gradient_feather', { id: 1 }], ['place_image_in_shape', { shapeId: 1, imagePath: png }], ['duplicate_object', { id: 1 }],
      ['align_objects', { ids: [1, 2], alignment: 'left' }], ['distribute_objects', { ids: [1, 2, 3], distribution: 'left_edges' }], ['get_object_info', { id: 1 }],
      ['render_preview', {}], ['measure_text', { frameIndex: 0 }], ['place_graphic', { filePath: path.join(TMP, 'logo.svg') }],
      ['create_graphic_from_svg', { svg: SVG_MARKUP }], ['create_cmyk_pdf_shape', { shapes: [{ d: 'M0 0 L10 10 L0 10 Z', fill: [0, 0, 0, 100] }] }],
      ['update_color_swatch', { name: 'X', colorValues: [0, 0, 0, 0] }], ['delete_color_swatch', { name: 'X' }], ['list_color_swatches', {}],
      ['set_object_opacity', { id: 1, opacity: 50 }], ['rotate_object', { id: 1, angle: 5 }], ['group_objects', { ids: [1, 2] }], ['list_open_documents_dummy', null],
    ].filter(([tool]) => tool !== 'list_open_documents_dummy')) {
      await fails(tool, args, /No document open/);
    }
  });
}

// ---------------------------------------------------------------- run
try {
  await main();
} catch (error) {
  results.push({ name: 'test run', ok: false, error });
  console.log(`\nFATAL: ${error.message}`);
} finally {
  // Close only what the tests created
  for (const name of [...createdDocs]) {
    try { await call('close_document', { name, confirmDestructive: true }); } catch { /* already closed */ }
  }
  try { await client?.close(); } catch { /* ignore */ }
  fs.rmSync(TMP, { recursive: true, force: true });
}

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} passed`);
for (const f of failed) console.log(`  FAILED: ${f.name}\n    ${f.error?.stack?.split('\n').slice(0, 3).join('\n    ')}`);
process.exit(failed.length ? 1 : 0);
