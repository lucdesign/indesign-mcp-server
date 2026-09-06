/**
 * Test harness.
 *
 * Loads index.js and replaces executeInDesignScript with a collector, so the
 * generated ExtendScript can be inspected without InDesign running.
 *
 * index.js itself is not modified — the harness writes a copy into a temp
 * directory and removes the auto-start at the end of the file.
 */

import fs from 'fs';
import os from 'os';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath, pathToFileURL } from 'url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..');

let cached = null;

/** Load the server once, with script execution stubbed out. */
export async function loadServer() {
  if (cached) return cached;

  const dir = path.join(os.tmpdir(), `indesign-mcp-test-${crypto.randomBytes(6).toString('hex')}`);
  fs.mkdirSync(path.join(dir, 'lib'), { recursive: true });

  // Keep the real escaping helpers — they are part of what is under test.
  fs.copyFileSync(path.join(ROOT, 'lib', 'jsx-safe.js'), path.join(dir, 'lib', 'jsx-safe.js'));

  let code = fs.readFileSync(path.join(ROOT, 'index.js'), 'utf8');
  code = code.replace(/const server = new InDesignMCPServer\(\);\s*server\.run\(\)[^\n]*/, '');
  code += '\nexport { InDesignMCPServer, autoCaptureResult };\n';
  fs.writeFileSync(path.join(dir, 'index.js'), code, 'utf8');

  // Make node_modules reachable for the SDK import.
  try {
    fs.symlinkSync(path.join(ROOT, 'node_modules'), path.join(dir, 'node_modules'), 'junction');
  } catch {
    /* symlink not permitted; package.json below still sets the module type */
  }
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ type: 'module' }), 'utf8');

  const mod = await import(pathToFileURL(path.join(dir, 'index.js')).href);
  const server = new mod.InDesignMCPServer();

  // Collect scripts instead of sending them to InDesign.
  const captured = [];
  server.executeInDesignScript = async (script) => {
    captured.push(script);
    return '(mock)';
  };

  cached = { server, captured, autoCaptureResult: mod.autoCaptureResult, dir };
  return cached;
}

/**
 * Call a tool method and return the script it produced.
 * If the method throws — an argument was rejected, say — the error is
 * returned rather than raised.
 */
export async function runTool(method, args) {
  const { server, captured } = await loadServer();
  const before = captured.length;
  try {
    await server[method](args);
  } catch (e) {
    return { error: e, script: captured.length > before ? captured[captured.length - 1] : null };
  }
  return { error: null, script: captured.length > before ? captured[captured.length - 1] : null };
}

/** The auto-capture used by the executor, for tests that exercise it. */
export async function getAutoCapture() {
  const { autoCaptureResult } = await loadServer();
  return autoCaptureResult;
}

/**
 * Does an ExtendScript source parse?
 *
 * ES3 is a subset of the Node grammar, so new Function() works as a parser.
 * The reverse is not true: Node accepts ES2024 constructs that ExtendScript
 * does not, hence the additional deny-list.
 */
const ES3_FORBIDDEN = [
  [/`/, 'template literal (ExtendScript is ES3)'],
  [/\blet\s+[A-Za-z_$]/, 'let (ExtendScript only has var)'],
  [/\bconst\s+[A-Za-z_$]/, 'const (ExtendScript only has var)'],
  [/=>/, 'arrow function (ExtendScript is ES3)'],
];

export function parses(script) {
  try {
    // eslint-disable-next-line no-new-func
    new Function(script);
  } catch (e) {
    return { ok: false, message: e.message };
  }
  for (const [re, what] of ES3_FORBIDDEN) {
    if (re.test(script)) return { ok: false, message: 'not valid for ExtendScript: ' + what };
  }
  return { ok: true };
}
