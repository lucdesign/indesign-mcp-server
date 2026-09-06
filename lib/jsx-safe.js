/**
 * jsx-safe.js
 * Escaping and validation helpers for values that end up in ExtendScript.
 *
 * ExtendScript is not sandboxed: it reads and writes files and starts system
 * commands through app.system(). A value placed unchecked into source text is
 * therefore potentially executable code. This is not only about direct input —
 * tool arguments frequently originate from documents, CSV files or web
 * content, that is, from sources the caller does not control.
 *
 * Rule: no value is ever interpolated into a script as-is.
 * Strings -> str(), numbers -> num(), enums -> enumOf(), booleans -> bool().
 */

import path from 'path';
import os from 'os';

/** Embed a string as a JS/JSX literal, including backslashes and Unicode. */
export function str(value) {
  if (value === undefined || value === null) return '""';
  // JSON.stringify escapes backslashes, quotes and control characters
  // correctly. U+2028/U+2029 need separate handling: ES3 parsers treat them
  // as line terminators, which would end the literal early.
  return JSON.stringify(String(value))
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

/** Require a finite number. Throws on anything else. */
export function num(value, { min = -1e9, max = 1e9, name = 'value' } = {}) {
  const n = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(n)) {
    throw new Error(`Invalid number for '${name}': ${JSON.stringify(value)}`);
  }
  if (n < min || n > max) {
    throw new Error(`Value for '${name}' out of range [${min}, ${max}]: ${n}`);
  }
  return String(n);
}

/** Require a non-negative integer, e.g. a page index. */
export function index(value, { name = 'index', max = 100000 } = {}) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0 || n > max) {
    throw new Error(`Invalid index for '${name}': ${JSON.stringify(value)}`);
  }
  return String(n);
}

/** Boolean as a JSX literal. Never passes a caller-supplied value through. */
export function bool(value) {
  return value ? 'true' : 'false';
}

/**
 * Allow only values from a list — for anything written into the script as an
 * identifier rather than a string, such as Justification enum members.
 */
export function enumOf(value, allowed, { name = 'enum' } = {}) {
  const v = String(value);
  if (!allowed.includes(v)) {
    throw new Error(
      `Invalid value for '${name}': ${JSON.stringify(value)}. Allowed: ${allowed.join(', ')}`
    );
  }
  return v;
}

/** Allow-lists for the InDesign enums used by the tools. */
export const ALLOWED = {
  alignment: [
    'LEFT_ALIGN', 'CENTER_ALIGN', 'RIGHT_ALIGN',
    'LEFT_JUSTIFIED', 'RIGHT_JUSTIFIED', 'CENTER_JUSTIFIED', 'FULLY_JUSTIFIED',
  ],
  fitOption: [
    'CONTENT_TO_FRAME', 'FRAME_TO_CONTENT', 'PROPORTIONALLY',
    'FILL_PROPORTIONALLY', 'CENTER_CONTENT', 'APPLY_FRAME_FITTING_OPTIONS',
  ],
  colorSpace: ['CMYK', 'RGB', 'LAB'],
  colorModel: ['PROCESS', 'SPOT', 'REGISTRATION'],
  imageFormat: ['PNG', 'JPG', 'JPEG', 'TIFF', 'GIF'],
  pdfPreset: [
    'HighQualityPrint', 'PressQuality', 'SmallestFileSize',
    'PDFX1a2001', 'PDFX32002', 'PDFX42008',
  ],
  // UIColors — layer colours in the InDesign interface.
  uiColor: [
    'LIGHT_BLUE', 'RED', 'GREEN', 'BLUE', 'YELLOW', 'MAGENTA', 'CYAN',
    'GRAY', 'BLACK', 'ORANGE', 'DARK_GREEN', 'TEAL', 'TAN', 'BROWN',
    'VIOLET', 'GOLD', 'DARK_BLUE', 'PINK', 'LAVENDER', 'BRICK_RED',
    'OLIVE_GREEN', 'PEACH', 'BURGUNDY', 'GRASS_GREEN', 'OCHRE',
    'PURPLE', 'LIGHT_GRAY', 'CHARCOAL', 'GRID_BLUE', 'GRID_ORANGE',
    'FIESTA', 'LIGHT_OLIVE', 'LIPSTICK', 'CUTE_TEAL', 'SULPHUR',
    'GRID_GREEN', 'WHITE',
  ],
};

/**
 * Embed a data structure as a JS literal (arrays and objects from tool
 * arguments). JSON.stringify already produces valid JS, but lets U+2028 and
 * U+2029 through — line terminators to an ES3 parser.
 */
export function json(value) {
  return JSON.stringify(value === undefined ? null : value)
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

/** A list of numbers as JSX array contents, e.g. for colorValue. */
export function numList(values, { name = 'values', min, max } = {}) {
  if (!Array.isArray(values)) {
    throw new Error(`Expected a list for '${name}': ${JSON.stringify(values)}`);
  }
  const opts = { name };
  if (min !== undefined) opts.min = min;
  if (max !== undefined) opts.max = max;
  return values.map((v) => num(v, opts)).join(', ');
}

/**
 * A measurement with a unit, e.g. "20mm". Number plus known unit only, and it
 * supplies its own quotes — so no additional quoting at the call site.
 */
export function measure(value, { unit = 'mm', name = 'measure' } = {}) {
  const allowedUnits = ['mm', 'cm', 'in', 'pt', 'px', 'p'];
  if (!allowedUnits.includes(unit)) {
    throw new Error(`Invalid unit for '${name}': ${unit}`);
  }
  const n = Number(value);
  if (!Number.isFinite(n)) {
    throw new Error(`Invalid measurement for '${name}': ${JSON.stringify(value)}`);
  }
  return `"${n}${unit}"`;
}

/**
 * Directories file operations are confined to: the home directory plus
 * anything listed in INDESIGN_ALLOWED_DIRS.
 *
 * The list is split on path.delimiter — ';' on Windows. Splitting on ':'
 * would tear a drive letter apart: 'C:\projects' becomes ['C', '\projects'],
 * and the second entry resolves on every drive.
 */
export function buildAllowedDirs() {
  const dirs = [os.homedir()];
  const fromEnv = process.env.INDESIGN_ALLOWED_DIRS;
  if (fromEnv) {
    for (const d of fromEnv.split(path.delimiter)) {
      const trimmed = d.trim();
      if (trimmed) dirs.push(trimmed);
    }
  }
  return dirs.map((d) => path.resolve(d));
}

const WIN_FORBIDDEN = [
  'C:\\Windows',
  'C:\\Program Files',
  'C:\\Program Files (x86)',
  'C:\\ProgramData',
];
const POSIX_FORBIDDEN = ['/etc', '/System', '/usr', '/bin', '/sbin', '/Library'];

/**
 * Path validation for Windows and macOS.
 *
 * Platform-dependent deny-list, case-insensitive comparison on Windows, UNC
 * paths and alternate data streams rejected. A POSIX-only deny-list would
 * match nothing on Windows, leaving the system directories unprotected.
 */
export function validateFilePath(filePath, allowedDirs = buildAllowedDirs()) {
  if (!filePath || typeof filePath !== 'string') {
    throw new Error('Invalid file path.');
  }
  // Reject UNC paths and alternate data streams.
  if (filePath.startsWith('\\\\') || filePath.includes(':', 2)) {
    if (process.platform === 'win32' && /:[^\\/]/.test(filePath.slice(2))) {
      throw new Error(`Invalid path (alternate data stream suspected): ${filePath}`);
    }
  }

  const resolved = path.resolve(filePath);
  const cmp = process.platform === 'win32' ? (s) => s.toLowerCase() : (s) => s;

  const inAllowed = allowedDirs.some((dir) => {
    const d = cmp(path.resolve(dir));
    const r = cmp(resolved);
    return r === d || r.startsWith(d + path.sep);
  });
  if (!inAllowed) {
    throw new Error(
      `Access denied: '${filePath}' is outside the allowed directories (${allowedDirs.join(', ')}).`
    );
  }

  const forbidden = process.platform === 'win32' ? WIN_FORBIDDEN : POSIX_FORBIDDEN;
  const hit = forbidden.some((f) => {
    const d = cmp(path.resolve(f));
    const r = cmp(resolved);
    return r === d || r.startsWith(d + path.sep);
  });
  if (hit) {
    throw new Error(`Access denied: system directory '${resolved}'.`);
  }

  return resolved;
}

/** A path for ExtendScript File() — forward slashes, safe as a literal. */
export function jsxPath(resolvedPath) {
  return str(String(resolvedPath).replace(/\\/g, '/'));
}
