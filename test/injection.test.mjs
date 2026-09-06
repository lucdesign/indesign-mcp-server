/**
 * Injection regression.
 *
 * Every payload must have one of two outcomes:
 *   (a) the argument is rejected, or
 *   (b) the value ends up in the script as data — an escaped string literal,
 *       never executable code.
 *
 * What is checked is the generated ExtendScript source, not runtime behaviour:
 * showing that the payload never leaves its literal is enough.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { runTool, parses } from './harness.mjs';
import {
  str, num, index, measure, enumOf, bool, ALLOWED, validateFilePath,
} from '../lib/jsx-safe.js';

const BS = String.fromCharCode(92);      // backslash
const LS = String.fromCharCode(0x2028);  // line separator
const PS = String.fromCharCode(0x2029);  // paragraph separator

const BASE = {
  x: 10, y: 10, width: 100, height: 50, pageIndex: 0, frameIndex: 0,
  fontSize: 12, fontFamily: 'Helvetica Neue', fontStyle: 'Regular',
  textColor: 'Black', alignment: 'LEFT_ALIGN', content: 'ok',
};

const PAYLOADS = {
  trailingBackslash: 'end' + BS,
  quoteBreakout: '"; app.system("calc"); var x="',
  indexBreakout: '0]; app.quit(); //',
  enumBreakout: 'LEFT_ALIGN; app.quit()',
  fontBreakout: 'Arial"); app.quit(); ("',
  lineSeparator: 'a' + LS + 'app.quit(); ' + PS + 'b',
  commentBreakout: '*/ app.quit(); /*',
};

/** Fails if anything in the script looks like escaped-out code. */
function assertNoBreakout(script, t, context) {
  const res = parses(script);
  assert.ok(res.ok, `${context}: script no longer parses — ${res.message}`);

  // An escaped payload always sits inside a string literal. Strip every
  // literal; nothing recognisable may remain.
  const withoutStrings = script
    .replace(/"(?:[^"\\]|\\.)*"/g, '""')
    .replace(/'(?:[^'\\]|\\.)*'/g, "''");
  assert.doesNotMatch(withoutStrings, /app\s*\.\s*quit\s*\(/, `${context}: app.quit() outside a literal`);
  assert.doesNotMatch(withoutStrings, /app\s*\.\s*system\s*\(/, `${context}: app.system() outside a literal`);
}

// ------------------------------------------------------- sharpness of the check

/**
 * Counter-check: the same assertion against deliberately insufficient
 * escaping.
 *
 * Without this test a green run would be no evidence, only the absence of a
 * finding — a check that never fires looks exactly like one with nothing to
 * find.
 *
 * This escapes quotes and newlines but not the backslash. A common mistake,
 * and precisely the gap str() closes.
 */
function insufficientEscaping(value) {
  return '"' + String(value).replace(/"/g, BS + '"').replace(/\n/g, BS + 'n') + '"';
}

/**
 * The breakout does not go through the quote — that gets escaped — but
 * through the backslash before it: `\` plus `"` becomes `\` plus `\"`, an
 * escaped backslash followed by the closing quote. The literal ends earlier
 * than intended and the rest is code. This is why str() does not scan for
 * quotes but relies on JSON.stringify.
 */
const BREAKOUT = 'x' + BS + '"; app.quit(); //';

test('counter-check: insufficient escaping allows code execution', (t) => {
  const brokenOut = 'textFrame.contents = ' + insufficientEscaping(BREAKOUT) + ';';

  // Evidence 1: app.quit() sits outside every string literal.
  const withoutStrings = brokenOut.replace(/"(?:[^"\\]|\\.)*"/g, '""');
  assert.match(withoutStrings, /app\s*\.\s*quit\s*\(/,
    'test setup is wrong — this payload does not break out at all');

  // Evidence 2: the assertion used by the regression tests fires on it.
  assert.throws(
    () => assertNoBreakout(brokenOut, t, 'counter-check'),
    /outside a literal|no longer parses/,
    'assertNoBreakout misses a real breakout — the tests would be worthless'
  );
});

test('counter-check: the same payload through str() does not break out', (t) => {
  assertNoBreakout('textFrame.contents = ' + str(BREAKOUT) + ';', t, 'hardened');
  assert.equal(eval(str(BREAKOUT)), BREAKOUT, 'value does not survive round-trip');
});

test('counter-check: the same payload through createTextFrame does not break out', async (t) => {
  const { error, script } = await runTool('createTextFrame', { ...BASE, content: BREAKOUT });
  assert.ok(!error, 'tool threw unexpectedly: ' + (error && error.message));
  assertNoBreakout(script, t, 'createTextFrame/breakout');
  assert.ok(script.includes(str(BREAKOUT)), 'payload is not present as an escaped literal');
});

// ------------------------------------------------------------------- helpers

test('str(): a trailing backslash does not break out', () => {
  assert.equal(eval(str(PAYLOADS.trailingBackslash)), PAYLOADS.trailingBackslash);
});

test('str(): embedded quotes stay data', () => {
  assert.equal(eval(str(PAYLOADS.quoteBreakout)), PAYLOADS.quoteBreakout);
});

test('str(): U+2028/U+2029 are escaped', () => {
  const out = str(PAYLOADS.lineSeparator);
  assert.doesNotMatch(out, new RegExp('[' + LS + PS + ']'), 'raw line separator in the literal');
  assert.equal(eval(out), PAYLOADS.lineSeparator);
});

test('num(): non-numeric values are rejected', () => {
  assert.throws(() => num(PAYLOADS.indexBreakout, { name: 'x' }), /Invalid number/);
  assert.throws(() => num('1; app.quit()', { name: 'x' }), /Invalid number/);
  assert.throws(() => num(NaN, { name: 'x' }), /Invalid number/);
});

test('index(): breakout attempts are rejected', () => {
  assert.throws(() => index(PAYLOADS.indexBreakout, { name: 'pageIndex' }), /Invalid index/);
  assert.throws(() => index(-1, { name: 'pageIndex' }), /Invalid index/);
  assert.throws(() => index(1.5, { name: 'pageIndex' }), /Invalid index/);
});

test('measure(): number plus known unit only', () => {
  assert.equal(measure(20, { unit: 'mm' }), '"20mm"');
  assert.throws(() => measure('20mm"; app.quit(); //', { name: 'x' }), /Invalid measurement/);
  assert.throws(() => measure(10, { unit: 'mm"; app.quit()' }), /Invalid unit/);
});

test('enumOf(): allow-list only', () => {
  assert.equal(enumOf('LEFT_ALIGN', ALLOWED.alignment), 'LEFT_ALIGN');
  assert.throws(
    () => enumOf(PAYLOADS.enumBreakout, ALLOWED.alignment, { name: 'alignment' }),
    /Invalid value/
  );
});

test('bool(): always yields a literal', () => {
  for (const v of ['true; app.quit()', {}, [], 'x', 0, null]) {
    assert.match(bool(v), /^(true|false)$/);
  }
});

test('validateFilePath(): system and outside paths are rejected', () => {
  const outside = process.platform === 'win32'
    ? 'C:' + BS + 'Windows' + BS + 'System32' + BS + 'x.indd'
    : '/etc/x.indd';
  assert.throws(() => validateFilePath(outside, [process.cwd()]), /Access denied/);
  assert.throws(() => validateFilePath('', [process.cwd()]), /Invalid file path/);
});

// ---------------------------------------------------------------- end-to-end

test('createTextFrame: content payloads stay data', async (t) => {
  let checked = 0;
  for (const [name, payload] of Object.entries(PAYLOADS)) {
    const { error, script } = await runTool('createTextFrame', { ...BASE, content: payload });
    if (error) continue;                    // rejection is a valid outcome
    assertNoBreakout(script, t, `content/${name}`);
    assert.ok(script.includes(str(payload)), `content/${name}: value is not an escaped literal`);
    checked++;
  }
  assert.equal(checked, Object.keys(PAYLOADS).length,
    'not every content payload reached a script — the test checks too little');
});

test('createTextFrame: fontFamily payload stays data', async (t) => {
  const { error, script } = await runTool('createTextFrame', { ...BASE, fontFamily: PAYLOADS.fontBreakout });
  if (!error) assertNoBreakout(script, t, 'fontFamily');
});

test('createTextFrame: pageIndex payload is rejected', async () => {
  const { error, script } = await runTool('createTextFrame', { ...BASE, pageIndex: PAYLOADS.indexBreakout });
  assert.ok(error, 'pageIndex payload was not rejected');
  assert.match(error.message, /Invalid index/);
  assert.equal(script, null, 'a script was produced despite rejection');
});

test('createTextFrame: alignment payload is rejected', async () => {
  const { error } = await runTool('createTextFrame', { ...BASE, alignment: PAYLOADS.enumBreakout });
  assert.ok(error, 'alignment payload was not rejected');
  assert.match(error.message, /Invalid value/);
});

test('createTextFrame: coordinate payload is rejected', async () => {
  const { error } = await runTool('createTextFrame', { ...BASE, x: '10mm"]; app.quit(); //' });
  assert.ok(error, 'x payload was not rejected');
  assert.match(error.message, /Invalid measurement/);
});

test('findReplaceText: search and replacement stay data', async (t) => {
  const { error, script } = await runTool('findReplaceText', {
    findText: PAYLOADS.quoteBreakout,
    replaceText: PAYLOADS.trailingBackslash,
    scope: 'document',
  });
  if (!error) assertNoBreakout(script, t, 'findReplaceText');
});

test('insertMarkdownText: markdown stays data', async (t) => {
  const { error, script } = await runTool('insertMarkdownText', {
    markdownText: PAYLOADS.quoteBreakout + '\n' + PAYLOADS.trailingBackslash,
    frameIndex: 0, pageIndex: 0,
  });
  if (!error) assertNoBreakout(script, t, 'insertMarkdownText');
});

test('createColorSwatch: colour values must be numbers', async () => {
  const { error } = await runTool('createColorSwatch', {
    name: 'x', colorModel: 'PROCESS', colorSpace: 'CMYK',
    colorValues: ['0]; app.quit(); //', 0, 0, 0],
  });
  assert.ok(error, 'colour value payload was not rejected');
});

test('populateTable: table data stays data', async (t) => {
  const { error, script } = await runTool('populateTable', {
    tableIndex: 0, pageIndex: 0,
    data: [[PAYLOADS.quoteBreakout, PAYLOADS.lineSeparator]],
  });
  if (!error) assertNoBreakout(script, t, 'populateTable');
});

test('execute_indesign_code stays locked without the opt-in', async () => {
  const { error, script } = await runTool('executeInDesignCode', { code: 'app.quit();' });
  const locked = Boolean(error) || script === null || /disabled|not enabled/i.test(String(script));
  assert.ok(locked, 'executeInDesignCode was reachable without INDESIGN_ALLOW_ARBITRARY_CODE');
});

test('the destructive-operation guard actually blocks', async () => {
  const { loadServer } = await import('./harness.mjs');
  const { server } = await loadServer();
  assert.throws(
    () => server.validateDestructiveOperation({}, 'test operation', 'test target'),
    /confirmation required/i,
    'the guard did not throw - callers would carry on and perform the operation'
  );
  assert.doesNotThrow(
    () => server.validateDestructiveOperation({ confirmDestructive: true }, 'test', 'target'),
    'the guard blocked a confirmed operation'
  );
});
