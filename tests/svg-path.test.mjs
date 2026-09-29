#!/usr/bin/env node
// Offline unit tests for the SVG path parser (no InDesign needed).
import assert from 'assert';
import { svgPathToSubpaths, fitSubpaths, subpathBounds, polygonPoints } from '../lib/svg-path.js';

let passed = 0;
const test = (name, fn) => { try { fn(); passed++; console.log(`  ok    ${name}`); } catch (e) { console.log(`  FAIL  ${name}\n        ${e.message}`); process.exitCode = 1; } };
const near = (a, b, tol, msg) => assert.ok(Math.abs(a - b) <= tol, `${msg}: expected ~${b}, got ${a}`);
const pt = (p, x, y, msg) => { near(p.x, x, 1e-6, `${msg} x`); near(p.y, y, 1e-6, `${msg} y`); };

console.log('SVG path parser');
const CURVE = 'M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z';

test('the flyer band: 5 anchors, closed, handles as written, S reflects the previous handle', () => {
  const { subpaths } = svgPathToSubpaths(CURVE);
  assert.strictEqual(subpaths.length, 1);
  const sp = subpaths[0];
  assert.strictEqual(sp.closed, true);
  assert.strictEqual(sp.points.length, 5);
  pt(sp.points[0], 0, 8, 'p0'); pt(sp.points[1], 120, 8, 'p1'); pt(sp.points[2], 216, 6, 'p2'); pt(sp.points[3], 216, 50, 'p3'); pt(sp.points[4], 0, 50, 'p4');
  pt(sp.points[0].right, 40, 0, 'p0 right');
  pt(sp.points[1].left, 80, 14, 'p1 left');
  pt(sp.points[1].right, 160, 2, 'p1 right (reflection of 80,14 about 120,8)');
  pt(sp.points[2].left, 190, 0, 'p2 left');
  pt(sp.points[3].left, 216, 50, 'straight segment has no handle');
});

test('fitting into a viewBox and target rectangle (mm)', () => {
  const { subpaths } = svgPathToSubpaths(CURVE);
  const fit = fitSubpaths(subpaths, { viewBox: [0, 0, 216, 50], x: -3, y: 55, width: 216, height: 48 });
  pt(fit[0].points[0], -3, 55 + 8 * 0.96, 'first');
  pt(fit[0].points[2], 213, 55 + 6 * 0.96, 'third');
  pt(fit[0].points[3], 213, 103, 'bottom right');
  pt(fit[0].points[0].right, -3 + 40, 55, 'handle scaled too');
});

test('the top edge is a gentle S-curve (sampled deviation from the straight chord)', () => {
  const { subpaths } = svgPathToSubpaths(CURVE);
  const fit = fitSubpaths(subpaths, { viewBox: [0, 0, 216, 50], x: -3, y: 55, width: 216, height: 48 });
  const [a, b, c] = fit[0].points;
  const yAt = (p0, p1, p2, p3, t) => { const u = 1 - t; return u ** 3 * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t ** 3 * p3; };
  let min = Infinity, max = -Infinity;
  for (let i = 0; i <= 100; i++) {
    const y1 = yAt(a.y, a.right.y, b.left.y, b.y, i / 100);
    min = Math.min(min, y1); max = Math.max(max, y1);
  }
  assert.ok(min < a.y - 1 && max > a.y - 0.01 || max > b.y, 'the first half dips above the chord');
  const bounds = subpathBounds(fit);
  assert.ok(bounds.height > 44 && bounds.height < 48.5, `bounds height ${bounds.height}`);
});

test('relative commands and implicit repeated commands', () => {
  const { subpaths } = svgPathToSubpaths('m 10 10 20 0 0 20 h -20 z');
  assert.deepStrictEqual(subpaths[0].points.map((p) => [p.x, p.y]), [[10, 10], [30, 10], [30, 30], [10, 30]]);
  assert.strictEqual(subpaths[0].closed, true);
});

test('H, V, relative c/s and multiple sub-paths (compound)', () => {
  const { subpaths } = svgPathToSubpaths('M0 0 H10 V10 H0 Z M2 2 h6 v6 h-6 z');
  assert.strictEqual(subpaths.length, 2);
  assert.deepStrictEqual(subpaths[1].points.map((p) => [p.x, p.y]), [[2, 2], [8, 2], [8, 8], [2, 8]]);
  const c = svgPathToSubpaths('M0 0 c 10 0 20 10 30 10 s 20 10 30 0').subpaths[0];
  pt(c.points[1], 30, 10, 'c end'); pt(c.points[2], 60, 10, 's end'); pt(c.points[1].right, 40, 10, 's first control reflects (20,10) about (30,10)');
  assert.strictEqual(c.closed, false);
});

test('quadratic curves (Q, T) become cubic curves', () => {
  const sp = svgPathToSubpaths('M0 0 Q 50 100 100 0 T 200 0').subpaths[0];
  assert.strictEqual(sp.points.length, 3);
  pt(sp.points[0].right, 100 / 3, 200 / 3, 'first cubic control = p0 + 2/3 (q - p0)');
  pt(sp.points[1].left, 100 - 100 / 3, 200 / 3, 'second cubic control = p + 2/3 (q - p)');
  pt(sp.points[1].right, 100 + 100 / 3, -200 / 3, 'T reflects the quadratic control point');
});

test('arcs: a circle from two arcs stays round', () => {
  const { subpaths } = svgPathToSubpaths('M 0 50 A 50 50 0 1 1 100 50 A 50 50 0 1 1 0 50 Z');
  assert.strictEqual(subpaths[0].points.length, 4, 'four 90-degree segments');
  const b = subpathBounds(subpaths);
  near(b.x, 0, 0.05, 'x'); near(b.y, 0, 0.05, 'y'); near(b.width, 100, 0.1, 'w'); near(b.height, 100, 0.1, 'h');
  for (const p of subpaths[0].points) near(Math.hypot(p.x - 50, p.y - 50), 50, 1e-6, 'anchor on the circle');
});

test('arc flags without separators, degenerate arcs and radii scaling', () => {
  const sp = svgPathToSubpaths('M0 0a1 1 0 00.5.5').subpaths[0];
  assert.ok(sp.points.length >= 2);
  assert.strictEqual(svgPathToSubpaths('M0 0 A0 5 0 0 1 10 10').subpaths[0].points.length, 2, 'zero radius is a line');
  const big = svgPathToSubpaths('M0 0 A1 1 0 0 1 100 0').subpaths[0];   // radii too small: scaled up to a half circle
  near(subpathBounds([big]).width, 100, 0.1, 'half circle width');
});

test('numbers: exponents, signs, leading dots', () => {
  const sp = svgPathToSubpaths('M1e1-5L.5.25 -2E0,3').subpaths[0];
  pt(sp.points[0], 10, -5, 'exp'); pt(sp.points[1], 0.5, 0.25, 'dots'); pt(sp.points[2], -2, 3, 'comma');
});

test('fit without viewBox uses the path bounds; one dimension keeps the aspect ratio', () => {
  const { subpaths } = svgPathToSubpaths('M10 10 L30 10 L30 20 L10 20 Z');
  const a = fitSubpaths(subpaths, { x: 0, y: 0, width: 100, height: 50 });
  pt(a[0].points[0], 0, 0, 'a'); pt(a[0].points[2], 100, 50, 'a2');
  const b = fitSubpaths(subpaths, { x: 5, y: 5, width: 40 });
  pt(b[0].points[2], 45, 25, 'aspect kept: 20x10 -> 40x20');
});

test('errors are clear', () => {
  assert.throws(() => svgPathToSubpaths('L 1 2'), /must start with M/);
  assert.throws(() => svgPathToSubpaths('M 0 0 X 1'), /unexpected "X"/);
  assert.throws(() => svgPathToSubpaths('M 0'), /number expected/);
  assert.throws(() => svgPathToSubpaths(''), /Empty/);
  assert.throws(() => svgPathToSubpaths('M 5 5'), /no drawable/);
});

console.log('\nPolygons and stars');
test('regular hexagon, star and rounded corners', () => {
  const hex = polygonPoints({ cx: 50, cy: 50, radius: 10, sides: 6 });
  assert.strictEqual(hex.length, 6);
  hex.forEach((p) => near(Math.hypot(p.x - 50, p.y - 50), 10, 1e-9, 'on circle'));
  pt(hex[0], 50, 40, 'first vertex points up');
  const star = polygonPoints({ cx: 0, cy: 0, radius: 10, sides: 5, innerRatio: 0.4 });
  assert.strictEqual(star.length, 10);
  near(Math.hypot(star[1].x, star[1].y), 4, 1e-9, 'inner radius');
  const round = polygonPoints({ cx: 0, cy: 0, radius: 10, sides: 6, cornerRadius: 2 });
  assert.strictEqual(round.length, 12, 'two anchors per rounded corner');
  round.forEach((p) => assert.ok(Math.hypot(p.x, p.y) <= 10 + 1e-9, 'rounded points stay inside the polygon'));
});

console.log(`\n${passed} passed`);
