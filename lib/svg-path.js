// SVG path data -> native Bezier anchor points.
//
// Supports M m L l H h V v C c S s Q q T t A a Z z (absolute and relative, implicit repeated
// commands, compact arc flags). Quadratic Beziers and elliptical arcs are converted to cubic
// Beziers. The result is a list of sub-paths, each a list of anchor points with explicit
// handles; a straight segment has handles equal to its anchor.
//
//   const { subpaths } = svgPathToSubpaths('M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z');
//   const fitted = fitSubpaths(subpaths, { viewBox: [0, 0, 216, 50], x: -3, y: 55, width: 216, height: 48 });

const ARG_COUNT = { M: 2, L: 2, H: 1, V: 1, C: 6, S: 4, Q: 4, T: 2, A: 7, Z: 0 };
const EPS = 1e-9;

function tokenize(d) {
  const tokens = [];
  const src = String(d);
  let i = 0;
  const number = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/;
  const skip = () => { while (i < src.length && /[\s,]/.test(src[i])) i++; };
  while (i < src.length) {
    skip();
    if (i >= src.length) break;
    const cmd = src[i];
    if (!/[MmLlHhVvCcSsQqTtAaZz]/.test(cmd)) throw new Error(`Invalid SVG path data: unexpected "${cmd}" at position ${i}`);
    i++;
    const upper = cmd.toUpperCase();
    const args = [];
    if (upper !== 'Z') {
      for (;;) {
        skip();
        if (i >= src.length || /[A-Za-z]/.test(src[i])) break;
        const group = [];
        for (let k = 0; k < ARG_COUNT[upper]; k++) {
          skip();
          // Arc flags are single characters (0/1) that may be written without separators
          if (upper === 'A' && (k === 3 || k === 4)) {
            if (src[i] !== '0' && src[i] !== '1') throw new Error(`Invalid SVG path data: arc flag expected at position ${i}`);
            group.push(Number(src[i]));
            i++;
            continue;
          }
          const m = number.exec(src.slice(i));
          if (!m) throw new Error(`Invalid SVG path data: number expected at position ${i} (command ${cmd})`);
          group.push(Number(m[0]));
          i += m[0].length;
        }
        args.push(group);
      }
      if (args.length === 0) throw new Error(`Invalid SVG path data: command ${cmd} has no arguments`);
    }
    tokens.push({ cmd, args });
  }
  return tokens;
}

// Endpoint arc -> cubic Beziers (SVG spec F.6.5 / F.6.6), split into segments of at most 90 degrees
function arcToCubics(x1, y1, rx, ry, phiDeg, fa, fs, x2, y2) {
  if (Math.abs(x1 - x2) < EPS && Math.abs(y1 - y2) < EPS) return [];
  rx = Math.abs(rx); ry = Math.abs(ry);
  if (rx < EPS || ry < EPS) return [{ line: true, x: x2, y: y2 }];
  const phi = (phiDeg * Math.PI) / 180, cos = Math.cos(phi), sin = Math.sin(phi);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) { const s = Math.sqrt(lambda); rx *= s; ry *= s; }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let coef = den === 0 ? 0 : Math.sqrt(Math.max(0, num / den));
  if (fa === fs) coef = -coef;
  const cxp = (coef * rx * y1p) / ry, cyp = (-coef * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux, uy, vx, vy) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dTheta = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!fs && dTheta > 0) dTheta -= 2 * Math.PI;
  else if (fs && dTheta < 0) dTheta += 2 * Math.PI;

  const count = Math.max(1, Math.ceil(Math.abs(dTheta) / (Math.PI / 2) - 1e-9));
  const delta = dTheta / count;
  const t = (4 / 3) * Math.tan(delta / 4);
  const out = [];
  let a = theta1;
  const point = (ang) => {
    const ca = Math.cos(ang), sa = Math.sin(ang);
    return [cos * rx * ca - sin * ry * sa + cx, sin * rx * ca + cos * ry * sa + cy];
  };
  const deriv = (ang) => {
    const ca = Math.cos(ang), sa = Math.sin(ang);
    return [-cos * rx * sa - sin * ry * ca, -sin * rx * sa + cos * ry * ca];
  };
  for (let i = 0; i < count; i++) {
    const b = a + delta;
    const p0 = point(a), p3 = i === count - 1 ? [x2, y2] : point(b);
    const d0 = deriv(a), d3 = deriv(b);
    out.push({ c1: [p0[0] + t * d0[0], p0[1] + t * d0[1]], c2: [p3[0] - t * d3[0], p3[1] - t * d3[1]], p: p3 });
    a = b;
  }
  return out;
}

// Parse path data into sub-paths of anchor points with handles
export function svgPathToSubpaths(d) {
  const tokens = tokenize(d);
  if (tokens.length === 0) throw new Error('Empty SVG path data');
  if (!/[Mm]/.test(tokens[0].cmd)) throw new Error('Invalid SVG path data: it must start with M or m');

  const subs = [];       // { start, segs: [{ c1, c2, p, line }], closed }
  let cur = [0, 0], start = [0, 0], sub = null;
  let lastCubicC2 = null, lastQuadC = null;   // for S/T reflection

  const begin = (pt) => { sub = { start: pt.slice(), segs: [], closed: false }; subs.push(sub); start = pt.slice(); };
  const ensureSub = () => { if (!sub || sub.closed) begin(cur); };
  const lineTo = (p) => { ensureSub(); sub.segs.push({ line: true, p: p.slice() }); cur = p.slice(); };
  const cubicTo = (c1, c2, p) => { ensureSub(); sub.segs.push({ c1, c2, p: p.slice() }); cur = p.slice(); };

  for (const { cmd, args } of tokens) {
    const rel = cmd === cmd.toLowerCase();
    const upper = cmd.toUpperCase();
    const ox = rel ? cur[0] : 0, oy = rel ? cur[1] : 0;
    if (upper === 'Z') {
      if (sub && !sub.closed) sub.closed = true;
      cur = start.slice();
      lastCubicC2 = lastQuadC = null;
      continue;
    }
    args.forEach((a, index) => {
      const bx = rel ? cur[0] : 0, by = rel ? cur[1] : 0;   // relative to the current point at the start of this group
      let nextCubic = null, nextQuad = null;
      switch (upper) {
        case 'M': {
          const pt = [a[0] + bx, a[1] + by];
          if (index === 0) { cur = pt; begin(pt); } else lineTo(pt);
          break;
        }
        case 'L': lineTo([a[0] + bx, a[1] + by]); break;
        case 'H': lineTo([a[0] + bx, cur[1]]); break;
        case 'V': lineTo([cur[0], a[0] + by]); break;
        case 'C': {
          const c1 = [a[0] + bx, a[1] + by], c2 = [a[2] + bx, a[3] + by], p = [a[4] + bx, a[5] + by];
          cubicTo(c1, c2, p); nextCubic = c2; break;
        }
        case 'S': {
          const c1 = lastCubicC2 ? [2 * cur[0] - lastCubicC2[0], 2 * cur[1] - lastCubicC2[1]] : cur.slice();
          const c2 = [a[0] + bx, a[1] + by], p = [a[2] + bx, a[3] + by];
          cubicTo(c1, c2, p); nextCubic = c2; break;
        }
        case 'Q': {
          const q = [a[0] + bx, a[1] + by], p = [a[2] + bx, a[3] + by];
          const p0 = cur.slice();
          cubicTo([p0[0] + (2 / 3) * (q[0] - p0[0]), p0[1] + (2 / 3) * (q[1] - p0[1])], [p[0] + (2 / 3) * (q[0] - p[0]), p[1] + (2 / 3) * (q[1] - p[1])], p);
          nextQuad = q; break;
        }
        case 'T': {
          const q = lastQuadC ? [2 * cur[0] - lastQuadC[0], 2 * cur[1] - lastQuadC[1]] : cur.slice();
          const p = [a[0] + bx, a[1] + by], p0 = cur.slice();
          cubicTo([p0[0] + (2 / 3) * (q[0] - p0[0]), p0[1] + (2 / 3) * (q[1] - p0[1])], [p[0] + (2 / 3) * (q[0] - p[0]), p[1] + (2 / 3) * (q[1] - p[1])], p);
          nextQuad = q; break;
        }
        case 'A': {
          const p = [a[5] + bx, a[6] + by];
          const pieces = arcToCubics(cur[0], cur[1], a[0], a[1], a[2], a[3], a[4], p[0], p[1]);
          if (pieces.length === 0) break;
          for (const piece of pieces) {
            if (piece.line) lineTo([piece.x, piece.y]); else cubicTo(piece.c1, piece.c2, piece.p);
          }
          break;
        }
        default: break;
      }
      lastCubicC2 = nextCubic;
      lastQuadC = nextQuad;
    });
  }

  const subpaths = [];
  for (const s of subs) {
    if (s.segs.length === 0) continue;   // lone "M x y" draws nothing
    const points = [{ x: s.start[0], y: s.start[1], left: { x: s.start[0], y: s.start[1] }, right: { x: s.start[0], y: s.start[1] } }];
    for (const seg of s.segs) {
      const prev = points[points.length - 1];
      if (!seg.line) prev.right = { x: seg.c1[0], y: seg.c1[1] };
      points.push({
        x: seg.p[0], y: seg.p[1],
        left: seg.line ? { x: seg.p[0], y: seg.p[1] } : { x: seg.c2[0], y: seg.c2[1] },
        right: { x: seg.p[0], y: seg.p[1] },
      });
    }
    if (s.closed) {
      // A closed path that ends where it started shares that anchor: merge the last point into the first
      const first = points[0], last = points[points.length - 1];
      if (points.length > 1 && Math.abs(first.x - last.x) < EPS && Math.abs(first.y - last.y) < EPS) {
        first.left = last.left;
        points.pop();
      }
    }
    subpaths.push({ points, closed: s.closed });
  }
  if (subpaths.length === 0) throw new Error('The SVG path data contains no drawable segments');
  return { subpaths };
}

const cubicAt = (p0, p1, p2, p3, t) => {
  const u = 1 - t;
  return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3;
};

// Tight bounds of the drawn curve (sampled), in the path's own coordinates
export function subpathBounds(subpaths) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  const add = (x, y) => { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); };
  for (const sp of subpaths) {
    const n = sp.points.length;
    const segCount = sp.closed ? n : n - 1;
    for (let i = 0; i < segCount; i++) {
      const a = sp.points[i], b = sp.points[(i + 1) % n];
      for (let s = 0; s <= 40; s++) {
        const t = s / 40;
        add(cubicAt(a.x, a.right.x, b.left.x, b.x, t), cubicAt(a.y, a.right.y, b.left.y, b.y, t));
      }
    }
    if (n === 1) add(sp.points[0].x, sp.points[0].y);
  }
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
}

// Scale sub-paths from a viewBox (or the path bounds) into a target rectangle.
// Missing width/height: the other one keeps the aspect ratio; both missing: 1 unit = 1 target unit.
export function fitSubpaths(subpaths, { viewBox, x = 0, y = 0, width, height } = {}) {
  const box = viewBox
    ? { x: viewBox[0], y: viewBox[1], width: viewBox[2], height: viewBox[3] }
    : subpathBounds(subpaths);
  if (!(box.width > 0) || !(box.height > 0)) {
    if (!viewBox && (box.width > 0 || box.height > 0)) {
      // a perfectly horizontal or vertical line has no area: avoid dividing by zero
      box.width = box.width > 0 ? box.width : 1;
      box.height = box.height > 0 ? box.height : 1;
    } else {
      throw new Error('The viewBox / path has no width or height');
    }
  }
  let sx, sy;
  if (width !== undefined && height !== undefined) { sx = width / box.width; sy = height / box.height; }
  else if (width !== undefined) { sx = sy = width / box.width; }
  else if (height !== undefined) { sx = sy = height / box.height; }
  else { sx = sy = 1; }
  const tx = (px) => x + (px - box.x) * sx;
  const ty = (py) => y + (py - box.y) * sy;
  const map = (p) => ({ x: tx(p.x), y: ty(p.y) });
  return subpaths.map((sp) => ({
    closed: sp.closed,
    points: sp.points.map((pt) => ({ ...map(pt), left: map(pt.left), right: map(pt.right) })),
  }));
}

// Regular polygon / star anchor points (optionally with rounded corners). Angles in degrees; the
// first vertex points up, then clockwise (y grows downwards).
export function polygonPoints({ cx, cy, radius, sides, innerRatio, rotation = 0, cornerRadius = 0 }) {
  const n = innerRatio ? sides * 2 : sides;
  const verts = [];
  for (let i = 0; i < n; i++) {
    const r = innerRatio && i % 2 === 1 ? radius * innerRatio : radius;
    const ang = ((rotation - 90) * Math.PI) / 180 + (i * 2 * Math.PI) / n;
    verts.push({ x: cx + r * Math.cos(ang), y: cy + r * Math.sin(ang) });
  }
  if (!(cornerRadius > 0)) {
    return verts.map((v) => ({ x: v.x, y: v.y, left: { x: v.x, y: v.y }, right: { x: v.x, y: v.y } }));
  }
  // Round every corner with a circular fillet (two anchors per corner, cubic approximation of the arc)
  const out = [];
  for (let i = 0; i < n; i++) {
    const p = verts[i], a = verts[(i + n - 1) % n], b = verts[(i + 1) % n];
    const ua = { x: a.x - p.x, y: a.y - p.y }, ub = { x: b.x - p.x, y: b.y - p.y };
    const la = Math.hypot(ua.x, ua.y), lb = Math.hypot(ub.x, ub.y);
    ua.x /= la; ua.y /= la; ub.x /= lb; ub.y /= lb;
    const cosT = Math.max(-1, Math.min(1, ua.x * ub.x + ua.y * ub.y));
    const theta = Math.acos(cosT);                 // interior angle at the corner
    if (theta < 1e-6 || Math.abs(theta - Math.PI) < 1e-6) { out.push({ x: p.x, y: p.y, left: { ...p }, right: { ...p } }); continue; }
    let t = cornerRadius / Math.tan(theta / 2);    // distance from the corner to the tangent points
    t = Math.min(t, 0.5 * la, 0.5 * lb);
    const r = t * Math.tan(theta / 2);             // effective radius after clamping
    const phi = Math.PI - theta;                   // arc sweep
    const k = ((4 / 3) * Math.tan(phi / 4) * r) / t;
    const t1 = { x: p.x + ua.x * t, y: p.y + ua.y * t }, t2 = { x: p.x + ub.x * t, y: p.y + ub.y * t };
    out.push({ x: t1.x, y: t1.y, left: { ...t1 }, right: { x: t1.x + (p.x - t1.x) * k, y: t1.y + (p.y - t1.y) * k } });
    out.push({ x: t2.x, y: t2.y, left: { x: t2.x + (p.x - t2.x) * k, y: t2.y + (p.y - t2.y) * k }, right: { ...t2 } });
  }
  return out;
}
