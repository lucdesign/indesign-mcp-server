// Native vector shape tools: paths from anchor points or SVG path data, line and polygon
// creation, shape conversion, corner options, pathfinder, fill/stroke/gradient editing,
// clipping, duplication, alignment and object inspection.
//
// Coordinates are millimetres, origin at the top-left of the page's trim area (negative values
// reach into the bleed). Objects are addressed by their stable `id`.

import { jsStr, jsJson, jsNum, jsEnum } from '../lib/script-runtime.js';
import { svgPathToSubpaths, fitSubpaths, polygonPoints } from '../lib/svg-path.js';

const has = (v) => v !== undefined && v !== null;

const OBJECT_PROPS = {
  id: { type: 'number', description: 'Stable object id (returned by every create_* tool, place_image and list_page_items)' },
  pageIndex: { type: 'number', description: 'Page index (0-based); only used with index' },
  index: { type: 'number', description: 'Position in list_page_items for that page (shifts when objects are added or reordered)' },
};

const XY = { type: 'object', description: 'Position in mm: { x, y }', properties: { x: { type: 'number' }, y: { type: 'number' } } };

const STYLE_PROPS = {
  pageIndex: { type: 'number', description: 'Page index (0-based)', default: 0 },
  fill: { type: 'string', description: 'Fill swatch name, or "none"' },
  stroke: { type: 'string', description: 'Stroke swatch name, or "none"' },
  strokeWeight: { type: 'number', description: 'Stroke weight in points' },
  opacity: { type: 'number', description: 'Object opacity in percent (0-100)' },
  name: { type: 'string', description: 'Object name (shown in the Layers panel)' },
  label: { type: 'string', description: 'Script label (shown by list_page_items)' },
};

const POINT_TYPES = { corner: 'CORNER', smooth: 'SMOOTH', symmetrical: 'SYMMETRICAL', plain: 'PLAIN', auto: 'auto' };
const ARROWS = {
  none: 'NONE', simple: 'SIMPLE_ARROW_HEAD', 'simple-wide': 'SIMPLE_WIDE_ARROW_HEAD', triangle: 'TRIANGLE_ARROW_HEAD',
  'triangle-wide': 'TRIANGLE_WIDE_ARROW_HEAD', barbed: 'BARBED_ARROW_HEAD', curved: 'CURVED_ARROW_HEAD',
  circle: 'CIRCLE_ARROW_HEAD', 'circle-solid': 'CIRCLE_SOLID_ARROW_HEAD', square: 'SQUARE_ARROW_HEAD',
  'square-solid': 'SQUARE_SOLID_ARROW_HEAD', bar: 'BAR_ARROW_HEAD',
};
const CORNERS = { none: 'NONE', rounded: 'ROUNDED_CORNER', inverse_rounded: 'INVERSE_ROUNDED_CORNER', inset: 'INSET_CORNER', bevel: 'BEVEL_CORNER', fancy: 'FANCY_CORNER' };
const CORNER_NAMES = { top_left: 'topLeft', top_right: 'topRight', bottom_left: 'bottomLeft', bottom_right: 'bottomRight' };

export function createVectorTools(server) {
  const fmt = (result, operation) => server.formatResponse(result, operation);
  const run = (script) => server.executeInDesignScript(script);
  const num = (v, label) => (has(v) ? jsNum(v, label) : 'null');
  const objExpr = (args) => `__resolve(doc, ${num(args.id, 'id')}, ${num(args.pageIndex, 'pageIndex')}, ${num(args.index, 'index')})`;

  // ---- argument helpers -------------------------------------------------------------------
  const xy = (v, label) => {
    const p = Array.isArray(v) ? { x: v[0], y: v[1] } : v;
    if (!p || typeof p !== 'object') throw new Error(`${label} must be { x, y } (mm)`);
    return { x: jsNum(p.x, `${label}.x`), y: jsNum(p.y, `${label}.y`) };
  };

  // One anchor point with optional handles -> { x, y, lx, ly, rx, ry, type }
  const anchor = (pt, label) => {
    const a = xy(pt, label);
    const l = has(pt.leftDirection) ? xy(pt.leftDirection, `${label}.leftDirection`) : a;
    const r = has(pt.rightDirection) ? xy(pt.rightDirection, `${label}.rightDirection`) : a;
    let type = 'auto';
    if (has(pt.pointType)) {
      if (!POINT_TYPES[pt.pointType]) throw new Error(`Invalid pointType at ${label}: ${pt.pointType} (use ${Object.keys(POINT_TYPES).join(', ')})`);
      type = POINT_TYPES[pt.pointType];
    }
    return { x: a.x, y: a.y, lx: l.x, ly: l.y, rx: r.x, ry: r.y, type };
  };

  const fromSvgPoint = (p) => ({ x: p.x, y: p.y, lx: p.left.x, ly: p.left.y, rx: p.right.x, ry: p.right.y, type: 'auto' });

  const styleOpts = (args) => {
    const o = {};
    for (const k of ['fill', 'stroke', 'name', 'label']) if (has(args[k])) o[k] = String(args[k]);
    if (has(args.strokeWeight)) {
      o.strokeWeight = jsNum(args.strokeWeight, 'strokeWeight');
      if (o.strokeWeight < 0) throw new Error('strokeWeight must be >= 0');
    }
    if (has(args.opacity)) {
      o.opacity = jsNum(args.opacity, 'opacity');
      if (o.opacity < 0 || o.opacity > 100) throw new Error('opacity must be 0-100');
    }
    return o;
  };

  // ---- definitions ------------------------------------------------------------------------
  const definitions = [
    {
      name: 'create_path',
      description: 'Create an editable open or closed Bezier path from anchor points. Each point: x, y (mm) and optionally leftDirection / rightDirection (handle positions, same coordinate system) and pointType. Returns the object id and the resulting points. Example (S-curve band): points [{x:0,y:8,rightDirection:{x:40,y:0}}, {x:120,y:8,leftDirection:{x:80,y:14},rightDirection:{x:160,y:2}}, ...], closed true.',
      inputSchema: {
        type: 'object',
        properties: {
          points: {
            type: 'array',
            description: 'At least 2 anchor points',
            items: {
              type: 'object',
              properties: {
                x: { type: 'number', description: 'Anchor x in mm' },
                y: { type: 'number', description: 'Anchor y in mm' },
                leftDirection: { ...XY, description: 'Incoming handle position (default: the anchor, i.e. no handle)' },
                rightDirection: { ...XY, description: 'Outgoing handle position (default: the anchor)' },
                pointType: { type: 'string', enum: Object.keys(POINT_TYPES), description: 'corner (independent handles), smooth (collinear), symmetrical, plain (no handles) or auto' },
              },
              required: ['x', 'y'],
            },
          },
          closed: { type: 'boolean', description: 'Close the path', default: false },
          ...STYLE_PROPS,
        },
        required: ['points'],
      },
    },
    {
      name: 'create_path_from_svg',
      description: 'Create native, editable path(s) from SVG path data (commands M m L l H h V v C c S s Q q T t A a Z z; quadratics and arcs become cubic Beziers). The data is scaled from viewBox (default: the path bounds) into the target rectangle x, y, width, height (mm); one dimension only keeps the aspect ratio. Several sub-paths become one compound path. Example: d "M0 8 C40 0 80 14 120 8 S190 0 216 6 L216 50 L0 50 Z", viewBox "0 0 216 50", x -3, y 55, width 216, height 48.',
      inputSchema: {
        type: 'object',
        properties: {
          d: { type: 'string', description: 'SVG path data' },
          viewBox: { type: 'string', description: 'SVG viewBox "minX minY width height" the data was drawn in (default: bounds of the path)' },
          x: { type: 'number', description: 'Target left edge in mm', default: 0 },
          y: { type: 'number', description: 'Target top edge in mm', default: 0 },
          width: { type: 'number', description: 'Target width in mm' },
          height: { type: 'number', description: 'Target height in mm' },
          ...STYLE_PROPS,
        },
        required: ['d'],
      },
    },
    {
      name: 'edit_path_points',
      description: 'Read or edit the anchor points and handles of an existing path (any rectangle, oval, polygon or line). action: read | add | move | delete | set_type | reverse | set_closed. Coordinates are mm (same system as create_path). Moving an anchor moves its handles with it unless moveHandles is false.',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          action: { type: 'string', enum: ['read', 'add', 'move', 'delete', 'set_type', 'reverse', 'set_closed'], default: 'read' },
          pathIndex: { type: 'number', description: 'Sub-path of a compound path (0-based)', default: 0 },
          pointIndex: { type: 'number', description: 'Point to move / delete / retype; for add: insert before this index (default: append)' },
          point: {
            type: 'object',
            description: 'add: the new point { x, y, leftDirection?, rightDirection?, pointType? }',
            properties: { x: { type: 'number' }, y: { type: 'number' }, leftDirection: XY, rightDirection: XY, pointType: { type: 'string', enum: Object.keys(POINT_TYPES) } },
          },
          x: { type: 'number', description: 'move: new anchor x in mm' },
          y: { type: 'number', description: 'move: new anchor y in mm' },
          leftDirection: { ...XY, description: 'move: new absolute position of the left handle' },
          rightDirection: { ...XY, description: 'move: new absolute position of the right handle' },
          moveHandles: { type: 'boolean', description: 'move: translate the handles together with the anchor', default: true },
          pointType: { type: 'string', enum: Object.keys(POINT_TYPES), description: 'set_type: new point type' },
          closed: { type: 'boolean', description: 'set_closed: true closes, false opens the path' },
        },
      },
    },
    {
      name: 'create_line',
      description: 'Create a straight or curved line between two points (mm). Curve: give controlPoint1 and/or controlPoint2 (Bezier handles). Options: weight (pt), stroke swatch, dash (solid, a stroke style name, or [dash, gap, ...] in pt), endCap, arrowheads (none, simple, simple-wide, triangle, triangle-wide, barbed, curved, circle, circle-solid, square, square-solid, bar).',
      inputSchema: {
        type: 'object',
        properties: {
          x1: { type: 'number', description: 'Start x in mm' },
          y1: { type: 'number', description: 'Start y in mm' },
          x2: { type: 'number', description: 'End x in mm' },
          y2: { type: 'number', description: 'End y in mm' },
          controlPoint1: { ...XY, description: 'Curve: handle leaving the start point' },
          controlPoint2: { ...XY, description: 'Curve: handle arriving at the end point' },
          weight: { type: 'number', description: 'Stroke weight in points' },
          stroke: { type: 'string', description: 'Stroke swatch name' },
          dash: { type: 'array', items: { type: 'number' }, description: 'Custom dash/gap lengths in pt, e.g. [6, 3]. Use dashStyle for a named style or solid.' },
          dashStyle: { type: 'string', description: '"solid" or the name of a stroke style (e.g. Dotted, Dashed)' },
          endCap: { type: 'string', enum: ['butt', 'round', 'projecting'] },
          arrowStart: { type: 'string', enum: Object.keys(ARROWS), description: 'Arrowhead at the start' },
          arrowEnd: { type: 'string', enum: Object.keys(ARROWS), description: 'Arrowhead at the end' },
          opacity: { type: 'number', description: 'Opacity in percent (0-100)' },
          pageIndex: STYLE_PROPS.pageIndex,
          name: STYLE_PROPS.name,
          label: STYLE_PROPS.label,
        },
        required: ['x1', 'y1', 'x2', 'y2'],
      },
    },
    {
      name: 'create_polygon',
      description: 'Create a regular polygon or a star as an editable path. x, y is the centre and radius the outer radius (mm). innerRadiusRatio (0-1) makes a star; cornerRadius (mm) rounds the corners; rotation in degrees (0 = first corner points up).',
      inputSchema: {
        type: 'object',
        properties: {
          x: { type: 'number', description: 'Centre x in mm' },
          y: { type: 'number', description: 'Centre y in mm' },
          radius: { type: 'number', description: 'Outer radius in mm' },
          sides: { type: 'number', description: 'Number of sides (3-100) or star points', default: 5 },
          innerRadiusRatio: { type: 'number', description: 'Star: inner radius / outer radius (0-1)' },
          cornerRadius: { type: 'number', description: 'Round corners with this radius in mm', default: 0 },
          rotation: { type: 'number', description: 'Rotation in degrees', default: 0 },
          ...STYLE_PROPS,
        },
        required: ['x', 'y', 'radius'],
      },
    },
    {
      name: 'convert_shape',
      description: 'Convert a shape with InDesign\'s Convert Shape: to rectangle, rounded_rectangle, beveled_rectangle, inverse_rounded_rectangle, oval, triangle, polygon (sides, insetPercent, cornerRadius), line, straight_line, open_path or closed_path. Every shape (rectangle, oval, polygon) is an editable path: use edit_path_points to change its points.',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          to: { type: 'string', enum: ['rectangle', 'rounded_rectangle', 'beveled_rectangle', 'inverse_rounded_rectangle', 'oval', 'triangle', 'polygon', 'line', 'straight_line', 'open_path', 'closed_path'] },
          sides: { type: 'number', description: 'polygon: number of sides' },
          insetPercent: { type: 'number', description: 'polygon: star inset in percent' },
          cornerRadius: { type: 'number', description: 'polygon / rounded shapes: corner radius in mm' },
        },
        required: ['to'],
      },
    },
    {
      name: 'set_corner_options',
      description: 'Set corner effects and radius (mm): option none | rounded | inverse_rounded | inset | bevel | fancy, for all corners or per corner (corners: { top_left: { option, radius }, ... }).',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          option: { type: 'string', enum: Object.keys(CORNERS), description: 'Effect for all corners (unless corners is given)' },
          radius: { type: 'number', description: 'Radius in mm for all corners' },
          corners: {
            type: 'object',
            description: 'Per corner settings; keys top_left, top_right, bottom_left, bottom_right; each { option?, radius? }',
            properties: Object.fromEntries(Object.keys(CORNER_NAMES).map((k) => [k, { type: 'object', properties: { option: { type: 'string', enum: Object.keys(CORNERS) }, radius: { type: 'number' } } }])),
          },
        },
      },
    },
    {
      name: 'pathfinder',
      description: 'Pathfinder on a list of object ids; returns the id of the result. union / intersect / exclude_overlap combine all shapes. subtract: the FIRST id is kept and all other shapes are cut out of it. minus_back: the frontmost shape is kept and all shapes behind it are cut out. The source objects are consumed.',
      inputSchema: {
        type: 'object',
        properties: {
          operation: { type: 'string', enum: ['union', 'subtract', 'intersect', 'exclude_overlap', 'minus_back'] },
          ids: { type: 'array', items: { type: 'number' }, description: 'Ids of the shapes (at least 2, on the same page)' },
        },
        required: ['operation', 'ids'],
      },
    },
    {
      name: 'set_object_fill',
      description: 'Change the fill of any existing object (rectangle, oval, path, text frame) in place, so the stacking order is kept. swatch (name or "none"), tint (0-100), opacity of the fill (0-100), or a gradient (name of a gradient swatch) with gradientAngle.',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          ids: { type: 'array', items: { type: 'number' }, description: 'Alternative to id: change several objects at once' },
          swatch: { type: 'string', description: 'Fill swatch name, or "none"' },
          tint: { type: 'number', description: 'Fill tint in percent (0-100)' },
          opacity: { type: 'number', description: 'Fill opacity in percent (0-100)' },
          gradient: { type: 'string', description: 'Name of a gradient swatch (see create_gradient_swatch)' },
          gradientAngle: { type: 'number', description: 'Gradient angle in degrees' },
        },
      },
    },
    {
      name: 'set_object_stroke',
      description: 'Change the stroke of any existing object in place: swatch (name or "none"), weight (pt), tint, dash (custom [dash, gap, ...] pt, or dashStyle "solid" / a stroke style name), alignment (center, inside, outside), join (miter, round, bevel), miterLimit, cap (butt, round, projecting), opacity of the stroke (0-100).',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          ids: { type: 'array', items: { type: 'number' }, description: 'Alternative to id: change several objects at once' },
          swatch: { type: 'string', description: 'Stroke swatch name, or "none"' },
          weight: { type: 'number', description: 'Stroke weight in points' },
          tint: { type: 'number', description: 'Stroke tint in percent (0-100)' },
          dash: { type: 'array', items: { type: 'number' }, description: 'Custom dash/gap lengths in pt, e.g. [6, 3]' },
          dashStyle: { type: 'string', description: '"solid" or the name of a stroke style' },
          alignment: { type: 'string', enum: ['center', 'inside', 'outside'] },
          join: { type: 'string', enum: ['miter', 'round', 'bevel'] },
          miterLimit: { type: 'number', description: 'Miter limit (1-500)' },
          cap: { type: 'string', enum: ['butt', 'round', 'projecting'] },
          opacity: { type: 'number', description: 'Stroke opacity in percent (0-100)' },
        },
      },
    },
    {
      name: 'create_gradient_swatch',
      description: 'Create a linear or radial gradient swatch from colour stops. stops: [{ swatch, position (0-100), midpoint? (13-87, between this and the previous stop) }]; the first stop must be at 0 and the last at 100. Apply it with set_object_fill (gradient, gradientAngle).',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Gradient name' },
          type: { type: 'string', enum: ['linear', 'radial'], default: 'linear' },
          stops: {
            type: 'array',
            description: 'At least 2 stops',
            items: { type: 'object', properties: { swatch: { type: 'string' }, position: { type: 'number' }, midpoint: { type: 'number' } }, required: ['swatch', 'position'] },
          },
        },
        required: ['name', 'stops'],
      },
    },
    {
      name: 'set_gradient_feather',
      description: 'Fade an object out with a gradient feather (opacity gradient): type linear/radial, angle, startOpacity (default 100) and endOpacity (default 0). remove: true switches it off.',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          type: { type: 'string', enum: ['linear', 'radial'], default: 'linear' },
          angle: { type: 'number', description: 'Angle in degrees', default: 90 },
          startOpacity: { type: 'number', description: 'Opacity at the start (0-100)', default: 100 },
          endOpacity: { type: 'number', description: 'Opacity at the end (0-100)', default: 0 },
          remove: { type: 'boolean', description: 'Remove the gradient feather', default: false },
        },
      },
    },
    {
      name: 'place_image_in_shape',
      description: 'Place an image into an existing shape (path, rectangle, oval, polygon) which acts as the mask. fitOption FILL_PROPORTIONALLY (crop to fill, default), PROPORTIONALLY, CONTENT_TO_FRAME, CENTER_CONTENT or NONE; then optional contentOffsetX/Y (mm) and contentScale (%). Returns the shape id.',
      inputSchema: {
        type: 'object',
        properties: {
          shapeId: { type: 'number', description: 'Id of the shape that becomes the mask' },
          imagePath: { type: 'string', description: 'Path of the image file' },
          fitOption: { type: 'string', enum: ['FILL_PROPORTIONALLY', 'PROPORTIONALLY', 'CONTENT_TO_FRAME', 'CENTER_CONTENT', 'NONE'], default: 'FILL_PROPORTIONALLY' },
          contentOffsetX: { type: 'number', description: 'Move the image inside the shape (mm)' },
          contentOffsetY: { type: 'number', description: 'Move the image inside the shape (mm)' },
          contentScale: { type: 'number', description: 'Scale of the image content in percent' },
        },
        required: ['shapeId', 'imagePath'],
      },
    },
    {
      name: 'duplicate_object',
      description: 'Duplicate an object (count copies, each moved by offsetX / offsetY mm relative to the previous one, optionally onto another page). Returns the ids of the copies.',
      inputSchema: {
        type: 'object',
        properties: {
          ...OBJECT_PROPS,
          count: { type: 'number', description: 'Number of copies (1-100)', default: 1 },
          offsetX: { type: 'number', description: 'Horizontal offset in mm between copies', default: 0 },
          offsetY: { type: 'number', description: 'Vertical offset in mm between copies', default: 0 },
          toPageIndex: { type: 'number', description: 'Put the copies on this page (0-based)' },
        },
      },
    },
    {
      name: 'align_objects',
      description: 'Align objects by id. alignment: left, right, top, bottom, horizontal_center, vertical_center. relativeTo: selection (bounds of the objects), page, margins, spread, bleed, key_object (with keyObjectId).',
      inputSchema: {
        type: 'object',
        properties: {
          ids: { type: 'array', items: { type: 'number' }, description: 'Object ids' },
          alignment: { type: 'string', enum: ['left', 'right', 'top', 'bottom', 'horizontal_center', 'vertical_center'] },
          relativeTo: { type: 'string', enum: ['selection', 'page', 'margins', 'spread', 'bleed', 'key_object'], default: 'selection' },
          keyObjectId: { type: 'number', description: 'relativeTo key_object: the object that stays where it is' },
        },
        required: ['ids', 'alignment'],
      },
    },
    {
      name: 'distribute_objects',
      description: 'Distribute objects by id. distribution: left_edges, horizontal_centers, right_edges, top_edges, vertical_centers, bottom_edges, horizontal_space, vertical_space. With spacing (mm) the gaps are fixed (horizontal_space / vertical_space).',
      inputSchema: {
        type: 'object',
        properties: {
          ids: { type: 'array', items: { type: 'number' }, description: 'Object ids (at least 3 for a useful result)' },
          distribution: { type: 'string', enum: ['left_edges', 'horizontal_centers', 'right_edges', 'top_edges', 'vertical_centers', 'bottom_edges', 'horizontal_space', 'vertical_space'] },
          spacing: { type: 'number', description: 'Fixed gap in mm (horizontal_space / vertical_space)' },
          relativeTo: { type: 'string', enum: ['selection', 'page', 'margins', 'spread', 'bleed'], default: 'selection' },
        },
        required: ['ids', 'distribution'],
      },
    },
    {
      name: 'get_object_info',
      description: 'Full details of an object: type, page, geometry (mm), rotation, fill (swatch, tint, values), stroke, opacity and blend mode, effects, applied object style, layer, name/label, group membership, and the anchor points of paths.',
      inputSchema: { type: 'object', properties: { ...OBJECT_PROPS } },
    },
  ];

  // ---- handlers ---------------------------------------------------------------------------
  const handlers = {};

  handlers.create_path = async (args) => {
    if (!Array.isArray(args.points) || args.points.length < 2) throw new Error('points must contain at least 2 anchor points');
    const pts = args.points.map((p, i) => anchor(p, `points[${i}]`));
    const closed = args.closed === true;
    const opts = styleOpts(args);
    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${jsNum(has(args.pageIndex) ? args.pageIndex : 0, 'pageIndex')});
      var item = page.polygons.add();
      try {
        __writePath(item.paths[0], ${jsJson(pts)}, ${closed});
        __styleShape(doc, item, ${jsJson(opts)});
      } catch (e) {
        item.remove();
        throw e;
      }
      "Path created: " + __describePath(item) + "\\n" + __fmtPoints(item.paths[0]);
    `;
    return fmt(await run(script), 'Create Path');
  };

  handlers.create_path_from_svg = async (args) => {
    if (typeof args.d !== 'string' || !args.d.trim()) throw new Error('d (SVG path data) is required');
    let viewBox;
    if (has(args.viewBox)) {
      viewBox = String(args.viewBox).trim().split(/[\s,]+/).map(Number);
      if (viewBox.length !== 4 || viewBox.some((n) => !Number.isFinite(n)) || viewBox[2] <= 0 || viewBox[3] <= 0) {
        throw new Error('viewBox must be "minX minY width height" with positive width and height');
      }
    }
    const { subpaths } = svgPathToSubpaths(args.d);
    const fitted = fitSubpaths(subpaths, {
      viewBox,
      x: has(args.x) ? jsNum(args.x, 'x') : 0,
      y: has(args.y) ? jsNum(args.y, 'y') : 0,
      width: has(args.width) ? jsNum(args.width, 'width') : undefined,
      height: has(args.height) ? jsNum(args.height, 'height') : undefined,
    });
    const paths = fitted.map((sp) => ({ closed: sp.closed, points: sp.points.map(fromSvgPoint) }));
    const opts = styleOpts(args);
    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${jsNum(has(args.pageIndex) ? args.pageIndex : 0, 'pageIndex')});
      var data = ${jsJson(paths)};
      var item = page.polygons.add();
      try {
        __writePath(item.paths[0], data[0].points, data[0].closed);
        for (var k = 1; k < data.length; k++) {
          var extra = item.paths.add();
          __writePath(extra, data[k].points, data[k].closed);
        }
        __styleShape(doc, item, ${jsJson(opts)});
      } catch (e) {
        item.remove();
        throw e;
      }
      var listing = [];
      for (var p = 0; p < item.paths.length; p++) listing.push("path " + p + ":\\n" + __fmtPoints(item.paths[p]));
      "Path created from SVG data: " + __describePath(item) + "\\n" + listing.join("\\n");
    `;
    return fmt(await run(script), 'Create Path From SVG');
  };

  handlers.edit_path_points = async (args) => {
    const action = args.action || 'read';
    const pathIndex = has(args.pathIndex) ? jsNum(args.pathIndex, 'pathIndex') : 0;
    const idx = has(args.pointIndex) ? jsNum(args.pointIndex, 'pointIndex') : null;
    let body = '';
    if (action === 'add') {
      if (!args.point) throw new Error('add needs point { x, y, ... }');
      body = `
        var np = ${jsJson(anchor(args.point, 'point'))};
        var at = ${idx === null ? 'pts.length' : idx};
        if (at < 0 || at > pts.length) throw new Error("Invalid pointIndex " + at + " (the path has " + pts.length + " points)");
        pts.splice(at, 0, np);`;
    } else if (action === 'move') {
      if (idx === null) throw new Error('move needs pointIndex');
      if (![args.x, args.y, args.leftDirection, args.rightDirection].some(has)) throw new Error('move needs x/y and/or leftDirection / rightDirection');
      const l = has(args.leftDirection) ? xy(args.leftDirection, 'leftDirection') : null;
      const r = has(args.rightDirection) ? xy(args.rightDirection, 'rightDirection') : null;
      body = `
        var i = ${idx};
        if (i < 0 || i >= pts.length) throw new Error("Invalid pointIndex " + i + " (the path has " + pts.length + " points)");
        var p = pts[i];
        var nx = ${has(args.x) ? jsNum(args.x, 'x') : 'p.x'}, ny = ${has(args.y) ? jsNum(args.y, 'y') : 'p.y'};
        var dx = nx - p.x, dy = ny - p.y;
        ${args.moveHandles === false ? '' : 'p.lx += dx; p.ly += dy; p.rx += dx; p.ry += dy;'}
        p.x = nx; p.y = ny;
        ${l ? `p.lx = ${l.x}; p.ly = ${l.y};` : ''}
        ${r ? `p.rx = ${r.x}; p.ry = ${r.y};` : ''}`;
    } else if (action === 'delete') {
      if (idx === null) throw new Error('delete needs pointIndex');
      body = `
        var i = ${idx};
        if (i < 0 || i >= pts.length) throw new Error("Invalid pointIndex " + i + " (the path has " + pts.length + " points)");
        if (pts.length <= (closed ? 3 : 2)) throw new Error("Cannot delete: a " + (closed ? "closed" : "open") + " path needs at least " + (closed ? 3 : 2) + " points");
        pts.splice(i, 1);`;
    } else if (action === 'set_type') {
      if (idx === null || !has(args.pointType)) throw new Error('set_type needs pointIndex and pointType');
      if (!POINT_TYPES[args.pointType]) throw new Error(`Invalid pointType: ${args.pointType}`);
      body = `
        var i = ${idx};
        if (i < 0 || i >= pts.length) throw new Error("Invalid pointIndex " + i + " (the path has " + pts.length + " points)");
        pts[i].type = ${jsStr(POINT_TYPES[args.pointType])};
        ${POINT_TYPES[args.pointType] === 'PLAIN' ? 'pts[i].lx = pts[i].rx = pts[i].x; pts[i].ly = pts[i].ry = pts[i].y;' : ''}`;
    } else if (action === 'reverse') {
      body = `
        pts.reverse();
        for (var q = 0; q < pts.length; q++) {
          var tx = pts[q].lx, ty = pts[q].ly;
          pts[q].lx = pts[q].rx; pts[q].ly = pts[q].ry; pts[q].rx = tx; pts[q].ry = ty;
        }`;
    } else if (action === 'set_closed') {
      if (typeof args.closed !== 'boolean') throw new Error('set_closed needs closed: true or false');
      body = `closed = ${args.closed};`;
    } else if (action !== 'read') {
      throw new Error(`Invalid action: ${action}`);
    }
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      var kind = item.constructor.name;
      if (kind !== "Rectangle" && kind !== "Oval" && kind !== "Polygon" && kind !== "GraphicLine") throw new Error("Object " + item.id + " is a " + kind + " and has no editable path");
      var pathIndex = ${pathIndex};
      if (pathIndex < 0 || pathIndex >= item.paths.length) throw new Error("Invalid pathIndex " + pathIndex + " (the object has " + item.paths.length + " path(s))");
      var path = item.paths[pathIndex];
      var pts = __readPath(path);
      var closed = __isClosed(path);
      ${body}
      ${action === 'read' ? '' : '__writePath(path, pts, closed);'}
      "${action === 'read' ? 'Points' : 'Path edited (' + action + ')'}: " + __describePath(item) + "\\n" + __fmtPoints(path);
    `;
    return fmt(await run(script), action === 'read' ? 'Read Path Points' : 'Edit Path Points');
  };

  handlers.create_line = async (args) => {
    const p1 = { x: jsNum(args.x1, 'x1'), y: jsNum(args.y1, 'y1') }, p2 = { x: jsNum(args.x2, 'x2'), y: jsNum(args.y2, 'y2') };
    const c1 = has(args.controlPoint1) ? xy(args.controlPoint1, 'controlPoint1') : p1;
    const c2 = has(args.controlPoint2) ? xy(args.controlPoint2, 'controlPoint2') : p2;
    const pts = [
      { x: p1.x, y: p1.y, lx: p1.x, ly: p1.y, rx: c1.x, ry: c1.y, type: has(args.controlPoint1) ? 'CORNER' : 'auto' },
      { x: p2.x, y: p2.y, lx: c2.x, ly: c2.y, rx: p2.x, ry: p2.y, type: has(args.controlPoint2) ? 'CORNER' : 'auto' },
    ];
    for (const k of ['arrowStart', 'arrowEnd']) if (has(args[k]) && !ARROWS[args[k]]) throw new Error(`Invalid ${k}: ${args[k]}`);
    const caps = { butt: 'BUTT_END_CAP', round: 'ROUND_END_CAP', projecting: 'PROJECTING_END_CAP' };
    if (has(args.endCap) && !caps[args.endCap]) throw new Error(`Invalid endCap: ${args.endCap}`);
    const dash = has(args.dash) ? args.dash.map((v, i) => jsNum(v, `dash[${i}]`)) : null;
    const opts = styleOpts({ ...args, fill: undefined });
    if (has(args.weight)) opts.strokeWeight = jsNum(args.weight, 'weight');
    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${jsNum(has(args.pageIndex) ? args.pageIndex : 0, 'pageIndex')});
      var item = page.graphicLines.add();
      try {
        __writePath(item.paths[0], ${jsJson(pts)}, false);
        __styleShape(doc, item, ${jsJson(opts)});
        ${has(args.endCap) ? `item.endCap = EndCap.${caps[args.endCap]};` : ''}
        ${has(args.arrowStart) ? `item.leftLineEnd = ArrowHead.${ARROWS[args.arrowStart]};` : ''}
        ${has(args.arrowEnd) ? `item.rightLineEnd = ArrowHead.${ARROWS[args.arrowEnd]};` : ''}
        ${dash ? `item.strokeType = __strokeStyle(doc, "dashed"); item.strokeDashAndGap = ${jsJson(dash)};` : ''}
        ${has(args.dashStyle) ? `item.strokeType = __strokeStyle(doc, ${jsStr(args.dashStyle)});` : ''}
      } catch (e) {
        item.remove();
        throw e;
      }
      "Line created: " + __describePath(item) + "\\n" + __fmtPoints(item.paths[0]);
    `;
    return fmt(await run(script), 'Create Line');
  };

  handlers.create_polygon = async (args) => {
    const sides = Math.floor(has(args.sides) ? jsNum(args.sides, 'sides') : 5);
    if (sides < 3 || sides > 100) throw new Error('sides must be 3-100');
    const radius = jsNum(args.radius, 'radius');
    if (!(radius > 0)) throw new Error('radius must be > 0');
    let innerRatio;
    if (has(args.innerRadiusRatio)) {
      innerRatio = jsNum(args.innerRadiusRatio, 'innerRadiusRatio');
      if (!(innerRatio > 0 && innerRatio < 1)) throw new Error('innerRadiusRatio must be between 0 and 1');
    }
    const cornerRadius = has(args.cornerRadius) ? jsNum(args.cornerRadius, 'cornerRadius') : 0;
    if (cornerRadius < 0) throw new Error('cornerRadius must be >= 0');
    const pts = polygonPoints({
      cx: jsNum(args.x, 'x'), cy: jsNum(args.y, 'y'), radius, sides, innerRatio,
      rotation: has(args.rotation) ? jsNum(args.rotation, 'rotation') : 0, cornerRadius,
    }).map(fromSvgPoint);
    const opts = styleOpts(args);
    const script = `
      var doc = __requireDoc();
      var page = __page(doc, ${jsNum(has(args.pageIndex) ? args.pageIndex : 0, 'pageIndex')});
      var item = page.polygons.add();
      try {
        __writePath(item.paths[0], ${jsJson(pts)}, true);
        __styleShape(doc, item, ${jsJson(opts)});
      } catch (e) {
        item.remove();
        throw e;
      }
      "${innerRatio ? 'Star' : 'Polygon'} created: " + __describePath(item);
    `;
    return fmt(await run(script), innerRatio ? 'Create Star' : 'Create Polygon');
  };

  handlers.convert_shape = async (args) => {
    const map = {
      rectangle: 'CONVERT_TO_RECTANGLE', rounded_rectangle: 'CONVERT_TO_ROUNDED_RECTANGLE', beveled_rectangle: 'CONVERT_TO_BEVELED_RECTANGLE',
      inverse_rounded_rectangle: 'CONVERT_TO_INVERSE_ROUNDED_RECTANGLE', oval: 'CONVERT_TO_OVAL', triangle: 'CONVERT_TO_TRIANGLE',
      polygon: 'CONVERT_TO_POLYGON', line: 'CONVERT_TO_LINE', straight_line: 'CONVERT_TO_STRAIGHT_LINE', open_path: 'CONVERT_TO_OPEN_PATH', closed_path: 'CONVERT_TO_CLOSED_PATH',
    };
    if (!map[args.to]) throw new Error(`Invalid target shape: ${args.to}`);
    const extra = [];
    if (args.to === 'polygon' && has(args.sides)) {
      extra.push(Math.floor(jsNum(args.sides, 'sides')));
      extra.push(has(args.insetPercent) ? jsNum(args.insetPercent, 'insetPercent') : 0);
      extra.push(has(args.cornerRadius) ? jsNum(args.cornerRadius, 'cornerRadius') : 0);
    } else if (has(args.cornerRadius)) {
      extra.push(6, 0, jsNum(args.cornerRadius, 'cornerRadius'));
    }
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      var before = __pathCount(item) + " points (" + item.constructor.name + ")";
      item.convertShape(ConvertShapeOptions.${map[args.to]}${extra.length ? ', ' + extra.join(', ') : ''});
      item = __findById(doc, item.id) || item;
      "Converted to ${args.to}: was " + before + ", now " + __describePath(item);
    `;
    return fmt(await run(script), 'Convert Shape');
  };

  handlers.set_corner_options = async (args) => {
    const settings = {};   // corner -> { option, radius }
    const put = (corner, option, radius) => {
      const s = (settings[corner] = settings[corner] || {});
      if (has(option)) { if (!CORNERS[option]) throw new Error(`Invalid corner option: ${option}`); s.option = CORNERS[option]; }
      if (has(radius)) { s.radius = jsNum(radius, 'radius'); if (s.radius < 0) throw new Error('radius must be >= 0'); }
    };
    if (has(args.corners)) {
      for (const [k, v] of Object.entries(args.corners)) {
        if (!CORNER_NAMES[k]) throw new Error(`Unknown corner: ${k} (use ${Object.keys(CORNER_NAMES).join(', ')})`);
        put(CORNER_NAMES[k], v?.option, v?.radius);
      }
    } else {
      for (const c of Object.values(CORNER_NAMES)) put(c, args.option, args.radius);
    }
    if (Object.values(settings).every((s) => s.option === undefined && s.radius === undefined)) throw new Error('Pass option and/or radius (or corners)');
    const lines = Object.entries(settings).map(([c, s]) => `
        ${s.option ? `item.${c}CornerOption = CornerOptions.${s.option};` : ''}
        ${has(s.radius) ? `item.${c}CornerRadius = ${s.radius};` : ''}`).join('');
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      ${lines}
      var report = [];
      var names = ["topLeft", "topRight", "bottomLeft", "bottomRight"];
      for (var i = 0; i < names.length; i++) report.push(names[i] + "=" + String(item[names[i] + "CornerOption"]) + "/" + __r(item[names[i] + "CornerRadius"]) + "mm");
      "Corner options set: " + __describe(item) + "\\n" + report.join(" ");
    `;
    return fmt(await run(script), 'Set Corner Options');
  };

  handlers.pathfinder = async (args) => {
    const ops = { union: 'addPath', intersect: 'intersectPath', exclude_overlap: 'excludeOverlapPath', subtract: 'subtractPath', minus_back: 'subtractPath' };
    if (!ops[args.operation]) throw new Error(`Invalid operation: ${args.operation}`);
    if (!Array.isArray(args.ids) || args.ids.length < 2) throw new Error('ids must list at least 2 objects');
    const ids = args.ids.map((v, i) => jsNum(v, `ids[${i}]`));
    const op = args.operation;
    const script = `
      var doc = __requireDoc();
      var ids = ${jsJson(ids)};
      var items = [];
      for (var i = 0; i < ids.length; i++) {
        var it = __findById(doc, ids[i]);
        if (!it) throw new Error("No object with id " + ids[i]);
        var kind = it.constructor.name;
        if (kind !== "Rectangle" && kind !== "Oval" && kind !== "Polygon" && kind !== "GraphicLine") throw new Error("Object " + ids[i] + " is a " + kind + " and cannot be used in a pathfinder operation");
        items.push(it);
      }
      var result;
      try {
        ${op === 'union' || op === 'intersect' || op === 'exclude_overlap' ? `
          result = items[0];
          for (var k = 1; k < items.length; k++) result = result.${ops[op]}(items[k]);
        ` : `
          var base = items[0];
          ${op === 'minus_back' ? `
            // the frontmost shape (lowest index in the stacking order of its page) is the one that stays
            var stack = items[0].parentPage ? items[0].parentPage.allPageItems : null;
            if (!stack) throw new Error("minus_back needs objects that are placed on a page");
            var bestIndex = 1e9;
            for (var m = 0; m < items.length; m++) {
              for (var s = 0; s < stack.length; s++) { if (stack[s].id === items[m].id) { if (s < bestIndex) { bestIndex = s; base = items[m]; } break; } }
            }
          ` : ''}
          result = base;
          for (var n = 0; n < items.length; n++) {
            if (items[n].id === base.id) continue;
            // X.subtractPath(Y) leaves Y minus X, so the cutter is the receiver
            result = items[n].subtractPath(result);
          }
        `}
      } catch (e) {
        if (e.number === 11276) throw new Error("The pathfinder result is empty: the shapes do not overlap the way ${op} needs");
        throw e;
      }
      "Pathfinder ${op} done. Result: " + __describePath(result);
    `;
    return fmt(await run(script), `Pathfinder ${op}`);
  };

  const targets = (args) => {
    if (Array.isArray(args.ids) && args.ids.length) return { list: args.ids.map((v, i) => jsNum(v, `ids[${i}]`)) };
    return { single: true };
  };
  const eachTarget = (args, inner) => {
    const t = targets(args);
    return `
      var doc = __requireDoc();
      var targets = [];
      ${t.single ? `targets.push(${objExpr(args)});` : `
        var tid = ${jsJson(t.list)};
        for (var ti = 0; ti < tid.length; ti++) {
          var found = __findById(doc, tid[ti]);
          if (!found) throw new Error("No object with id " + tid[ti]);
          targets.push(found);
        }`}
      var report = [];
      for (var ti2 = 0; ti2 < targets.length; ti2++) {
        var item = targets[ti2];
        ${inner}
        report.push(__describe(item) + " fill=" + (function () { try { return item.fillColor.name + "/" + item.fillTint; } catch (e) { return "n/a"; } })() + " stroke=" + (function () { try { return item.strokeColor.name + "/" + __r(item.strokeWeight) + "pt"; } catch (e) { return "n/a"; } })());
      }
      report.join("\\n");`;
  };
  const pct = (v, label) => { const n = jsNum(v, label); if (n < 0 || n > 100) throw new Error(`${label} must be 0-100`); return n; };

  handlers.set_object_fill = async (args) => {
    if (![args.swatch, args.tint, args.opacity, args.gradient, args.gradientAngle].some(has)) throw new Error('Pass swatch, tint, opacity or gradient');
    if (has(args.swatch) && has(args.gradient)) throw new Error('Pass either swatch or gradient, not both');
    const inner = `
        ${has(args.swatch) ? `item.fillColor = __swatchOrNone(doc, ${jsStr(args.swatch)});` : ''}
        ${has(args.gradient) ? `
          var grad = doc.gradients.itemByName(${jsStr(args.gradient)});
          if (!grad.isValid) throw new Error("Gradient not found: " + ${jsStr(args.gradient)});
          item.fillColor = grad;` : ''}
        ${has(args.gradientAngle) ? `item.gradientFillAngle = ${jsNum(args.gradientAngle, 'gradientAngle')};` : ''}
        ${has(args.tint) ? `item.fillTint = ${pct(args.tint, 'tint')};` : ''}
        ${has(args.opacity) ? `item.fillTransparencySettings.blendingSettings.opacity = ${pct(args.opacity, 'opacity')};` : ''}`;
    return fmt(await run(eachTarget(args, inner)), 'Set Object Fill');
  };

  handlers.set_object_stroke = async (args) => {
    if (![args.swatch, args.weight, args.tint, args.dash, args.dashStyle, args.alignment, args.join, args.miterLimit, args.cap, args.opacity].some(has)) {
      throw new Error('Pass at least one stroke property');
    }
    const align = { center: 'CENTER_ALIGNMENT', inside: 'INSIDE_ALIGNMENT', outside: 'OUTSIDE_ALIGNMENT' };
    const join = { miter: 'MITER_END_JOIN', round: 'ROUND_END_JOIN', bevel: 'BEVEL_END_JOIN' };
    const caps = { butt: 'BUTT_END_CAP', round: 'ROUND_END_CAP', projecting: 'PROJECTING_END_CAP' };
    if (has(args.alignment) && !align[args.alignment]) throw new Error(`Invalid alignment: ${args.alignment}`);
    if (has(args.join) && !join[args.join]) throw new Error(`Invalid join: ${args.join}`);
    if (has(args.cap) && !caps[args.cap]) throw new Error(`Invalid cap: ${args.cap}`);
    const dash = has(args.dash) ? args.dash.map((v, i) => jsNum(v, `dash[${i}]`)) : null;
    const inner = `
        ${has(args.swatch) ? `item.strokeColor = __swatchOrNone(doc, ${jsStr(args.swatch)});` : ''}
        ${has(args.weight) ? `item.strokeWeight = ${jsNum(args.weight, 'weight')};` : ''}
        ${has(args.tint) ? `item.strokeTint = ${pct(args.tint, 'tint')};` : ''}
        ${has(args.dashStyle) ? `item.strokeType = __strokeStyle(doc, ${jsStr(args.dashStyle)});` : ''}
        ${dash ? `item.strokeType = __strokeStyle(doc, "dashed"); item.strokeDashAndGap = ${jsJson(dash)};` : ''}
        ${has(args.alignment) ? `item.strokeAlignment = StrokeAlignment.${align[args.alignment]};` : ''}
        ${has(args.join) ? `item.endJoin = OutlineJoin.${join[args.join]};` : ''}
        ${has(args.miterLimit) ? `item.miterLimit = ${jsNum(args.miterLimit, 'miterLimit')};` : ''}
        ${has(args.cap) ? `item.endCap = EndCap.${caps[args.cap]};` : ''}
        ${has(args.opacity) ? `item.strokeTransparencySettings.blendingSettings.opacity = ${pct(args.opacity, 'opacity')};` : ''}`;
    return fmt(await run(eachTarget(args, inner)), 'Set Object Stroke');
  };

  handlers.create_gradient_swatch = async (args) => {
    const type = args.type || 'linear';
    if (!['linear', 'radial'].includes(type)) throw new Error(`Invalid type: ${type}`);
    if (!Array.isArray(args.stops) || args.stops.length < 2) throw new Error('stops must contain at least 2 colour stops');
    const stops = args.stops.map((s, i) => ({
      swatch: String(s.swatch),
      position: pct(s.position, `stops[${i}].position`),
      midpoint: has(s.midpoint) ? jsNum(s.midpoint, `stops[${i}].midpoint`) : null,
    })).sort((a, b) => a.position - b.position);
    if (stops[0].position !== 0 || stops[stops.length - 1].position !== 100) throw new Error('The first stop must be at position 0 and the last at position 100');
    for (let i = 1; i < stops.length; i++) if (stops[i].position <= stops[i - 1].position) throw new Error('Stop positions must be strictly increasing');
    const script = `
      var doc = __requireDoc();
      if (doc.gradients.itemByName(${jsStr(args.name)}).isValid) throw new Error("Gradient already exists: " + ${jsStr(args.name)});
      var stops = ${jsJson(stops)};
      var colours = [];
      for (var c = 0; c < stops.length; c++) colours.push(__swatch(doc, stops[c].swatch));
      var g = doc.gradients.add({ name: ${jsStr(args.name)}, type: GradientType.${type === 'radial' ? 'RADIAL' : 'LINEAR'} });
      try {
        // a new gradient has two stops (0 and 100); the others are inserted in between
        g.gradientStops[0].stopColor = colours[0];
        g.gradientStops[1].stopColor = colours[colours.length - 1];
        for (var k = 1; k < stops.length - 1; k++) g.gradientStops.add({ stopColor: colours[k], location: stops[k].position });
        for (var m = 1; m < stops.length; m++) if (stops[m].midpoint !== null) g.gradientStops[m].midpoint = stops[m].midpoint;
      } catch (e) {
        g.remove();
        throw e;
      }
      var listing = [];
      for (var i = 0; i < g.gradientStops.length; i++) {
        var st = g.gradientStops[i], mid = "";
        try { mid = " midpoint=" + __r(st.midpoint); } catch (e) {}
        listing.push("  " + __r(st.location) + "% " + st.stopColor.name + mid);
      }
      "Gradient " + ${jsStr(args.name)} + " created (" + String(g.type).toLowerCase() + ", " + g.gradientStops.length + " stops):\\n" + listing.join("\\n");
    `;
    return fmt(await run(script), 'Create Gradient Swatch');
  };

  handlers.set_gradient_feather = async (args) => {
    const type = args.type || 'linear';
    if (!['linear', 'radial'].includes(type)) throw new Error(`Invalid type: ${type}`);
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      var f = item.transparencySettings.gradientFeatherSettings;
      ${args.remove ? `
        f.applied = false;
        "Gradient feather removed: " + __describe(item);
      ` : `
        f.applied = true;
        f.type = GradientType.${type === 'radial' ? 'RADIAL' : 'LINEAR'};
        f.angle = ${has(args.angle) ? jsNum(args.angle, 'angle') : 90};
        f.opacityGradientStops[0].opacity = ${has(args.startOpacity) ? pct(args.startOpacity, 'startOpacity') : 100};
        f.opacityGradientStops[f.opacityGradientStops.length - 1].opacity = ${has(args.endOpacity) ? pct(args.endOpacity, 'endOpacity') : 0};
        "Gradient feather set (" + String(f.type).toLowerCase() + ", angle " + __r(f.angle) + "): " + __describe(item);
      `}
    `;
    return fmt(await run(script), 'Set Gradient Feather');
  };

  handlers.place_image_in_shape = async (args) => {
    const validated = server.validateFilePath(args.imagePath);
    const fit = has(args.fitOption) ? args.fitOption : 'FILL_PROPORTIONALLY';
    const fits = { FILL_PROPORTIONALLY: 'FitOptions.FILL_PROPORTIONALLY', PROPORTIONALLY: 'FitOptions.PROPORTIONALLY', CONTENT_TO_FRAME: 'FitOptions.CONTENT_TO_FRAME', CENTER_CONTENT: 'FitOptions.CENTER_CONTENT', NONE: null };
    if (!(fit in fits)) throw new Error(`Invalid fitOption: ${fit}`);
    const script = `
      var doc = __requireDoc();
      var shape = __resolve(doc, ${jsNum(args.shapeId, 'shapeId')}, null, null);
      var file = File(${jsStr(validated)});
      if (!file.exists) throw new Error("Image file not found: " + file.fsName);
      var kind = shape.constructor.name;
      if (kind !== "Rectangle" && kind !== "Oval" && kind !== "Polygon") throw new Error("Object " + shape.id + " is a " + kind + " and cannot be used as a mask (use a rectangle, oval or polygon/path)");
      shape.place(file);
      ${fits[fit] ? `shape.fit(${fits[fit]});` : ''}
      ${has(args.contentScale) ? `shape.graphics[0].horizontalScale = ${jsNum(args.contentScale, 'contentScale')}; shape.graphics[0].verticalScale = ${jsNum(args.contentScale, 'contentScale')};` : ''}
      ${has(args.contentOffsetX) || has(args.contentOffsetY) ? `shape.graphics[0].move(undefined, [${has(args.contentOffsetX) ? jsNum(args.contentOffsetX, 'contentOffsetX') : 0}, ${has(args.contentOffsetY) ? jsNum(args.contentOffsetY, 'contentOffsetY') : 0}]);` : ''}
      "Image placed in shape: " + __describePath(shape);
    `;
    return fmt(await run(script), 'Place Image In Shape');
  };

  handlers.duplicate_object = async (args) => {
    const count = Math.floor(has(args.count) ? jsNum(args.count, 'count') : 1);
    if (count < 1 || count > 100) throw new Error('count must be 1-100');
    const ox = has(args.offsetX) ? jsNum(args.offsetX, 'offsetX') : 0, oy = has(args.offsetY) ? jsNum(args.offsetY, 'offsetY') : 0;
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      ${has(args.toPageIndex) ? `var target = __page(doc, ${jsNum(args.toPageIndex, 'toPageIndex')});` : 'var target = null;'}
      var ids = [], last = item, lines = [];
      for (var i = 0; i < ${count}; i++) {
        var copy = target && i === 0 ? last.duplicate(target) : last.duplicate();
        if (${ox} !== 0 || ${oy} !== 0) copy.move(undefined, [${ox}, ${oy}]);
        ids.push(copy.id);
        lines.push(__describe(copy));
        last = copy;
      }
      "Duplicated " + ids.length + " time(s), ids: " + ids.join(", ") + "\\n" + lines.join("\\n");
    `;
    return fmt(await run(script), 'Duplicate Object');
  };

  const boundsMap = { selection: 'ITEM_BOUNDS', page: 'PAGE_BOUNDS', margins: 'MARGIN_BOUNDS', spread: 'SPREAD_BOUNDS', bleed: 'BLEED_BOUNDS', key_object: 'KEY_OBJECT' };
  const idList = (ids, min) => {
    if (!Array.isArray(ids) || ids.length < min) throw new Error(`ids must list at least ${min} object(s)`);
    return ids.map((v, i) => jsNum(v, `ids[${i}]`));
  };
  const itemsFromIds = `
      var items = [];
      for (var i = 0; i < ids.length; i++) {
        var it = __findById(doc, ids[i]);
        if (!it) throw new Error("No object with id " + ids[i]);
        items.push(it);
      }`;

  handlers.align_objects = async (args) => {
    const opts = { left: 'LEFT_EDGES', right: 'RIGHT_EDGES', top: 'TOP_EDGES', bottom: 'BOTTOM_EDGES', horizontal_center: 'HORIZONTAL_CENTERS', vertical_center: 'VERTICAL_CENTERS' };
    if (!opts[args.alignment]) throw new Error(`Invalid alignment: ${args.alignment}`);
    const rel = args.relativeTo || 'selection';
    if (!boundsMap[rel]) throw new Error(`Invalid relativeTo: ${rel}`);
    const ids = idList(args.ids, rel === 'selection' || rel === 'key_object' ? 2 : 1);
    if (rel === 'key_object' && !has(args.keyObjectId)) throw new Error('relativeTo key_object needs keyObjectId');
    const script = `
      var doc = __requireDoc();
      var ids = ${jsJson(ids)};
      ${itemsFromIds}
      ${rel === 'key_object' ? `
        var key = __findById(doc, ${jsNum(args.keyObjectId, 'keyObjectId')});
        if (!key) throw new Error("No object with id ${jsNum(args.keyObjectId, 'keyObjectId')}");
        var found = false;
        for (var f = 0; f < items.length; f++) if (items[f].id === key.id) found = true;
        if (!found) items.push(key);
        doc.align(items, AlignOptions.${opts[args.alignment]}, AlignDistributeBounds.KEY_OBJECT, key);
      ` : `doc.align(items, AlignOptions.${opts[args.alignment]}, AlignDistributeBounds.${boundsMap[rel]});`}
      var lines = [];
      for (var j = 0; j < items.length; j++) lines.push(__describe(items[j]));
      "Aligned (${args.alignment}, relative to ${rel}):\\n" + lines.join("\\n");
    `;
    return fmt(await run(script), 'Align Objects');
  };

  handlers.distribute_objects = async (args) => {
    const opts = { left_edges: 'LEFT_EDGES', horizontal_centers: 'HORIZONTAL_CENTERS', right_edges: 'RIGHT_EDGES', top_edges: 'TOP_EDGES', vertical_centers: 'VERTICAL_CENTERS', bottom_edges: 'BOTTOM_EDGES', horizontal_space: 'HORIZONTAL_SPACE', vertical_space: 'VERTICAL_SPACE' };
    if (!opts[args.distribution]) throw new Error(`Invalid distribution: ${args.distribution}`);
    const rel = args.relativeTo || 'selection';
    if (!boundsMap[rel] || rel === 'key_object') throw new Error(`Invalid relativeTo: ${rel}`);
    const ids = idList(args.ids, 2);
    const spacing = has(args.spacing) ? jsNum(args.spacing, 'spacing') : null;
    if (spacing !== null && !/_space$/.test(args.distribution)) throw new Error('spacing only applies to horizontal_space / vertical_space');
    const script = `
      var doc = __requireDoc();
      var ids = ${jsJson(ids)};
      ${itemsFromIds}
      doc.distribute(items, DistributeOptions.${opts[args.distribution]}, AlignDistributeBounds.${boundsMap[rel]}${spacing !== null ? `, true, ${spacing}` : ''});
      var lines = [];
      for (var j = 0; j < items.length; j++) lines.push(__describe(items[j]));
      "Distributed (${args.distribution}):\\n" + lines.join("\\n");
    `;
    return fmt(await run(script), 'Distribute Objects');
  };

  handlers.get_object_info = async (args) => {
    const script = `
      var doc = __requireDoc();
      var item = ${objExpr(args)};
      function on(o) { try { return o.applied === true; } catch (e) { return false; } }
      var L = [];
      L.push("=== OBJECT " + item.id + " (" + item.constructor.name + ") ===");
      L.push("geometry: " + __describe(item));
      try { L.push("name=" + JSON_STR(item.name) + " label=" + JSON_STR(item.label) + " layer=" + item.itemLayer.name + " locked=" + item.locked + " visible=" + item.visible); } catch (e) {}
      try { L.push("rotation=" + __r(item.rotationAngle) + " shear=" + __r(item.shearAngle)); } catch (e) {}
      try {
        var fc = item.fillColor, fd = fc.name;
        try { if (fc.constructor.name === "Swatch" || fc.constructor.name === "Color") { var e1 = fc.getElements()[0]; if (e1.colorValue) fd += " [" + String(e1.space) + " " + (function (v) { var a = []; for (var i = 0; i < v.length; i++) a.push(__r(v[i])); return a.join(","); })(e1.colorValue) + "]"; } } catch (e) {}
        L.push("fill: " + fd + " tint=" + __r(item.fillTint) + (fc.constructor.name === "Gradient" ? " angle=" + __r(item.gradientFillAngle) : ""));
      } catch (e) { L.push("fill: n/a"); }
      try {
        var sc = item.strokeColor;
        var dash = "";
        try { dash = " type=" + item.strokeType.name; } catch (e) {}
        L.push("stroke: " + sc.name + " weight=" + __r(item.strokeWeight) + "pt tint=" + __r(item.strokeTint) + " align=" + String(item.strokeAlignment) + " join=" + String(item.endJoin) + " cap=" + String(item.endCap) + dash);
      } catch (e) { L.push("stroke: n/a"); }
      try {
        var ts = item.transparencySettings;
        L.push("opacity=" + __r(ts.blendingSettings.opacity) + " blend=" + String(ts.blendingSettings.blendMode) + " fillOpacity=" + __r(item.fillTransparencySettings.blendingSettings.opacity) + " strokeOpacity=" + __r(item.strokeTransparencySettings.blendingSettings.opacity));
        var fx = [];
        try { if (String(ts.dropShadowSettings.mode) !== "NONE") fx.push("dropShadow"); } catch (e) {}
        try { if (on(ts.outerGlowSettings)) fx.push("outerGlow"); } catch (e) {}
        try { if (on(ts.innerGlowSettings)) fx.push("innerGlow"); } catch (e) {}
        try { if (on(ts.bevelAndEmbossSettings)) fx.push("bevelEmboss"); } catch (e) {}
        try { if (on(ts.satinSettings)) fx.push("satin"); } catch (e) {}
        try { if (on(ts.featherSettings)) fx.push("feather"); } catch (e) {}
        try { if (on(ts.directionalFeatherSettings)) fx.push("directionalFeather"); } catch (e) {}
        try { if (on(item.transparencySettings.gradientFeatherSettings)) fx.push("gradientFeather"); } catch (e) {}
        L.push("effects: " + (fx.length ? fx.join(", ") : "none"));
      } catch (e) {}
      try { L.push("objectStyle=" + item.appliedObjectStyle.name); } catch (e) {}
      try { if (item.parent && item.parent.constructor.name === "Group") L.push("group=" + item.parent.id + " (members: " + item.parent.pageItems.length + ")"); else L.push("group=none"); } catch (e) {}
      try { if (item.constructor.name === "Group") { var ids = []; for (var g = 0; g < item.pageItems.length; g++) ids.push(item.pageItems[g].id); L.push("members: " + ids.join(", ")); } } catch (e) {}
      try {
        if (item.allGraphics && item.allGraphics.length > 0) {
          var gr = item.allGraphics[0], ppi = "";
          try { ppi = ", effective ppi " + __r(gr.effectivePpi[0]); } catch (e) {}   // vector graphics have none
          L.push("image: " + gr.itemLink.name + " (" + String(gr.itemLink.status) + ", " + gr.imageTypeName + ppi + ")");
        }
      } catch (e) {}
      try { if (item.constructor.name === "TextFrame") L.push("text: " + __overflowNote(item) + " paragraphs=" + item.parentStory.paragraphs.length); } catch (e) {}
      try {
        if (item.paths && item.paths.length > 0) {
          L.push("paths=" + item.paths.length + " points=" + __pathCount(item));
          for (var p = 0; p < item.paths.length; p++) L.push("path " + p + " (" + (__isClosed(item.paths[p]) ? "closed" : "open") + "):\\n" + __fmtPoints(item.paths[p]));
        }
      } catch (e) {}
      L.join("\\n");
    `;
    return fmt(await run(script), 'Object Info');
  };

  return { definitions, handlers };
}
