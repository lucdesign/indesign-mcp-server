// Swatches and fonts: update / delete swatches, list swatches with values and usage, list fonts.

import { jsStr, jsJson, jsNum } from '../lib/script-runtime.js';

const has = (v) => v !== undefined && v !== null;

export function createResourceTools(server) {
  const fmt = (result, operation) => server.formatResponse(result, operation);
  const run = (script) => server.executeInDesignScript(script);

  const definitions = [
    {
      name: 'update_color_swatch',
      description: 'Change an existing colour swatch in place (all objects using it follow): new values (colorValues CMYK 0-100 or RGB 0-255, or hex), spot/process, and/or a new name. RGB or hex values are converted to CMYK in print documents with the colour management of InDesign unless keepRgb is true. Built-in swatches (Black, Paper, Registration, None) cannot be changed.',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Name of the swatch to change' },
          newName: { type: 'string', description: 'Rename the swatch' },
          colorModel: { type: 'string', enum: ['CMYK', 'RGB'], description: 'Model of colorValues', default: 'CMYK' },
          colorValues: { type: 'array', items: { type: 'number' }, description: '[C,M,Y,K] 0-100 or [R,G,B] 0-255' },
          hex: { type: 'string', description: 'Alternative to colorValues: RGB hex such as #FF6600' },
          keepRgb: { type: 'boolean', description: 'Keep RGB values as RGB even in a print document', default: false },
          conversion: { type: 'string', enum: ['profile', 'simple'], description: 'RGB/hex to CMYK: profile = InDesign colour management (accurate, default); simple = formula without profile', default: 'profile' },
          spotColor: { type: 'boolean', description: 'true = spot colour, false = process colour' },
        },
        required: ['name'],
      },
    },
    {
      name: 'delete_color_swatch',
      description: 'Delete a swatch. If objects or styles use it you must say what replaces it (replaceWith: a swatch name, or "none"); the error lists how it is used.',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Name of the swatch to delete' },
          replaceWith: { type: 'string', description: 'Swatch that takes its place where it is used, or "none"' },
        },
        required: ['name'],
      },
    },
    {
      name: 'list_color_swatches',
      description: 'List all swatches with kind, model and values (colours), base colour and tint (tints), stops (gradients) and where each is used: number of fills, strokes and styles, plus the ids of up to 10 objects. Local text colour overrides inside stories are not counted.',
      inputSchema: {
        type: 'object',
        properties: {
          includeUsage: { type: 'boolean', description: 'Also count where each swatch is used (scans the whole document)', default: true },
        },
      },
    },
    {
      name: 'list_fonts',
      description: 'List installed font families with their styles (use them for fontFamily / fontStyle). Filter with search (substring of the family name) or family (exact name); limit caps the number of families shown.',
      inputSchema: {
        type: 'object',
        properties: {
          search: { type: 'string', description: 'Case-insensitive substring of the family name' },
          family: { type: 'string', description: 'Exact family name: list only its styles' },
          limit: { type: 'number', description: 'Maximum number of families (1-1000)', default: 60 },
        },
      },
    },
  ];

  const handlers = {};

  handlers.update_color_swatch = async (args) => {
    let model = args.colorModel || 'CMYK';
    let values = args.colorValues;
    if (has(args.hex)) { model = 'RGB'; values = server.parseHex(args.hex); }
    if (!['CMYK', 'RGB'].includes(model)) throw new Error(`Invalid colorModel: ${model}`);
    if (has(values)) {
      if (!Array.isArray(values) || values.length !== (model === 'CMYK' ? 4 : 3) || values.some((v) => !Number.isFinite(Number(v)))) {
        throw new Error(`colorValues must be ${model === 'CMYK' ? '[C,M,Y,K] (0-100)' : '[R,G,B] (0-255)'}`);
      }
      values = values.map(Number);
    }
    if (has(args.conversion) && !['profile', 'simple'].includes(args.conversion)) throw new Error(`Invalid conversion: ${args.conversion}`);
    if (!has(values) && !has(args.newName) && !has(args.spotColor)) throw new Error('Nothing to change: pass colorValues / hex, newName and/or spotColor');
    const rgbToCmyk = has(values) && model === 'RGB' ? server.rgbToCmyk(...values) : [0, 0, 0, 0];

    const script = `
      var doc = __requireDoc();
      var color = doc.colors.itemByName(${jsStr(args.name)});
      if (!color.isValid) throw new Error("Swatch not found: " + ${jsStr(args.name)} + " (only colour swatches can be updated; gradients and tints are separate)");
      var usage = __usageMap(doc)[color.name] || { fills: 0, strokes: 0, styles: 0 };
      var note = "";
      try {
        ${has(values) ? `
          var model = ${jsStr(model)}, values = ${jsJson(values)};
          if (model === "RGB" && ${args.keepRgb ? 'false' : 'true'} && String(doc.documentPreferences.intent).indexOf("PRINT") === 0) {
            ${args.conversion === 'simple'
              ? `values = ${jsJson(rgbToCmyk)}; note = " (RGB converted to CMYK with the simple formula, no colour profile)";`
              : `values = __cmsCmyk(doc, values); note = " (RGB converted to CMYK with the document profile " + doc.cmykProfile + ")";`}
            model = "CMYK";
          }
          color.space = (model === "CMYK") ? ColorSpace.CMYK : ColorSpace.RGB;
          color.colorValue = values;
        ` : ''}
        ${has(args.spotColor) ? `color.model = ${args.spotColor ? 'ColorModel.SPOT' : 'ColorModel.PROCESS'};` : ''}
        ${has(args.newName) ? `
          if (doc.swatches.itemByName(${jsStr(args.newName)}).isValid) throw new Error("Swatch already exists: " + ${jsStr(args.newName)});
          color.name = ${jsStr(args.newName)};
        ` : ''}
      } catch (e) {
        if (e.number && e.number !== 1) throw new Error("The swatch " + ${jsStr(args.name)} + " cannot be modified (built-in swatches such as Black, Paper, Registration and None are read-only). InDesign error " + e.number);
        throw e;
      }
      var out = [];
      for (var i = 0; i < color.colorValue.length; i++) out.push(__r(color.colorValue[i]));
      "Swatch " + color.name + " updated: " + String(color.space) + " " + (String(color.model) === "SPOT" ? "spot" : "process") + " " + out.join(",") + note +
        ". Used by " + usage.fills + " fill(s), " + usage.strokes + " stroke(s), " + usage.styles + " style(s).";
    `;
    return fmt(await run(script), 'Update Color Swatch');
  };

  handlers.delete_color_swatch = async (args) => {
    const script = `
      var doc = __requireDoc();
      var sw = doc.swatches.itemByName(${jsStr(args.name)});
      if (!sw.isValid) throw new Error("Swatch not found: " + ${jsStr(args.name)});
      var u = __usageMap(doc)[sw.name] || { fills: 0, strokes: 0, styles: 0, ids: [] };
      var used = u.fills + u.strokes + u.styles;
      ${has(args.replaceWith) ? `
        var rep = __swatchOrNone(doc, ${jsStr(args.replaceWith)});
        if (rep.name === sw.name) throw new Error("replaceWith must be a different swatch");
        try { sw.remove(rep); } catch (e) { throw new Error("The swatch " + ${jsStr(args.name)} + " cannot be deleted (built-in swatches are read-only). InDesign error " + e.number); }
      ` : `
        if (used > 0) throw new Error("Swatch " + ${jsStr(args.name)} + " is used by " + u.fills + " fill(s), " + u.strokes + " stroke(s) and " + u.styles + " style(s) (object ids: " + u.ids.join(", ") + "). Pass replaceWith (a swatch name, or 'none').");
        try { sw.remove(); } catch (e) { throw new Error("The swatch " + ${jsStr(args.name)} + " cannot be deleted (built-in swatches are read-only). InDesign error " + e.number); }
      `}
      "Swatch " + ${jsStr(args.name)} + " deleted" + (used > 0 ? " (" + used + " use(s) replaced)" : "") + ".";
    `;
    return fmt(await run(script), 'Delete Color Swatch');
  };

  handlers.list_color_swatches = async (args) => {
    const usage = args.includeUsage !== false;
    const script = `
      var doc = __requireDoc();
      var usage = ${usage ? '__usageMap(doc)' : '{}'};
      var out = ["=== SWATCHES (" + doc.swatches.length + ") ==="];
      function vals(v) { var a = []; for (var i = 0; i < v.length; i++) a.push(__r(v[i])); return a.join(", "); }
      for (var i = 0; i < doc.swatches.length; i++) {
        var e = doc.swatches[i].getElements()[0];
        var kind = e.constructor.name, d = "";
        try {
          if (kind === "Color") {
            d = String(e.space) + " " + (String(e.model) === "SPOT" ? "spot" : String(e.model).toLowerCase()) + " [" + vals(e.colorValue) + "]";
          } else if (kind === "Tint") {
            d = e.tintValue + "% of " + e.baseColor.name;
          } else if (kind === "Gradient") {
            var stops = [];
            for (var g = 0; g < e.gradientStops.length; g++) stops.push(__r(e.gradientStops[g].location) + "% " + e.gradientStops[g].stopColor.name);
            d = String(e.type).toLowerCase() + ": " + stops.join(" | ");
          } else if (kind === "MixedInk" || kind === "MixedInkGroup") {
            d = "mixed ink";
          }
        } catch (err) {}
        var line = "\\u2022 " + e.name + "  (" + (kind === "Swatch" ? "built-in" : kind) + (d ? ", " + d : "") + ")";
        ${usage ? `
          var u = usage[e.name];
          line += "  used: " + (u ? u.fills + " fill(s), " + u.strokes + " stroke(s), " + u.styles + " style(s)" + (u.ids.length ? ", ids " + u.ids.join(",") : "") : "not used");` : ''}
        out.push(line);
      }
      out.join("\\n");
    `;
    return fmt(await run(script), 'List Color Swatches');
  };

  handlers.list_fonts = async (args) => {
    const limit = Math.floor(has(args.limit) ? jsNum(args.limit, 'limit') : 60);
    if (limit < 1 || limit > 1000) throw new Error('limit must be 1-1000');
    const script = `
      var all = __fontFamilies();
      var search = ${has(args.search) ? jsStr(String(args.search).toLowerCase()) : 'null'};
      var exact = ${has(args.family) ? jsStr(args.family) : 'null'};
      var names = [];
      for (var i = 0; i < all.order.length; i++) {
        var f = all.order[i];
        if (exact !== null && f !== exact) continue;
        if (search !== null && f.toLowerCase().indexOf(search) === -1) continue;
        names.push(f);
      }
      if (exact !== null && names.length === 0) throw new Error(__fontHelp(exact, ""));
      var out = ["=== FONTS: " + names.length + " of " + all.order.length + " installed families ==="];
      for (var k = 0; k < names.length && k < ${limit}; k++) out.push(names[k] + ": " + all.map[names[k]].join(", "));
      if (names.length > ${limit}) out.push("... " + (names.length - ${limit}) + " more (narrow down with search, or raise limit)");
      out.push("Use fontFamily = the family and fontStyle = one of its styles.");
      out.join("\\n");
    `;
    return fmt(await run(script), 'List Fonts');
  };

  return { definitions, handlers };
}
