/**
 * XlsxFormat — OOXML (.xlsx / .xlsm / .xltx) formatting reader for ExcelTab.
 *
 * The bundled SheetJS community build only exposes cell values, number formats and a
 * partial fill object, so this module reads the workbook package directly (through the
 * zip reader SheetJS ships as XLSX.CFB) and resolves everything needed to reproduce the
 * original appearance: theme/indexed/RGB colors with tints, fonts, fills, borders,
 * alignment, rich-text runs, column widths, row heights, hidden rows/columns/sheets,
 * default sheet formatting, gridlines, freeze panes, tab colors and page setup.
 *
 * Usage: const fmt = XlsxFormat.parse(arrayBuffer);  // null when not an OOXML package
 *        fmt.sheets[sheetName] -> sheet formatting model (see parseSheet)
 */
(function (global) {
  'use strict';

  const PX_PER_PT = 96 / 72;
  const MAX_DIGIT_WIDTH = 7; // Calibri 11 / Aptos Narrow 11 max digit width in px

  // Office 2013-2022 default theme, used when the workbook has no theme part.
  // Order follows the theme index used by cell colors: lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
  const DEFAULT_THEME = [
    '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31',
    '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'
  ];

  const DEFAULT_INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
  ].map(h => '#' + h);

  // Share of the foreground color shown by each fill pattern (approximated as a flat blend).
  const PATTERN_COVERAGE = {
    darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625,
    darkHorizontal: 0.5, darkVertical: 0.5, darkDown: 0.5, darkUp: 0.5, darkGrid: 0.5, darkTrellis: 0.75,
    lightHorizontal: 0.25, lightVertical: 0.25, lightDown: 0.25, lightUp: 0.25, lightGrid: 0.25, lightTrellis: 0.25
  };

  const BORDER_STYLES = {
    thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed',
    dashDot: '1px dashed', dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed',
    mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed',
    thick: '3px solid', double: '3px double'
  };

  const FORMAT_COLORS = {
    black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
    magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00'
  };

  const SERIF_FONTS = /^(times|georgia|cambria|garamond|book antiqua|palatino|century|constantia|bookman|baskerville|caladea|mincho|ms mincho|batang|simsun)/i;
  const MONO_FONTS = /^(courier|consolas|lucida console|monaco|menlo|andale mono|cascadia|ms gothic)/i;
  const FONT_SUBSTITUTES = { calibri: 'Carlito', cambria: 'Caladea', aptos: 'Calibri', 'aptos narrow': 'Calibri' };

  /* ---------------------------------------------------------------- XML helpers */

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) return null;
    return doc;
  }

  function kids(el, name) {
    const out = [];
    if (!el) return out;
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (n.localName === name) out.push(n);
    }
    return out;
  }

  function kid(el, name) {
    if (!el) return null;
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (n.localName === name) return n;
    }
    return null;
  }

  function attr(el, name) {
    if (!el) return null;
    if (el.hasAttribute(name)) return el.getAttribute(name);
    const attrs = el.attributes;
    for (let i = 0; i < attrs.length; i++) {
      if (attrs[i].localName === name) return attrs[i].value;
    }
    return null;
  }

  function isTrue(v) {
    return v === '1' || v === 'true';
  }

  // Boolean child element such as <b/>, <b val="1"/> or <b val="0"/>
  function flag(el, name) {
    const k = kid(el, name);
    if (!k) return false;
    const v = attr(k, 'val');
    return v === null || isTrue(v);
  }

  function num(v, fallback) {
    if (v === null || v === undefined || v === '') return fallback;
    const n = parseFloat(v);
    return isNaN(n) ? fallback : n;
  }

  function parseAttrs(str) {
    const out = {};
    const re = /([\w:]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let m;
    while ((m = re.exec(str))) {
      const name = m[1].indexOf(':') >= 0 ? m[1].slice(m[1].indexOf(':') + 1) : m[1];
      out[name] = m[2] !== undefined ? m[2] : m[3];
    }
    return out;
  }

  function decodeCellRef(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref || '');
    if (!m) return null;
    let c = 0;
    const letters = m[1].toUpperCase();
    for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
  }

  /* ---------------------------------------------------------------- Zip package */

  function openPackage(data) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4B) return null;
    if (!global.XLSX || !global.XLSX.CFB) return null;
    const zip = global.XLSX.CFB.read(bytes, { type: 'array' });
    const decoder = new TextDecoder('utf-8');

    return {
      text(path) {
        if (!path) return null;
        const clean = '/' + path.replace(/^\/+/, '');
        const entry = global.XLSX.CFB.find(zip, clean);
        if (!entry || !entry.content || !entry.content.length) return null;
        const content = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content);
        return decoder.decode(content);
      }
    };
  }

  function resolveTarget(baseDir, target) {
    if (!target) return null;
    if (target.charAt(0) === '/') return target.slice(1);
    const parts = (baseDir + target).split('/');
    const out = [];
    parts.forEach(p => {
      if (p === '..') out.pop();
      else if (p && p !== '.') out.push(p);
    });
    return out.join('/');
  }

  function readRels(pkg, path) {
    const rels = {};
    const text = pkg.text(path);
    if (!text) return rels;
    const doc = parseXml(text);
    if (!doc) return rels;
    kids(doc.documentElement, 'Relationship').forEach(rel => {
      rels[attr(rel, 'Id')] = { type: attr(rel, 'Type') || '', target: attr(rel, 'Target') };
    });
    return rels;
  }

  /* ---------------------------------------------------------------- Colors */

  function hexToRgb(hex) {
    const n = parseInt(hex.replace('#', ''), 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }

  function rgbToHex(r, g, b) {
    const h = v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
    return ('#' + h(r) + h(g) + h(b)).toUpperCase();
  }

  // Excel tint: shift HSL luminance toward white (tint > 0) or black (tint < 0).
  function applyTint(hex, tint) {
    if (!tint) return hex;
    let [r, g, b] = hexToRgb(hex).map(v => v / 255);
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    if (s === 0) {
      r = g = b = l;
    } else {
      const hue2rgb = (p, q, t) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
      };
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;
      r = hue2rgb(p, q, h + 1 / 3);
      g = hue2rgb(p, q, h);
      b = hue2rgb(p, q, h - 1 / 3);
    }
    return rgbToHex(r * 255, g * 255, b * 255);
  }

  function blend(fg, bg, amount) {
    const a = hexToRgb(fg), b = hexToRgb(bg);
    return rgbToHex(a[0] * amount + b[0] * (1 - amount), a[1] * amount + b[1] * (1 - amount), a[2] * amount + b[2] * (1 - amount));
  }

  // Resolves a <color>/<fgColor>/<bgColor>-style element to '#RRGGBB', or null for automatic.
  function makeColorResolver(themeColors, indexedColors) {
    return function resolve(el, autoColor) {
      if (!el) return null;
      let hex = null;
      const rgb = attr(el, 'rgb');
      const theme = attr(el, 'theme');
      const indexed = attr(el, 'indexed');
      if (rgb && /^[0-9a-f]{6,8}$/i.test(rgb)) {
        hex = '#' + rgb.slice(-6).toUpperCase(); // Excel ignores the alpha byte
      } else if (theme !== null) {
        hex = themeColors[parseInt(theme, 10)] || null;
      } else if (indexed !== null) {
        const i = parseInt(indexed, 10);
        if (i === 64) hex = autoColor || null;           // system foreground
        else if (i === 65) hex = null;                   // system background
        else hex = indexedColors[i] || null;
      } else if (isTrue(attr(el, 'auto'))) {
        hex = autoColor || null;
      }
      if (hex) {
        const tint = num(attr(el, 'tint'), 0);
        if (tint) hex = applyTint(hex, tint);
      }
      return hex;
    };
  }

  function parseTheme(text) {
    const colors = DEFAULT_THEME.slice();
    const doc = text && parseXml(text);
    if (!doc) return colors;
    const scheme = doc.getElementsByTagNameNS('*', 'clrScheme')[0];
    if (!scheme) return colors;
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    order.forEach((name, i) => {
      const node = kid(scheme, name);
      if (!node) return;
      const srgb = kid(node, 'srgbClr');
      const sys = kid(node, 'sysClr');
      const val = srgb ? attr(srgb, 'val') : (sys ? (attr(sys, 'lastClr') || (attr(sys, 'val') === 'window' ? 'FFFFFF' : '000000')) : null);
      if (val && /^[0-9a-f]{6}$/i.test(val)) colors[i] = '#' + val.toUpperCase();
    });
    return colors;
  }

  /* ---------------------------------------------------------------- Styles */

  function cleanFontName(name) {
    return String(name || '').replace(/["'\\;<>{}&\u0000-\u001f]/g, '').trim();
  }

  function fontStack(name) {
    const clean = cleanFontName(name);
    if (!clean) return null;
    const stack = [`'${clean}'`];
    const sub = FONT_SUBSTITUTES[clean.toLowerCase()];
    if (sub) stack.push(`'${sub}'`);
    stack.push(SERIF_FONTS.test(clean) ? 'serif' : MONO_FONTS.test(clean) ? 'monospace' : 'sans-serif');
    return stack.join(', ');
  }

  function parseFontEl(el, resolveColor) {
    const f = {};
    const name = kid(el, 'name') || kid(el, 'rFont');
    if (name) f.name = attr(name, 'val');
    const sz = kid(el, 'sz');
    if (sz) f.sz = num(attr(sz, 'val'), undefined);
    if (flag(el, 'b')) f.b = true;
    if (flag(el, 'i')) f.i = true;
    if (flag(el, 'strike')) f.strike = true;
    const u = kid(el, 'u');
    if (u) {
      const val = attr(u, 'val') || 'single';
      if (val !== 'none') f.u = val;
    }
    const va = kid(el, 'vertAlign');
    if (va) f.vertAlign = attr(va, 'val');
    const color = kid(el, 'color');
    if (color) {
      const c = resolveColor(color, '#000000');
      if (c) f.color = c;
    }
    return f;
  }

  function fontCss(f, base) {
    const css = [];
    if (f.name && (!base || f.name !== base.name)) {
      const stack = fontStack(f.name);
      if (stack) css.push(`font-family:${stack};`);
    }
    if (f.sz && (!base || f.sz !== base.sz)) css.push(`font-size:${f.sz}pt;`);
    if (f.b) css.push('font-weight:bold;');
    if (f.i) css.push('font-style:italic;');
    const deco = [];
    if (f.u) deco.push('underline');
    if (f.strike) deco.push('line-through');
    if (deco.length) {
      css.push(`text-decoration:${deco.join(' ')};`);
      if (f.u === 'double' || f.u === 'doubleAccounting') css.push('text-decoration-style:double;');
    }
    if (f.color && (!base || f.color !== base.color)) css.push(`color:${f.color};`);
    return css.join('');
  }

  function parseFill(el, resolveColor) {
    const pattern = kid(el, 'patternFill');
    if (pattern) {
      const type = attr(pattern, 'patternType') || (kid(pattern, 'fgColor') ? 'solid' : 'none');
      if (type === 'none') return null;
      const fg = resolveColor(kid(pattern, 'fgColor'), '#000000');
      const bg = resolveColor(kid(pattern, 'bgColor'), '#FFFFFF');
      let color;
      if (type === 'solid') {
        color = fg || bg;
      } else {
        color = blend(fg || '#000000', bg || '#FFFFFF', PATTERN_COVERAGE[type] !== undefined ? PATTERN_COVERAGE[type] : 0.5);
      }
      return color ? { color, css: `background-color:${color};` } : null;
    }

    const gradient = kid(el, 'gradientFill');
    if (gradient) {
      const stops = kids(gradient, 'stop').map(s => ({
        pos: num(attr(s, 'position'), 0),
        color: resolveColor(kid(s, 'color'), '#FFFFFF') || '#FFFFFF'
      }));
      if (!stops.length) return null;
      const list = stops.map(s => `${s.color} ${Math.round(s.pos * 100)}%`).join(', ');
      const image = attr(gradient, 'type') === 'path'
        ? `radial-gradient(circle, ${list})`
        : `linear-gradient(${90 + num(attr(gradient, 'degree'), 0)}deg, ${list})`;
      return { color: stops[0].color, css: `background-color:${stops[0].color};background-image:${image};` };
    }
    return null;
  }

  function parseBorder(el, resolveColor) {
    const side = names => {
      for (const n of names) {
        const s = kid(el, n);
        const style = s && attr(s, 'style');
        if (style && style !== 'none' && BORDER_STYLES[style]) {
          return `${BORDER_STYLES[style]} ${resolveColor(kid(s, 'color'), '#000000') || '#000000'}`;
        }
      }
      return null;
    };
    return {
      left: side(['left', 'start']),
      right: side(['right', 'end']),
      top: side(['top']),
      bottom: side(['bottom'])
    };
  }

  function parseAlignment(el) {
    const a = {};
    if (!el) return a;
    const h = attr(el, 'horizontal');
    if (h) a.h = h;
    const v = attr(el, 'vertical');
    if (v) a.v = v;
    if (isTrue(attr(el, 'wrapText'))) a.wrap = true;
    if (isTrue(attr(el, 'shrinkToFit'))) a.shrink = true;
    const indent = num(attr(el, 'indent'), 0);
    if (indent) a.indent = indent;
    const rot = num(attr(el, 'textRotation'), 0);
    if (rot) a.rot = rot;
    return a;
  }

  const H_ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify', distributed: 'justify', centerContinuous: 'center', fill: 'left' };
  const V_ALIGN = { top: 'top', center: 'middle', bottom: 'bottom', justify: 'middle', distributed: 'middle' };

  function alignCss(a) {
    const css = [];
    const h = H_ALIGN[a.h];
    if (h) css.push(`text-align:${h};`);
    if (a.v && V_ALIGN[a.v]) css.push(`vertical-align:${V_ALIGN[a.v]};`);
    if (a.wrap || a.h === 'justify' || a.h === 'distributed' || a.v === 'justify' || a.v === 'distributed') {
      css.push('white-space:normal;overflow-wrap:break-word;');
    }
    if (a.indent) {
      const pad = 3 + Math.round(a.indent * 9);
      css.push(a.h === 'right' ? `padding-right:${pad}px;` : `padding-left:${pad}px;`);
    }
    return css.join('');
  }

  function parseStyles(text, themeColors) {
    const doc = text && parseXml(text);
    const root = doc && doc.documentElement;

    let indexed = DEFAULT_INDEXED;
    const custom = root && kid(kid(root, 'colors'), 'indexedColors');
    if (custom) {
      indexed = DEFAULT_INDEXED.slice();
      kids(custom, 'rgbColor').forEach((c, i) => {
        const rgb = attr(c, 'rgb');
        if (rgb && /^[0-9a-f]{6,8}$/i.test(rgb)) indexed[i] = '#' + rgb.slice(-6).toUpperCase();
      });
    }
    const resolveColor = makeColorResolver(themeColors, indexed);

    const fonts = root ? kids(kid(root, 'fonts'), 'font').map(f => parseFontEl(f, resolveColor)) : [];
    const fills = root ? kids(kid(root, 'fills'), 'fill').map(f => parseFill(f, resolveColor)) : [];
    const borders = root ? kids(kid(root, 'borders'), 'border').map(b => parseBorder(b, resolveColor)) : [];

    const baseFont = Object.assign({ name: 'Calibri', sz: 11, color: '#000000' }, fonts[0] || {});
    const noBorder = { left: null, right: null, top: null, bottom: null };

    const xfEls = root ? kids(kid(root, 'cellXfs'), 'xf') : [];
    const xfs = xfEls.map(xf => {
      const font = fonts[num(attr(xf, 'fontId'), 0)] || baseFont;
      const fill = fills[num(attr(xf, 'fillId'), 0)] || null;
      const border = borders[num(attr(xf, 'borderId'), 0)] || noBorder;
      const align = parseAlignment(kid(xf, 'alignment'));
      return {
        font,
        fill,
        border,
        align,
        css: fontCss(font, baseFont) + alignCss(align)
      };
    });
    if (!xfs.length) xfs.push({ font: baseFont, fill: null, border: noBorder, align: {}, css: '' });

    return { xfs, baseFont, resolveColor, indexed };
  }

  /* ---------------------------------------------------------------- Rich text */

  function parseRuns(container, resolveColor) {
    const runEls = kids(container, 'r');
    if (!runEls.length) return null;
    return runEls.map(r => {
      const t = kids(r, 't').map(n => n.textContent).join('');
      const rPr = kid(r, 'rPr');
      return { text: t, font: rPr ? parseFontEl(rPr, resolveColor) : null };
    });
  }

  function parseSharedStrings(text, resolveColor) {
    const out = [];
    const doc = text && parseXml(text);
    if (!doc) return out;
    kids(doc.documentElement, 'si').forEach(si => out.push(parseRuns(si, resolveColor)));
    return out;
  }

  /* ---------------------------------------------------------------- Worksheets */

  function colWidthToPx(width) {
    return Math.max(0, Math.trunc(((256 * width + Math.trunc(128 / MAX_DIGIT_WIDTH)) / 256) * MAX_DIGIT_WIDTH));
  }

  // Splits a worksheet into its (potentially huge) <sheetData> body and the rest of the XML.
  function splitSheetData(text) {
    const open = /<((?:\w+:)?sheetData)\b[^>]*?(\/?)>/.exec(text);
    if (!open) return { rest: text, data: '' };
    const start = open.index;
    if (open[2] === '/') {
      return { rest: text.slice(0, start) + text.slice(start + open[0].length), data: '' };
    }
    const closeTag = '</' + open[1] + '>';
    const end = text.indexOf(closeTag, start);
    if (end < 0) return { rest: text, data: '' };
    return {
      rest: text.slice(0, start) + text.slice(end + closeTag.length),
      data: text.slice(start + open[0].length, end)
    };
  }

  function parseSheet(text, ctx) {
    const { rest, data } = splitSheetData(text);
    const doc = parseXml(rest);
    const root = doc ? doc.documentElement : null;

    const sheet = {
      xfs: ctx.styles.xfs,
      baseFont: ctx.styles.baseFont,
      indexed: ctx.styles.indexed,
      rows: new Map(),         // r -> { px, hidden, s }
      cells: new Map(),         // r -> Map(c -> xf index)
      rich: new Map(),          // "r,c" -> runs
      cols: [],                 // [{ min, max, px, hidden, style }]
      defaultRowPx: Math.round(15 * PX_PER_PT),
      defaultColPx: 64,
      zeroHeight: false,
      showGrid: true,
      frozenRows: 0,
      frozenCols: 0,
      tabColor: null,
      margins: null,
      pageSetup: {},
      printOptions: {}
    };

    if (root) {
      const sheetPr = kid(root, 'sheetPr');
      const tab = kid(sheetPr, 'tabColor');
      if (tab) sheet.tabColor = ctx.styles.resolveColor(tab, null);
      const setUpPr = kid(sheetPr, 'pageSetUpPr');
      if (setUpPr && isTrue(attr(setUpPr, 'fitToPage'))) sheet.pageSetup.fitToPage = true;

      const fmt = kid(root, 'sheetFormatPr');
      if (fmt) {
        const defaultRowHeight = num(attr(fmt, 'defaultRowHeight'), 15);
        sheet.defaultRowPx = Math.round(defaultRowHeight * PX_PER_PT);
        sheet.zeroHeight = isTrue(attr(fmt, 'zeroHeight'));
        const defaultColWidth = num(attr(fmt, 'defaultColWidth'), null);
        if (defaultColWidth !== null) {
          sheet.defaultColPx = colWidthToPx(defaultColWidth);
        } else {
          const base = num(attr(fmt, 'baseColWidth'), 8);
          sheet.defaultColPx = Math.ceil((base * MAX_DIGIT_WIDTH + 5) / 8) * 8;
        }
      }

      kids(kid(root, 'cols'), 'col').forEach(col => {
        const width = num(attr(col, 'width'), null);
        const px = width === null ? null : colWidthToPx(width);
        const style = attr(col, 'style');
        sheet.cols.push({
          min: num(attr(col, 'min'), 1) - 1,
          max: num(attr(col, 'max'), 1) - 1,
          px,
          hidden: isTrue(attr(col, 'hidden')) || px === 0,
          style: style !== null ? parseInt(style, 10) : null
        });
      });

      const views = kids(kid(root, 'sheetViews'), 'sheetView');
      const view = views.find(v => isTrue(attr(v, 'tabSelected'))) || views[0];
      if (view) {
        if (attr(view, 'showGridLines') === '0' || attr(view, 'showGridLines') === 'false') sheet.showGrid = false;
        const pane = kid(view, 'pane');
        const state = pane && attr(pane, 'state');
        if (pane && (state === 'frozen' || state === 'frozenSplit')) {
          const origin = decodeCellRef(attr(view, 'topLeftCell')) || { r: 0, c: 0 };
          const ySplit = Math.round(num(attr(pane, 'ySplit'), 0));
          const xSplit = Math.round(num(attr(pane, 'xSplit'), 0));
          if (ySplit > 0) sheet.frozenRows = origin.r + ySplit;
          if (xSplit > 0) sheet.frozenCols = origin.c + xSplit;
        }
      }

      const margins = kid(root, 'pageMargins');
      if (margins) {
        sheet.margins = {
          left: num(attr(margins, 'left'), 0.7),
          right: num(attr(margins, 'right'), 0.7),
          top: num(attr(margins, 'top'), 0.75),
          bottom: num(attr(margins, 'bottom'), 0.75),
          header: num(attr(margins, 'header'), 0.3),
          footer: num(attr(margins, 'footer'), 0.3)
        };
      }

      const setup = kid(root, 'pageSetup');
      if (setup) {
        Object.assign(sheet.pageSetup, {
          orientation: attr(setup, 'orientation') || null,
          paperSize: num(attr(setup, 'paperSize'), 1),
          scale: num(attr(setup, 'scale'), 100),
          fitToWidth: num(attr(setup, 'fitToWidth'), 1),
          fitToHeight: num(attr(setup, 'fitToHeight'), 1)
        });
      }

      const print = kid(root, 'printOptions');
      if (print) {
        sheet.printOptions = {
          gridLines: isTrue(attr(print, 'gridLines')),
          headings: isTrue(attr(print, 'headings')),
          horizontalCentered: isTrue(attr(print, 'horizontalCentered')),
          verticalCentered: isTrue(attr(print, 'verticalCentered'))
        };
      }
    }

    parseSheetData(data, sheet, ctx);
    return sheet;
  }

  // Scans <sheetData> with regular expressions: far faster and lighter than a DOM for large sheets.
  function parseSheetData(data, sheet, ctx) {
    if (!data) return;
    const rowRe = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
    const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let rowMatch;
    let nextRow = 0;

    while ((rowMatch = rowRe.exec(data))) {
      const ra = parseAttrs(rowMatch[1]);
      const r = ra.r ? parseInt(ra.r, 10) - 1 : nextRow;
      nextRow = r + 1;

      const ht = num(ra.ht, null);
      const info = {
        px: ht === null ? null : Math.round(ht * PX_PER_PT),
        hidden: isTrue(ra.hidden) || ht === 0,
        custom: isTrue(ra.customHeight),
        s: isTrue(ra.customFormat) && ra.s !== undefined ? parseInt(ra.s, 10) : null
      };
      sheet.rows.set(r, info);

      const body = rowMatch[2];
      if (!body) continue;
      let rowCells = null;
      let nextCol = 0;
      let cellMatch;
      cellRe.lastIndex = 0;
      while ((cellMatch = cellRe.exec(body))) {
        const ca = parseAttrs(cellMatch[1]);
        const ref = ca.r ? decodeCellRef(ca.r) : null;
        const c = ref ? ref.c : nextCol;
        nextCol = c + 1;
        if (ca.s !== undefined) {
          if (!rowCells) {
            rowCells = new Map();
            sheet.cells.set(r, rowCells);
          }
          rowCells.set(c, parseInt(ca.s, 10) || 0);
        }

        const inner = cellMatch[2];
        if (!inner) continue;
        if (ca.t === 's') {
          const v = /<(?:\w+:)?v>\s*(\d+)\s*<\//.exec(inner);
          const runs = v && ctx.sharedStrings[parseInt(v[1], 10)];
          if (runs) sheet.rich.set(r + ',' + c, runs);
        } else if (ca.t === 'inlineStr' && /<(?:\w+:)?r>/.test(inner)) {
          const doc = parseXml('<root xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + inner.replace(/<(\/?)\w+:/g, '<$1') + '</root>');
          const is = doc && kid(doc.documentElement, 'is');
          const runs = is && parseRuns(is, ctx.styles.resolveColor);
          if (runs) sheet.rich.set(r + ',' + c, runs);
        }
      }
    }
  }

  /* ---------------------------------------------------------------- Workbook */

  function parse(data) {
    const pkg = openPackage(data);
    if (!pkg) return null;

    // Locate the workbook part through the package relationships.
    const rootRels = readRels(pkg, '_rels/.rels');
    let wbPath = 'xl/workbook.xml';
    Object.keys(rootRels).forEach(id => {
      if (/\/officeDocument$/.test(rootRels[id].type)) wbPath = resolveTarget('', rootRels[id].target);
    });
    const wbText = pkg.text(wbPath);
    const wbDoc = wbText && parseXml(wbText);
    if (!wbDoc) return null;

    const wbDir = wbPath.indexOf('/') >= 0 ? wbPath.slice(0, wbPath.lastIndexOf('/') + 1) : '';
    const wbRelsPath = wbDir + '_rels/' + wbPath.slice(wbDir.length) + '.rels';
    const rels = readRels(pkg, wbRelsPath);
    const partOfType = suffix => {
      const id = Object.keys(rels).find(k => rels[k].type.endsWith(suffix));
      return id ? resolveTarget(wbDir, rels[id].target) : null;
    };

    const themeColors = parseTheme(pkg.text(partOfType('/theme')));
    const styles = parseStyles(pkg.text(partOfType('/styles')), themeColors);
    const sharedStrings = parseSharedStrings(pkg.text(partOfType('/sharedStrings')), styles.resolveColor);
    const ctx = { styles, sharedStrings };

    const root = wbDoc.documentElement;
    const view = kid(kid(root, 'bookViews'), 'workbookView');
    const result = {
      baseFont: styles.baseFont,
      activeTab: view ? num(attr(view, 'activeTab'), 0) : 0,
      sheetOrder: [],
      hiddenSheets: {},
      sheets: {}
    };

    kids(kid(root, 'sheets'), 'sheet').forEach(el => {
      const name = attr(el, 'name');
      const state = attr(el, 'state');
      result.sheetOrder.push(name);
      if (state === 'hidden' || state === 'veryHidden') result.hiddenSheets[name] = true;
      const rel = rels[attr(el, 'id')];
      if (!rel || !/\/worksheet$/.test(rel.type)) return; // chartsheets / dialogsheets
      try {
        const text = pkg.text(resolveTarget(wbDir, rel.target));
        if (text) result.sheets[name] = parseSheet(text, ctx);
      } catch (e) {
        console.warn('XlsxFormat: could not read formatting for sheet "' + name + '"', e);
      }
    });

    return result;
  }

  /* ---------------------------------------------------------------- Sheet helpers */

  function colInfo(sheet, c) {
    for (let i = sheet.cols.length - 1; i >= 0; i--) {
      const col = sheet.cols[i];
      if (c >= col.min && c <= col.max) return col;
    }
    return null;
  }

  // Effective cell format: the cell's own xf, else the row style, else the column style.
  function styleAt(sheet, r, c) {
    const rowCells = sheet.cells.get(r);
    if (rowCells && rowCells.has(c)) return sheet.xfs[rowCells.get(c)] || sheet.xfs[0];
    const row = sheet.rows.get(r);
    if (row && row.s !== null && sheet.xfs[row.s]) return sheet.xfs[row.s];
    const col = colInfo(sheet, c);
    if (col && col.style !== null && sheet.xfs[col.style]) return sheet.xfs[col.style];
    return sheet.xfs[0];
  }

  function splitSections(fmt) {
    const out = [];
    let cur = '', quoted = false, bracket = false;
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '\\' && !quoted) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
      if (ch === '"' && !bracket) quoted = !quoted;
      else if (ch === '[' && !quoted) bracket = true;
      else if (ch === ']' && !quoted) bracket = false;
      if (ch === ';' && !quoted && !bracket) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  // Font color requested by a number format section, e.g. '#,##0;[Red]-#,##0'.
  function numberFormatColor(fmt, value, indexed) {
    if (!fmt || fmt.indexOf('[') < 0) return null;
    const sections = splitSections(fmt);
    let section;
    if (typeof value !== 'number') {
      section = sections.length >= 4 ? sections[3] : null;
    } else if (/\[\s*(<|>|=|<=|>=|<>)\s*-?[\d.]+\s*\]/.test(fmt)) {
      section = null;
      const cond = /\[\s*(<=|>=|<>|<|>|=)\s*(-?[\d.]+)\s*\]/;
      for (let i = 0; i < Math.min(sections.length, 2) && section === null; i++) {
        const m = cond.exec(sections[i]);
        if (!m) continue;
        const n = parseFloat(m[2]);
        const ok = { '<': value < n, '>': value > n, '=': value === n, '<=': value <= n, '>=': value >= n, '<>': value !== n }[m[1]];
        if (ok) section = sections[i];
      }
      if (section === null) section = sections[Math.min(2, sections.length - 1)];
    } else if (sections.length === 1 || value > 0 || (value === 0 && sections.length < 3)) {
      section = sections[0];
    } else if (value < 0) {
      section = sections[1];
    } else {
      section = sections[2];
    }
    if (!section) return null;
    const m = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*(\d+))\]/i.exec(section);
    if (!m) return null;
    if (m[2]) return (indexed || DEFAULT_INDEXED)[parseInt(m[2], 10) + 7] || null;
    return FORMAT_COLORS[m[1].toLowerCase()];
  }

  global.XlsxFormat = {
    parse,
    colInfo,
    styleAt,
    fontCss,
    fontStack,
    numberFormatColor,
    PX_PER_PT
  };
})(typeof window !== 'undefined' ? window : globalThis);
