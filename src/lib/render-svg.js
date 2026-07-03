import { Resvg } from '@resvg/resvg-js';
import { EMU_PER_PIXEL } from './units.js';

const MARGIN = 24;
const CELL_PAD = 3;
const CTX_CELLS = 2; // grid cells of context to draw around the drawings
const FONT = 11;
const MAX_CANVAS = 6000; // px guard so a stray anchor can't blow up the raster

function escapeXml(s) {
  return String(s).replace(/[<>&"']/g, (c) =>
    ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' }[c]),
  );
}

// Prefix offsets: cumulative[i] = pixels before column/row index i (0-based),
// extending past the known array with the sheet default.
function makePrefix(pxArray, defaultPx, count) {
  const prefix = [0];
  for (let i = 0; i < count; i++) {
    const w = i < pxArray.length ? pxArray[i] : defaultPx;
    prefix.push(prefix[i] + w);
  }
  return prefix;
}

// Pixel X of an anchor point (col index + EMU offset into that column).
function anchorX(prefix, pxArray, defaultPx, col, colOff) {
  const base = col < prefix.length ? prefix[col] : prefix[prefix.length - 1] + (col - (prefix.length - 1)) * defaultPx;
  return base + (colOff || 0) / EMU_PER_PIXEL;
}
function anchorY(prefix, pxArray, defaultPx, row, rowOff) {
  const base = row < prefix.length ? prefix[row] : prefix[prefix.length - 1] + (row - (prefix.length - 1)) * defaultPx;
  return base + (rowOff || 0) / EMU_PER_PIXEL;
}

// Resolve an object's pixel rectangle from its anchor (two-cell / one-cell).
function objectRect(obj, xPrefix, sheet, yPrefix) {
  const { colPx, rowPx, defaultColPx, defaultRowPx } = sheet;
  const g = obj.geom;
  const x1 = anchorX(xPrefix, colPx, defaultColPx, g.from.col, g.from.colOff);
  const y1 = anchorY(yPrefix, rowPx, defaultRowPx, g.from.row, g.from.rowOff);
  let x2, y2;
  if (g.to) {
    x2 = anchorX(xPrefix, colPx, defaultColPx, g.to.col, g.to.colOff);
    y2 = anchorY(yPrefix, rowPx, defaultRowPx, g.to.row, g.to.rowOff);
  } else if (g.ext) {
    x2 = x1 + Number(g.ext['@_cx'] || 0) / EMU_PER_PIXEL;
    y2 = y1 + Number(g.ext['@_cy'] || 0) / EMU_PER_PIXEL;
  } else {
    x2 = x1 + defaultColPx;
    y2 = y1 + defaultRowPx;
  }
  return { x: x1, y: y1, w: Math.max(1, x2 - x1), h: Math.max(1, y2 - y1) };
}

// Naive word wrap for textbox contents inside a fixed pixel width.
function wrapText(text, widthPx) {
  const maxChars = Math.max(4, Math.floor(widthPx / (FONT * 0.55)));
  const out = [];
  for (const rawLine of text.split('\n')) {
    let line = '';
    for (const word of rawLine.split(/(\s+)/)) {
      if ((line + word).length > maxChars && line) {
        out.push(line);
        line = word.trimStart();
      } else {
        line += word;
      }
    }
    out.push(line);
  }
  return out;
}

// Render a per-sheet snapshot PNG showing images + textboxes anchored over a
// light grid of the surrounding cells. Returns a Buffer, or null if the sheet
// has nothing to draw.
export function renderSheetSvgPng(sheet, drawings) {
  const objects = [...drawings.images, ...drawings.textboxes];
  if (objects.length === 0) return null;

  const maxCol = Math.max(...objects.map((o) => (o.geom.to?.col ?? o.geom.from.col) + 1)) + CTX_CELLS;
  const maxRow = Math.max(...objects.map((o) => (o.geom.to?.row ?? o.geom.from.row) + 1)) + CTX_CELLS;
  const xPrefix = makePrefix(sheet.colPx, sheet.defaultColPx, maxCol + 1);
  const yPrefix = makePrefix(sheet.rowPx, sheet.defaultRowPx, maxRow + 1);

  const rects = objects.map((o) => ({ obj: o, r: objectRect(o, xPrefix, sheet, yPrefix) }));

  // Bounding box across all objects, expanded by context cells.
  const minColIdx = Math.max(0, Math.min(...objects.map((o) => o.geom.from.col)) - CTX_CELLS);
  const minRowIdx = Math.max(0, Math.min(...objects.map((o) => o.geom.from.row)) - CTX_CELLS);
  const bx = xPrefix[minColIdx];
  const by = yPrefix[minRowIdx];
  let maxX = Math.max(...rects.map((it) => it.r.x + it.r.w));
  let maxY = Math.max(...rects.map((it) => it.r.y + it.r.h));
  maxX = Math.min(maxX, bx + MAX_CANVAS);
  maxY = Math.min(maxY, by + MAX_CANVAS);

  const width = Math.ceil(maxX - bx + MARGIN * 2);
  const height = Math.ceil(maxY - by + MARGIN * 2);
  const ox = MARGIN - bx; // world -> canvas offset
  const oy = MARGIN - by;

  const svg = [];
  svg.push(`<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`);
  svg.push(`<rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff"/>`);

  // Grid lines + cell text within the visible index range.
  svg.push(`<g stroke="#e0e0e0" stroke-width="1">`);
  for (let c = minColIdx; c <= maxCol; c++) {
    const x = xPrefix[c] + ox;
    if (x < 0 || x > width) continue;
    svg.push(`<line x1="${x}" y1="0" x2="${x}" y2="${height}"/>`);
  }
  for (let r = minRowIdx; r <= maxRow; r++) {
    const y = yPrefix[r] + oy;
    if (y < 0 || y > height) continue;
    svg.push(`<line x1="0" y1="${y}" x2="${width}" y2="${y}"/>`);
  }
  svg.push(`</g>`);

  // Cell text from the value matrix (0-based matrix -> 0-based grid indices).
  svg.push(`<g font-family="Segoe UI, Arial, sans-serif" font-size="${FONT}" fill="#333">`);
  for (let r = minRowIdx; r <= maxRow && r < sheet.matrix.length; r++) {
    for (let c = minColIdx; c <= maxCol && c < (sheet.matrix[r]?.length || 0); c++) {
      const val = sheet.matrix[r][c];
      if (!val) continue;
      const x = xPrefix[c] + ox + CELL_PAD;
      const y = yPrefix[r] + oy + FONT + CELL_PAD;
      const colW = (xPrefix[c + 1] ?? xPrefix[c] + sheet.defaultColPx) - xPrefix[c];
      const maxChars = Math.max(2, Math.floor((colW - CELL_PAD * 2) / (FONT * 0.55)));
      const txt = val.length > maxChars ? val.slice(0, maxChars - 1) + '…' : val;
      svg.push(`<text x="${x}" y="${y}">${escapeXml(txt)}</text>`);
    }
  }
  svg.push(`</g>`);

  // Images (raster embedded as data URIs; vector shown as a labelled box).
  for (const { obj, r } of rects) {
    if (obj.kind !== 'image') continue;
    const x = r.x + ox, y = r.y + oy;
    if (obj.asset && obj.asset.raster) {
      const mime = obj.asset.ext === 'jpg' ? 'jpeg' : obj.asset.ext;
      const b64 = obj.asset.buf.toString('base64');
      svg.push(`<image x="${x}" y="${y}" width="${r.w}" height="${r.h}" preserveAspectRatio="none" xlink:href="data:image/${mime};base64,${b64}"/>`);
    } else {
      svg.push(`<rect x="${x}" y="${y}" width="${r.w}" height="${r.h}" fill="#f5f5f5" stroke="#bbb" stroke-dasharray="4 3"/>`);
      svg.push(`<text x="${x + 4}" y="${y + 14}" font-family="Arial" font-size="10" fill="#999">${escapeXml(obj.asset?.ext || 'image')}</text>`);
    }
  }

  // Textboxes / callouts.
  for (const { obj, r } of rects) {
    if (obj.kind !== 'textbox') continue;
    const x = r.x + ox, y = r.y + oy;
    svg.push(`<rect x="${x}" y="${y}" width="${r.w}" height="${r.h}" fill="#fff8c4" fill-opacity="0.85" stroke="#e0b400" stroke-width="1"/>`);
    const lines = wrapText(obj.text, r.w - 8);
    svg.push(`<g font-family="Segoe UI, Arial, sans-serif" font-size="${FONT}" fill="#5a4b00">`);
    lines.forEach((ln, i) => {
      svg.push(`<text x="${x + 4}" y="${y + FONT + 3 + i * (FONT + 3)}">${escapeXml(ln)}</text>`);
    });
    svg.push(`</g>`);
  }

  svg.push(`</svg>`);

  const resvg = new Resvg(svg.join('\n'), { background: 'white' });
  return resvg.render().asPng();
}
