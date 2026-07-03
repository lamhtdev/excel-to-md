// Geometry + format helpers shared across the pipeline.

// EMU (English Metric Units) per pixel at 96 DPI: 914400 EMU/inch / 96 px/inch.
export const EMU_PER_PIXEL = 9525;

// Excel column "width" is measured in characters of the default font. For
// Calibri 11 the Maximum Digit Width (MDW) is 7px. Excel's own conversion:
//   px = round(width * MDW + 5)   (the +5 is cell padding / gridline)
const MDW = 7;
export function colWidthToPx(width, defaultWidth = 8.43) {
  const w = width == null ? defaultWidth : width;
  return Math.round(w * MDW + 5);
}

// Row height is stored in points. 96 DPI => px = pt * 96/72.
export function rowHeightToPt(height, defaultHeight = 15) {
  return height == null ? defaultHeight : height;
}
export function ptToPx(pt) {
  return Math.round((pt * 96) / 72);
}

// Sniff the real image type from magic bytes. Excel frequently stores images
// with a `.tmp` extension or a mismatched one, so we never trust the name.
export function sniffImageExt(buf) {
  if (!buf || buf.length < 12) return null;
  // PNG
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'png';
  // JPEG
  if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'jpg';
  // GIF
  if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return 'gif';
  // BMP
  if (buf[0] === 0x42 && buf[1] === 0x4d) return 'bmp';
  // WEBP: RIFF....WEBP
  if (buf[0] === 0x52 && buf[1] === 0x49 && buf[2] === 0x46 && buf[3] === 0x46 &&
      buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50) return 'webp';
  // EMF (Enhanced Metafile): record type 0x00000001 then 'EMF ' signature at 0x28.
  if (buf[0] === 0x01 && buf[1] === 0x00 && buf[2] === 0x00 && buf[3] === 0x00) return 'emf';
  // WMF
  if ((buf[0] === 0xd7 && buf[1] === 0xcd) || (buf[0] === 0x01 && buf[1] === 0x00 && buf[2] === 0x09)) return 'wmf';
  return null;
}

// Vector metafiles (EMF/WMF) don't render in Markdown or resvg; flag them so
// callers can warn instead of emitting a broken image link.
export function isRasterExt(ext) {
  return ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'].includes(ext);
}

// Filesystem-safe slug for sheet names used as output filenames. Preserves
// unicode letters (Vietnamese) but drops path-hostile characters.
export function slugify(name) {
  return String(name)
    .replace(/[\/\\:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/ /g, '_') || 'sheet';
}

// Escape a value for a GitHub-Flavored-Markdown table cell.
export function mdCell(value) {
  if (value == null) return '';
  let s = String(value);
  s = s.replace(/\r\n|\r|\n/g, '<br>'); // keep multi-line cells on one row
  s = s.replace(/\|/g, '\\|');
  return s;
}
