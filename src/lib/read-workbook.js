import ExcelJS from 'exceljs';
import { colWidthToPx, rowHeightToPt, ptToPx, mdCell } from './units.js';

const MAX_COLS = 256;
const MAX_ROWS = 20000;

// Flatten any ExcelJS cell value into a plain display string.
export function valueToString(v) {
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join('');
    if ('text' in v) return String(v.text); // hyperlink
    if ('result' in v) return v.result == null ? '' : String(v.result); // formula -> computed value
    if ('formula' in v) return ''; // formula with no cached result
    if ('error' in v) return String(v.error);
    if ('sharedFormula' in v) return v.result == null ? '' : String(v.result);
  }
  return String(v);
}

// Load the workbook and return per-sheet: display order, name, value matrix,
// and the column/row pixel geometry needed to place drawings for Path B.
export async function readWorkbook(xlsxPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(xlsxPath);

  const sheets = [];
  wb.eachSheet((ws) => {
    if (ws.state === 'veryHidden' || ws.state === 'hidden') {
      // Keep hidden sheets but flag them; the workbook order still matters.
    }
    const dim = ws.dimensions || {};
    let lastRow = Math.min(ws.actualRowCount || dim.bottom || 0, MAX_ROWS);
    let lastCol = Math.min(ws.actualColumnCount || dim.right || 0, MAX_COLS);
    const truncated =
      (ws.actualRowCount || 0) > MAX_ROWS || (ws.actualColumnCount || 0) > MAX_COLS;

    const defaultColWidth = ws.properties?.defaultColWidth ?? 8.43;
    const defaultRowHeight = ws.properties?.defaultRowHeight ?? 15;

    // Value matrix (rows x cols), 0-indexed for our own use.
    const matrix = [];
    for (let r = 1; r <= lastRow; r++) {
      const row = ws.getRow(r);
      const cells = [];
      for (let c = 1; c <= lastCol; c++) {
        const cell = row.getCell(c);
        // ExcelJS mirrors a merged value into every covered cell; keep it only
        // in the top-left master so the table isn't flooded with duplicates.
        if (cell.isMerged && cell.master && cell.master !== cell) {
          cells.push('');
        } else {
          cells.push(valueToString(cell.value));
        }
      }
      matrix.push(cells);
    }

    // Column pixel widths and offsets (prefix sums) for anchor math.
    const colPx = [];
    for (let c = 1; c <= Math.max(lastCol, 1); c++) {
      const w = ws.getColumn(c).width;
      colPx.push(colWidthToPx(w, defaultColWidth));
    }
    const rowPx = [];
    for (let r = 1; r <= Math.max(lastRow, 1); r++) {
      const h = ws.getRow(r).height;
      rowPx.push(ptToPx(rowHeightToPt(h, defaultRowHeight)));
    }

    sheets.push({
      id: ws.id,
      name: ws.name,
      state: ws.state,
      lastRow,
      lastCol,
      truncated,
      matrix,
      colPx,
      rowPx,
      defaultColPx: colWidthToPx(undefined, defaultColWidth),
      defaultRowPx: ptToPx(rowHeightToPt(undefined, defaultRowHeight)),
    });
  });

  return { sheets };
}

// Render a value matrix as a GitHub-Flavored-Markdown table. The first row is
// used as the header; if the sheet is empty, returns a note instead.
export function matrixToMarkdownTable(matrix, lastCol) {
  const nonEmpty = matrix.filter((row) => row.some((c) => c !== ''));
  if (nonEmpty.length === 0 || lastCol === 0) return '_(empty sheet)_';

  const width = lastCol;
  const pad = (row) => {
    const r = row.slice(0, width);
    while (r.length < width) r.push('');
    return r;
  };

  const [header, ...rest] = matrix;
  const headerCells = pad(header).map((c, i) => mdCell(c) || `Col ${i + 1}`);
  const sep = headerCells.map(() => '---');
  const lines = [
    `| ${headerCells.join(' | ')} |`,
    `| ${sep.join(' | ')} |`,
  ];
  for (const row of rest) {
    lines.push(`| ${pad(row).map(mdCell).join(' | ')} |`);
  }
  return lines.join('\n');
}
