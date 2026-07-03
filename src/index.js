#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { readWorkbook, matrixToMarkdownTable } from './lib/read-workbook.js';
import { parseXlsxParts } from './lib/parse-xlsx-parts.js';
import { renderSheetSvgPng } from './lib/render-svg.js';
import { exportSheetsToPdf, pdfToPngBuffers } from './lib/render-excel.js';
import { slugify } from './lib/units.js';

function parseArgs(argv) {
  const args = { input: null, out: null, engine: 'auto', snapshot: true };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--out' || a === '-o') args.out = argv[++i];
    else if (a === '--engine') args.engine = argv[++i]; // auto | excel | svg
    else if (a === '--no-snapshot') args.snapshot = false;
    else if (!a.startsWith('-')) args.input = a;
  }
  return args;
}

// A1-style column letter from a 0-based index.
function colLetter(idx) {
  let s = '';
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.input) {
    console.error('Usage: node src/index.js <workbook.xlsx> [--out dir] [--engine auto|excel|svg] [--no-snapshot]');
    process.exit(1);
  }
  const xlsxPath = path.resolve(args.input);
  if (!fs.existsSync(xlsxPath)) {
    console.error(`File not found: ${xlsxPath}`);
    process.exit(1);
  }
  const outDir = path.resolve(args.out || path.join(path.dirname(xlsxPath), 'output'));
  const assetsDir = path.join(outDir, 'assets');
  const workDir = path.join(outDir, '.work');
  fs.mkdirSync(assetsDir, { recursive: true });
  fs.mkdirSync(workDir, { recursive: true });

  console.log(`Reading ${path.basename(xlsxPath)} ...`);
  const { sheets } = await readWorkbook(xlsxPath);
  const { media, perSheet } = parseXlsxParts(xlsxPath);

  // Write extracted media (de-.tmp'd) to assets.
  const writtenMedia = new Set();
  for (const m of Object.values(media)) {
    const dest = path.join(assetsDir, m.assetName);
    if (!writtenMedia.has(m.assetName)) {
      fs.writeFileSync(dest, m.buf);
      writtenMedia.add(m.assetName);
    }
  }

  // Track each sheet's 1-based position (matches Excel's Worksheets order).
  sheets.forEach((s, i) => (s.pos = i + 1));

  // Which sheets actually have drawings worth snapshotting.
  const sheetsWithDrawings = sheets.filter((s) => {
    const d = perSheet[s.name];
    return d && (d.images.length || d.textboxes.length);
  });

  // Print-area target for each snapshotted sheet: covers the data used-range
  // AND the extent of its drawings, so Excel doesn't paginate the whole sheet.
  const targets = sheetsWithDrawings.map((s) => {
    const objs = [...perSheet[s.name].images, ...perSheet[s.name].textboxes];
    let maxCol = s.lastCol;
    let maxRow = s.lastRow;
    for (const o of objs) {
      maxCol = Math.max(maxCol, (o.geom.to?.col ?? o.geom.from.col) + 1);
      maxRow = Math.max(maxRow, (o.geom.to?.row ?? o.geom.from.row) + 1);
    }
    return { index: s.pos, printArea: `A1:${colLetter(Math.max(0, maxCol - 1))}${Math.max(1, maxRow)}` };
  });

  // --- Snapshots: try Path A (Excel) once, per-sheet fall back to Path B ---
  const snapshots = {}; // sheet name -> [assetName, ...]
  let engineUsed = 'none';
  let excelResult = null;

  if (args.snapshot && sheetsWithDrawings.length) {
    if (args.engine === 'auto' || args.engine === 'excel') {
      console.log('Attempting Path A (Windows Excel export) ...');
      excelResult = exportSheetsToPdf(xlsxPath, workDir, targets);
      if (excelResult.ok) {
        console.log('  Excel export OK.');
        engineUsed = 'excel';
        for (const s of sheetsWithDrawings) {
          const idx = Object.keys(excelResult.nameByIndex).find(
            (k) => excelResult.nameByIndex[k] === s.name,
          );
          const pdf = idx ? excelResult.pdfByIndex[idx] : null;
          if (!pdf) continue;
          const pngs = await pdfToPngBuffers(pdf);
          snapshots[s.name] = pngs.map((buf, i) => {
            const name = `snapshot_${slugify(s.name)}_${i + 1}.png`;
            fs.writeFileSync(path.join(assetsDir, name), buf);
            return name;
          });
        }
      } else {
        console.log(`  Path A unavailable (${excelResult.reason}). Falling back to Path B (SVG reconstruction).`);
      }
    }

    // Path B for any sheet Path A did not cover.
    for (const s of sheetsWithDrawings) {
      if (snapshots[s.name]?.length) continue;
      try {
        const png = renderSheetSvgPng(s, perSheet[s.name]);
        if (png) {
          const name = `snapshot_${slugify(s.name)}.png`;
          fs.writeFileSync(path.join(assetsDir, name), png);
          snapshots[s.name] = [name];
          if (engineUsed !== 'excel') engineUsed = 'svg';
        }
      } catch (e) {
        console.log(`  ! SVG snapshot failed for "${s.name}": ${e.message}`);
      }
    }
  }

  // --- Emit one Markdown file per sheet + an index ---
  const indexRows = [];
  for (const s of sheets) {
    const d = perSheet[s.name] || { images: [], textboxes: [] };
    const slug = slugify(s.name);
    const file = `${slug}.md`;
    const lines = [];
    lines.push(`# ${s.name}`);
    lines.push('');
    if (s.truncated) lines.push(`> ⚠️ Sheet is large; output capped to ${s.lastRow}×${s.lastCol}.`, '');

    lines.push('## Data', '');
    lines.push(matrixToMarkdownTable(s.matrix, s.lastCol), '');

    if (snapshots[s.name]?.length) {
      lines.push('## Snapshot', '');
      lines.push(
        engineUsed === 'excel'
          ? '_Rendered from Excel (images + annotations in position)._'
          : '_Reconstructed layout (images + annotations positioned over nearby cells)._',
        '',
      );
      for (const name of snapshots[s.name]) lines.push(`![snapshot](assets/${name})`, '');
    }

    if (d.textboxes.length) {
      lines.push('## Annotations (textboxes)', '');
      for (const t of d.textboxes) {
        const loc = t.geom?.from ? `\`${colLetter(t.geom.from.col)}${t.geom.from.row + 1}\`` : '';
        lines.push(`- ${loc ? loc + ': ' : ''}${t.text.replace(/\n/g, ' ')}`);
      }
      lines.push('');
    }

    if (d.images.length) {
      lines.push('## Embedded images', '');
      for (const im of d.images) {
        const loc = im.geom?.from ? `${colLetter(im.geom.from.col)}${im.geom.from.row + 1}` : '';
        if (im.asset && im.asset.raster) {
          lines.push(`- \`${loc}\` — ![${im.asset.assetName}](assets/${im.asset.assetName})`);
        } else {
          lines.push(`- \`${loc}\` — ${im.asset ? `${im.asset.assetName} (${im.asset.ext}, not previewable)` : 'unresolved image'}`);
        }
      }
      lines.push('');
    }

    fs.writeFileSync(path.join(outDir, file), lines.join('\n'));
    indexRows.push(
      `| [${s.name}](${encodeURI(file)}) | ${s.lastRow}×${s.lastCol} | ${d.images.length} | ${d.textboxes.length} | ${snapshots[s.name]?.length ? '✓' : ''} |`,
    );
  }

  const indexMd = [
    `# ${path.basename(xlsxPath)}`,
    '',
    `Converted ${sheets.length} sheets.` +
      (engineUsed === 'excel' ? ' Snapshots: Excel render.' : engineUsed === 'svg' ? ' Snapshots: reconstructed (SVG).' : ''),
    '',
    '| Sheet | Size (r×c) | Images | Textboxes | Snapshot |',
    '| --- | --- | --- | --- | --- |',
    ...indexRows,
    '',
  ].join('\n');
  fs.writeFileSync(path.join(outDir, 'index.md'), indexMd);

  // Clean up intermediate work dir (PDFs, ps1).
  try {
    fs.rmSync(workDir, { recursive: true, force: true });
  } catch {
    /* ignore */
  }

  console.log(`\nDone. Output: ${outDir}`);
  console.log(`  index.md + ${sheets.length} sheet files, ${writtenMedia.size} images in assets/.`);
  console.log(`  Snapshots: ${Object.keys(snapshots).length} (engine: ${engineUsed}).`);
  if (excelResult && !excelResult.ok) {
    console.log(`  Note: Excel path was unavailable — ${excelResult.reason}`);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
