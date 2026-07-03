import AdmZip from 'adm-zip';
import { XMLParser } from 'fast-xml-parser';
import { sniffImageExt, isRasterExt } from './units.js';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  removeNSPrefix: true, // <xdr:twoCellAnchor> -> twoCellAnchor, @_r:embed -> @_embed
  trimValues: false, // preserve leading/trailing spaces in textbox runs ("edit"+" button")
});

const arr = (x) => (x == null ? [] : Array.isArray(x) ? x : [x]);

// Resolve a relative path like "../media/image1.png" against a part dir.
function resolvePath(baseDir, target) {
  const parts = (baseDir + '/' + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

function dirOf(path) {
  return path.split('/').slice(0, -1).join('/');
}

// Parse a *.rels file into { rId -> absolute-part-path }.
function parseRels(zip, relsPath) {
  const entry = zip.getEntry(relsPath);
  if (!entry) return {};
  const xml = parser.parse(zip.readAsText(entry));
  const baseDir = dirOf(dirOf(relsPath)); // strip "_rels/x.rels" -> owning dir
  const map = {};
  for (const rel of arr(xml?.Relationships?.Relationship)) {
    const id = rel['@_Id'];
    const target = rel['@_Target'];
    if (!id || !target) continue;
    if (rel['@_TargetMode'] === 'External') {
      map[id] = { external: true, target };
    } else {
      map[id] = { external: false, target: resolvePath(baseDir, target) };
    }
  }
  return map;
}

// Collect every <a:t> text node under a shape's text body, preserving
// paragraph breaks between <a:p> blocks.
function extractShapeText(sp) {
  const txBody = sp?.txBody;
  if (!txBody) return '';
  const paras = arr(txBody.p).map((p) => {
    const runs = arr(p.r).map((r) => (r?.t == null ? '' : String(r.t)));
    // Some textboxes put text directly in <a:fld> or a bare <a:t>.
    if (p.t != null) runs.push(String(p.t));
    for (const fld of arr(p.fld)) if (fld?.t != null) runs.push(String(fld.t));
    return runs.join('');
  });
  return paras.join('\n').trim();
}

// Anchor -> { from:{col,colOff,row,rowOff}, to:{...} } in 0-based cell coords.
function readAnchorCells(anchor) {
  const cell = (node) =>
    node
      ? {
          col: Number(node.col ?? 0),
          colOff: Number(node.colOff ?? 0),
          row: Number(node.row ?? 0),
          rowOff: Number(node.rowOff ?? 0),
        }
      : null;
  return { from: cell(anchor.from), to: cell(anchor.to), ext: anchor.ext };
}

// Parse one drawing XML into a flat list of placed objects (pics + textboxes).
function parseDrawing(zip, drawingPath) {
  const entry = zip.getEntry(drawingPath);
  if (!entry) return [];
  const xml = parser.parse(zip.readAsText(entry));
  const ws = xml?.wsDr;
  if (!ws) return [];
  const rels = parseRels(zip, `${dirOf(drawingPath)}/_rels/${drawingPath.split('/').pop()}.rels`);

  const objects = [];
  const anchorTypes = ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'];
  for (const type of anchorTypes) {
    for (const anchor of arr(ws[type])) {
      const geom = readAnchorCells(anchor);

      // Picture
      for (const pic of arr(anchor.pic)) {
        const embed = pic?.blipFill?.blip?.['@_embed'];
        const rel = embed ? rels[embed] : null;
        objects.push({
          kind: 'image',
          mediaPath: rel && !rel.external ? rel.target : null,
          name: pic?.nvPicPr?.cNvPr?.['@_name'] || '',
          geom,
          anchorType: type,
        });
      }
      // Shapes (textboxes / callouts). Only keep those carrying text.
      for (const sp of arr(anchor.sp)) {
        const text = extractShapeText(sp);
        if (!text) continue;
        objects.push({
          kind: 'textbox',
          text,
          name: sp?.nvSpPr?.cNvPr?.['@_name'] || '',
          geom,
          anchorType: type,
        });
      }
    }
  }
  return objects;
}

// Top-level: read all workbook parts, extract & de-.tmp media, and map each
// sheet (in workbook order) to its placed images + textboxes.
export function parseXlsxParts(xlsxPath) {
  const zip = new AdmZip(xlsxPath);

  // 1) Workbook sheet order + rels -> physical worksheet files.
  const wbXml = parser.parse(zip.readAsText('xl/workbook.xml'));
  const wbRels = parseRels(zip, 'xl/_rels/workbook.xml.rels');
  const sheetOrder = arr(wbXml?.workbook?.sheets?.sheet).map((s) => ({
    name: s['@_name'],
    rid: s['@_id'],
    sheetPath: wbRels[s['@_id']]?.target,
  }));

  // 2) Extract every media part, sniffing its true type (Excel writes .tmp).
  const media = {}; // absolute part path -> { assetName, ext, buf, raster }
  let idx = 0;
  for (const entry of zip.getEntries()) {
    if (!/^xl\/media\//.test(entry.entryName)) continue;
    idx++;
    const buf = entry.getData();
    const ext = sniffImageExt(buf) || (entry.entryName.split('.').pop() || 'bin');
    media[entry.entryName] = {
      assetName: `image${idx}.${ext}`,
      ext,
      buf,
      raster: isRasterExt(ext),
    };
  }

  // 3) Per-sheet drawings.
  const perSheet = {}; // sheet name -> { images:[], textboxes:[] }
  for (const s of sheetOrder) {
    const result = { images: [], textboxes: [] };
    perSheet[s.name] = result;
    if (!s.sheetPath) continue;
    const sheetXml = parser.parse(zip.readAsText(s.sheetPath));
    const drawingRid = sheetXml?.worksheet?.drawing?.['@_id'];
    if (!drawingRid) continue;
    const sheetRels = parseRels(
      zip,
      `${dirOf(s.sheetPath)}/_rels/${s.sheetPath.split('/').pop()}.rels`,
    );
    const drawingPath = sheetRels[drawingRid]?.target;
    if (!drawingPath) continue;

    for (const obj of parseDrawing(zip, drawingPath)) {
      if (obj.kind === 'image') {
        const m = obj.mediaPath ? media[obj.mediaPath] : null;
        result.images.push({ ...obj, asset: m || null });
      } else {
        result.textboxes.push(obj);
      }
    }
  }

  return { media, perSheet, sheetOrder };
}
