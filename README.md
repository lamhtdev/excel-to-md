# excel-to-md

Convert a multi-sheet Excel workbook — including **embedded images** and **textbox
annotations** — into Markdown, one `.md` file per sheet.

The run environment is **Node + npm only**: no LibreOffice, poppler, or system
packages. High-fidelity rendering is delegated **off-WSL to the installed
Windows Excel**, with a fully self-contained fallback when Excel isn't usable.

## Usage

```bash
npm install
node src/index.js "<workbook.xlsx>" [--out <dir>] [--engine auto|excel|svg] [--no-snapshot]
```

- `--out <dir>` — output directory (default: `<workbook-dir>/output`).
- `--engine` — snapshot renderer:
  - `auto` (default) — try Excel, fall back to reconstruction per sheet.
  - `excel` — Path A only (Windows Excel).
  - `svg` — Path B only (self-contained reconstruction).
- `--no-snapshot` — skip snapshots (tables + extracted images only).

## Output

```
output/
  index.md                 # overview table linking every sheet
  <Sheet>.md               # one file per sheet
  assets/
    imageN.png             # extracted embedded images (extensions sniffed)
    snapshot_<Sheet>*.png  # rendered snapshots of annotated sheets
```

Each sheet file contains: a **Data** table (values-only GFM), a **Snapshot** of
any images/annotations in position, an **Annotations (textboxes)** list, and an
**Embedded images** gallery.

## How it works

1. **Data** — `exceljs` reads each sheet's values into a GFM table. Merged cells
   keep their value only in the top-left cell (no colspan) to avoid duplication.
2. **Images** — the `.xlsx` is unzipped directly; each `xl/media/*` part has its
   real type detected from magic bytes (Excel often stores images as `.tmp`) and
   is written to `assets/`.
3. **Textboxes** — `xl/drawings/*.xml` is parsed for shape text and anchor cells.
4. **Snapshots** — two paths:
   - **Path A (fidelity):** drive Windows Excel via `powershell.exe` to export each
     annotated sheet's print area to PDF, then rasterize to PNG
     (`pdf-to-png-converter`). Pixel-true to Excel.
   - **Path B (self-contained):** reconstruct an SVG from the parsed geometry
     (column widths / row heights / anchors) with images and textboxes overlaid on
     a light grid, rasterized with `@resvg/resvg-js`. No Excel, no printer.

   `auto` tries A and falls back to B for any sheet A couldn't render.

## Requirements & notes

- **Node 18+.**
- **Path A** needs Windows Excel reachable from WSL via `powershell.exe`, and a
  **default printer** must exist (e.g. enable *Microsoft Print to PDF*). Excel
  raises a *"missing printer"* error otherwise — the tool then falls back to Path B
  automatically.
- **Path B** layout is a faithful reconstruction: positions are accurate, but fonts
  and text wrapping are approximate, not identical to Excel.
- Vector images (EMF/WMF) can't preview in Markdown; they're listed but not inlined.
- Large sheets are capped (20000 rows × 256 cols) with a note in the output.
