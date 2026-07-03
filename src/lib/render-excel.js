import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

// Translate a WSL path to a Windows path via wslpath. Returns null off-WSL.
function toWin(p) {
  try {
    return execFileSync('wslpath', ['-w', p], { encoding: 'utf8' }).trim();
  } catch {
    return null;
  }
}

function findPowershell() {
  const candidates = [
    'powershell.exe',
    '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe',
    '/mnt/c/WINDOWS/System32/WindowsPowerShell/v1.0/powershell.exe',
  ];
  for (const c of candidates) {
    try {
      execFileSync(c, ['-NoProfile', '-Command', 'exit 0'], { stdio: 'ignore', timeout: 15000 });
      return c;
    } catch {
      /* try next */
    }
  }
  return null;
}

const PS_SCRIPT = `param([string]$Workbook, [string]$OutDir, [string]$TargetsJson)
$ErrorActionPreference = 'Stop'
$targets = @{}
if ($TargetsJson -and (Test-Path $TargetsJson)) {
  (Get-Content $TargetsJson -Raw | ConvertFrom-Json) | ForEach-Object { $targets[[int]$_.index] = $_.printArea }
}
try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  Write-Output "FATAL|Excel COM not available: $($_.Exception.Message)"
  exit 2
}
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($Workbook, $false, $true)
  $i = 0
  foreach ($ws in $wb.Worksheets) {
    $i++
    if (-not $targets.ContainsKey($i)) { continue }   # only requested sheets
    $pdf = Join-Path $OutDir ("sheet_" + $i + ".pdf")
    try {
      $ps = $ws.PageSetup
      $ps.Orientation = 2                # landscape -> annotated screenshots read wider
      if ($targets[$i]) { $ps.PrintArea = $targets[$i] }
      $ps.Zoom = $false                  # required for FitToPages to apply
      $ps.FitToPagesWide = 1             # never split a sheet horizontally
      $ps.FitToPagesTall = $false        # allow tall content to flow onto more pages
      $ps.LeftMargin = 10; $ps.RightMargin = 10; $ps.TopMargin = 10; $ps.BottomMargin = 10
      $ws.ExportAsFixedFormat(0, $pdf)   # 0 = xlTypePDF
      Write-Output ("SHEET|" + $i + "|" + $ws.Name)
    } catch {
      Write-Output ("SHEETERR|" + $i + "|" + $ws.Name + "|" + $_.Exception.Message)
    }
  }
  $wb.Close($false)
  Write-Output "DONE"
} catch {
  Write-Output ("FATAL|" + $_.Exception.Message)
  exit 3
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
`;

// Attempt to export the requested worksheets to PDF using the installed Windows
// Excel. `targets` is [{ index (1-based sheet position), printArea (A1 range) }].
// Returns { ok, pdfByIndex: {index -> pdfPath}, nameByIndex, reason }.
export function exportSheetsToPdf(xlsxPath, workDir, targets) {
  const ps = findPowershell();
  if (!ps) return { ok: false, reason: 'powershell.exe not reachable (not on WSL/Windows)' };

  const winWorkbook = toWin(path.resolve(xlsxPath));
  const outDir = path.join(workDir, 'pdf');
  fs.mkdirSync(outDir, { recursive: true });
  const winOutDir = toWin(outDir);
  if (!winWorkbook || !winOutDir) return { ok: false, reason: 'wslpath translation failed' };

  const scriptPath = path.join(workDir, 'export.ps1');
  fs.writeFileSync(scriptPath, PS_SCRIPT, 'utf8');
  const winScript = toWin(scriptPath);

  const targetsPath = path.join(workDir, 'targets.json');
  fs.writeFileSync(targetsPath, JSON.stringify(targets || []), 'utf8');
  const winTargets = toWin(targetsPath);

  let out = '';
  try {
    out = execFileSync(
      ps,
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', winScript,
       '-Workbook', winWorkbook, '-OutDir', winOutDir, '-TargetsJson', winTargets],
      { encoding: 'utf8', timeout: 180000 },
    );
  } catch (e) {
    out = (e.stdout || '') + (e.stderr || '');
    const fatal = /FATAL\|(.*)/.exec(out);
    return { ok: false, reason: fatal ? fatal[1].trim() : `Excel export failed: ${e.message}` };
  }

  const fatal = /FATAL\|(.*)/.exec(out);
  if (fatal) return { ok: false, reason: fatal[1].trim() };

  const pdfByIndex = {};
  const nameByIndex = {};
  const errors = [];
  for (const line of out.split(/\r?\n/)) {
    let m;
    if ((m = /^SHEET\|(\d+)\|(.*)$/.exec(line))) {
      const idx = Number(m[1]);
      const pdf = path.join(outDir, `sheet_${idx}.pdf`);
      if (fs.existsSync(pdf)) {
        pdfByIndex[idx] = pdf;
        nameByIndex[idx] = m[2];
      }
    } else if ((m = /^SHEETERR\|(\d+)\|([^|]*)\|(.*)$/.exec(line))) {
      errors.push({ index: Number(m[1]), name: m[2], message: m[3] });
    }
  }

  if (Object.keys(pdfByIndex).length === 0) {
    const why = errors[0]?.message || 'no PDFs produced';
    return { ok: false, reason: why, errors };
  }
  return { ok: true, pdfByIndex, nameByIndex, errors };
}

// Rasterize a PDF (all pages) to PNG buffers using the pure-npm converter.
export async function pdfToPngBuffers(pdfPath) {
  const { pdfToPng } = await import('pdf-to-png-converter');
  const pages = await pdfToPng(pdfPath, { viewportScale: 2.0 });
  return pages.map((p) => p.content);
}
