import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
export interface ChartBar {
  label: string;
  count: number;
  color?: string;
}

const CHART_SCRIPT = `
Add-Type -AssemblyName System.Drawing
$jsonPath = $args[0]
$outPath = $args[1]
$buckets = Get-Content -Raw -Encoding UTF8 -Path $jsonPath | ConvertFrom-Json
$width = 1080
$height = 36 + ($buckets.Count * 48)
$bmp = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$graphics.Clear([System.Drawing.Color]::White)
$labelFont = New-Object System.Drawing.Font "Segoe UI", 12
$countFont = New-Object System.Drawing.Font "Segoe UI", 12, ([System.Drawing.FontStyle]::Bold)
$muted = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(71, 85, 105))
$ink = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(15, 23, 42))
$trackBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(241, 245, 249))
$palette = @(
  [System.Drawing.Color]::FromArgb(203, 213, 225),
  [System.Drawing.Color]::FromArgb(148, 163, 184),
  [System.Drawing.Color]::FromArgb(245, 158, 11),
  [System.Drawing.Color]::FromArgb(249, 115, 22),
  [System.Drawing.Color]::FromArgb(225, 29, 72)
)
function Add-Rounded([System.Drawing.Drawing2D.GraphicsPath]$path, [single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
  if ($w -le 1) { return }
  $d = [Math]::Min($r * 2, [Math]::Min($w, $h))
  $path.AddArc($x, $y, $d, $d, 180, 90) | Out-Null
  $path.AddArc(($x + $w - $d), $y, $d, $d, 270, 90) | Out-Null
  $path.AddArc(($x + $w - $d), ($y + $h - $d), $d, $d, 0, 90) | Out-Null
  $path.AddArc($x, ($y + $h - $d), $d, $d, 90, 90) | Out-Null
  $path.CloseFigure() | Out-Null
}
$peak = 1
foreach ($bucket in $buckets) { if ([int]$bucket.count -gt $peak) { $peak = [int]$bucket.count } }
$index = 0
foreach ($bucket in $buckets) {
  $y = 16 + ($index * 48)
  $label = ([string]$bucket.label).Replace([string][char]0x2013, "-")
  $graphics.DrawString($label, $labelFont, $muted, 16, ($y + 4))
  $track = New-Object System.Drawing.Drawing2D.GraphicsPath
  Add-Rounded $track 230 ($y + 8) 720 16 8
  $graphics.FillPath($trackBrush, $track)
  $count = [int]$bucket.count
  $barWidth = 0
  if ($count -gt 0) { $barWidth = [Math]::Max(16, [int](($count / $peak) * 720)) }
  if ($barWidth -gt 0) {
    $bar = New-Object System.Drawing.Drawing2D.GraphicsPath
    $color = $palette[[Math]::Min($index, $palette.Length - 1)]
    if ($bucket.color) {
      $hex = ([string]$bucket.color).TrimStart('#')
      if ($hex.Length -eq 6) {
        $color = [System.Drawing.Color]::FromArgb(
          [Convert]::ToInt32($hex.Substring(0, 2), 16),
          [Convert]::ToInt32($hex.Substring(2, 2), 16),
          [Convert]::ToInt32($hex.Substring(4, 2), 16)
        )
      }
    }
    $brush = New-Object System.Drawing.SolidBrush $color
    Add-Rounded $bar 230 ($y + 8) $barWidth 16 8
    $graphics.FillPath($brush, $bar)
    $brush.Dispose()
    $bar.Dispose()
  }
  $graphics.DrawString([string]$count, $countFont, $ink, 964, ($y + 2))
  $track.Dispose()
  $index++
}
$bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bmp.Dispose()
$labelFont.Dispose()
$countFont.Dispose()
$muted.Dispose()
$ink.Dispose()
$trackBrush.Dispose()
`;

export function renderAgingChartPng(buckets: ChartBar[]): Buffer {
  const dir = mkdtempSync(path.join(tmpdir(), "pmt-chart-"));
  const jsonPath = path.join(dir, "buckets.json");
  const pngPath = path.join(dir, "chart.png");
  const scriptPath = path.join(dir, "chart.ps1");
  try {
    const rows = buckets.slice(0, 8).map((bucket) => ({
      ...bucket,
      label: bucket.label.replace(/[\u2013\u2014]/g, "-"),
    }));
    writeFileSync(jsonPath, JSON.stringify(rows), "utf8");
    writeFileSync(scriptPath, CHART_SCRIPT, "utf8");
    const result = spawnSync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, jsonPath, pngPath],
      { encoding: "utf8", windowsHide: true },
    );
    if (result.status !== 0) {
      const detail = (result.stderr || result.stdout || "Chart drawing failed.").trim();
      throw new Error(detail);
    }
    return readFileSync(pngPath);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

export function renderAgingChartBase64(buckets: ChartBar[]): string {
  return renderAgingChartPng(buckets).toString("base64");
}
