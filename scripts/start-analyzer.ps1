$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$url = "http://localhost:3000"

function Test-AnalyzerRunning {
  try {
    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 3
    return $response.StatusCode -ge 200
  } catch {
    return $false
  }
}

if (Test-AnalyzerRunning) {
  Start-Process $url
  exit 0
}

Set-Location $root

if (-not (Test-Path (Join-Path $root "node_modules"))) {
  Write-Host "Installing dependencies..."
  npm install
}

Write-Host "Starting M365 Transcript AI Analyzer..."
Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "npm run dev" -WorkingDirectory $root -WindowStyle Minimized

$deadline = (Get-Date).AddSeconds(90)
while (-not (Test-AnalyzerRunning)) {
  if ((Get-Date) -gt $deadline) {
    Write-Host "Server did not become ready in 90 seconds. Opening the browser anyway."
    break
  }
  Start-Sleep -Seconds 1
}

Start-Process $url
