param(
  [Parameter(Mandatory=$false)][string]$PayloadPath = "C:\Temp\_raw\payload_invalid.json",
  [Parameter(Mandatory=$false)][string]$ReportPath = "C:\Temp\_raw\payload_validation_report.txt"
)

function Get-Slice([string]$text, [int]$pos, [int]$radius = 120) {
  if ([string]::IsNullOrEmpty($text)) { return "" }
  if ($pos -lt 1) { $pos = 1 }
  if ($pos -gt $text.Length) { $pos = $text.Length }
  $start = [Math]::Max(0, $pos - 1 - $radius)
  $len = [Math]::Min($text.Length - $start, ($radius * 2) + 1)
  return $text.Substring($start, $len)
}

if (-not (Test-Path -LiteralPath $PayloadPath)) {
  Write-Error "Ficheiro não encontrado: $PayloadPath"
  exit 1
}

$raw = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8
$lines = @()
$ok = $true

try {
  $null = $raw | ConvertFrom-Json -Depth 100
  $lines += "PowerShell ConvertFrom-Json: PASS"
} catch {
  $ok = $false
  $msg = $_.Exception.Message
  $lines += "PowerShell ConvertFrom-Json: FAIL"
  $lines += "Erro: $msg"

  $pos = 0
  if ($msg -match 'position\s+(\d+)') { $pos = [int]$matches[1] }
  if ($pos -gt 0) {
    $lines += "Posição aproximada: $pos"
    $lines += "Slice:"
    $lines += (Get-Slice $raw $pos 120)
  }
}

$python = Get-Command python -ErrorAction SilentlyContinue
if ($python) {
  $tmp = [System.IO.Path]::GetTempFileName()
  try {
    python -m json.tool "$PayloadPath" *> $tmp
    if ($LASTEXITCODE -eq 0) {
      $lines += "python -m json.tool: PASS"
    } else {
      $ok = $false
      $lines += "python -m json.tool: FAIL"
      $lines += (Get-Content -LiteralPath $tmp -Raw)
    }
  } finally {
    Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
  }
} else {
  $lines += "python -m json.tool: SKIP (python não disponível)"
}

$dir = Split-Path -Parent $ReportPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
$lines | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$lines | ForEach-Object { Write-Host $_ }

if ($ok) { exit 0 } else { exit 2 }
