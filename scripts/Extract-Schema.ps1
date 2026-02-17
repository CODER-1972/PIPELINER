param(
  [Parameter(Mandatory=$false)][string]$PayloadPath = "C:\Temp\_raw\payload_invalid.json",
  [Parameter(Mandatory=$false)][string]$SchemaOutPath = "C:\Temp\schema_only.json"
)

if (-not (Test-Path -LiteralPath $PayloadPath)) {
  Write-Error "Ficheiro não encontrado: $PayloadPath"
  exit 1
}

try {
  $obj = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 100
} catch {
  Write-Error "Payload inválido; extração de schema indisponível: $($_.Exception.Message)"
  exit 2
}

$schema = $obj.text.format.schema
if ($null -eq $schema) {
  Write-Error "Schema não encontrado em text.format.schema"
  exit 3
}

$dir = Split-Path -Parent $SchemaOutPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$schema | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $SchemaOutPath -Encoding UTF8
Write-Host "Schema extraído para: $SchemaOutPath"
