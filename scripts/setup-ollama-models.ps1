$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modelDir = Join-Path $repoRoot 'ollama-models'

function Build-Model {
  param(
    [string]$Tag,
    [string]$FileName
  )

  $path = Join-Path $modelDir $FileName
  if (-not (Test-Path $path)) {
    throw "Missing modelfile: $path"
  }

  Write-Host "Building $Tag from $FileName"
  ollama create $Tag -f $path
}

Build-Model 'gpt-oss-20b-optimized:latest' 'gpt-oss-20b-optimized.modelfile'
Build-Model 'lorebound-pf2e:latest' 'lorebound-pf2e.modelfile'
Build-Model 'lorebound-pf2e-fast:latest' 'lorebound-pf2e-fast.modelfile'
Build-Model 'lorebound-pf2e-ultra-fast:latest' 'lorebound-pf2e-ultra-fast.modelfile'
Build-Model 'lorebound-pf2e-pure:latest' 'lorebound-pf2e-pure.modelfile'

Write-Host ''
Write-Host 'Done. You can now pick these models inside DM Helper.'
