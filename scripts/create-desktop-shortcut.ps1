[CmdletBinding()]
param(
  [string]$ShortcutName = 'DM Helper',
  [string]$ShortcutPath,
  [string]$TargetPath,
  [switch]$PreferSourceLauncher
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceLauncher = Join-Path $PSScriptRoot 'launch-dm-helper.cmd'
$portableExe = Join-Path $repoRoot 'dist\DM Helper 0.1.0.exe'
$unpackedExe = Join-Path $repoRoot 'dist\win-unpacked\DM Helper.exe'
$shortcutIcon = Join-Path $repoRoot 'build\icon.ico'
$electronExe = Join-Path $repoRoot 'node_modules\electron\dist\electron.exe'

if (-not $ShortcutPath) {
  $desktop = [Environment]::GetFolderPath('Desktop')
  $ShortcutPath = Join-Path $desktop "$ShortcutName.lnk"
}

if (-not $TargetPath) {
  $candidates = @()

  if ($PreferSourceLauncher) {
    $candidates += $sourceLauncher
  }

  $candidates += @(
    $portableExe,
    $unpackedExe,
    $sourceLauncher
  )

  $TargetPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

if (-not $TargetPath -or -not (Test-Path $TargetPath)) {
  throw "Could not find a launch target. Expected one of: $portableExe, $unpackedExe, or $sourceLauncher"
}

$workingDirectory = Split-Path -Parent $TargetPath
if ($TargetPath -eq $sourceLauncher) {
  $workingDirectory = $repoRoot
}

$iconLocation = if (Test-Path $shortcutIcon) {
  $shortcutIcon
} elseif (Test-Path $electronExe) {
  $electronExe
} elseif (Test-Path $portableExe) {
  $portableExe
} elseif (Test-Path $unpackedExe) {
  $unpackedExe
} else {
  "$env:SystemRoot\System32\SHELL32.dll,220"
}

$shortcutDir = Split-Path -Parent $ShortcutPath
if ($shortcutDir -and -not (Test-Path $shortcutDir)) {
  New-Item -ItemType Directory -Path $shortcutDir -Force | Out-Null
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $TargetPath
$shortcut.WorkingDirectory = $workingDirectory
$shortcut.IconLocation = $iconLocation
$shortcut.Description = 'Launch DM Helper'
$shortcut.Save()

Write-Host "Shortcut created: $ShortcutPath"
Write-Host "Target: $TargetPath"
