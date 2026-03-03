param(
  # Release version in vX.Y format (example: v1.6). If omitted, auto-bumps minor from latest release.
  [string]$ToVersion
)

# ==========================================================
# Part Finder - Professional Dev -> Release
#
# Repo layout:
#   <repo>\src\Part-Finder.ps1                 (DEV source of truth)
#   <repo>\tools\run-latest.ps1               (launcher shipped to users)
#   <repo>\releases\vX.Y\Part-Finder.ps1      (release snapshots)
#   <repo>\dist\PartFinder_vX.Y.zip           (zip to send to coworkers)
#
# Install layout on coworker machines:
#   C:\Scripts\PartFinder\run-latest.ps1
#   C:\Scripts\PartFinder\vX.Y\Part-Finder.ps1
#
# Notes:
# - run-latest picks the HIGHEST version folder (v1.10 > v1.9, and v1.2.3 supported).
# - This script updates the release PS1 version header + $script:Version.
# - After releasing, DEV is bumped to next minor.
# ==========================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve repo root from this script location: <repo>\tools\make-release-partfinder.ps1
$repoRoot = Split-Path -Parent $PSScriptRoot

$srcPs1 = Join-Path $repoRoot 'src\Part-Finder.ps1'
$launcherPs1 = Join-Path $repoRoot 'tools\run-latest.ps1'
$releasesDir = Join-Path $repoRoot 'releases'
$distDir = Join-Path $repoRoot 'dist'

function Parse-Version([string]$name) {
  if ($name -match '^v(\d+)\.(\d+)(?:\.(\d+))?$') {
    return [PSCustomObject]@{
      Major = [int]$matches[1]
      Minor = [int]$matches[2]
      Patch = if ($matches[3]) { [int]$matches[3] } else { 0 }
      Text  = $name
    }
  }
  return $null
}

function Get-LatestRelease([string]$root) {
  if (!(Test-Path -LiteralPath $root)) { return $null }
  $vers = Get-ChildItem -LiteralPath $root -Directory -ErrorAction Stop |
  ForEach-Object { Parse-Version $_.Name } |
  Where-Object { $_ -ne $null } |
  Sort-Object Major, Minor, Patch

  if (-not $vers -or $vers.Count -eq 0) { return $null }
  return $vers[-1]
}

function Next-VersionText($v) {
  # bump minor: v1.2(.3) -> v1.3
  return ("v{0}.{1}" -f $v.Major, ($v.Minor + 1))
}

function Assert-ToVersion([string]$v) {
  if ($v -notmatch '^v\d+\.\d+$') {
    throw "ToVersion must be vX.Y (example: v1.6). You provided: $v"
  }
}

function Bump-Minor([string]$vText) {
  if ($vText -notmatch '^v(\d+)\.(\d+)$') { throw "Expected vX.Y, got $vText" }
  $maj = [int]$matches[1]
  $min = [int]$matches[2]
  return ("v{0}.{1}" -f $maj, ($min + 1))
}

if (!(Test-Path -LiteralPath $srcPs1)) { throw "Missing DEV script: $srcPs1" }
if (!(Test-Path -LiteralPath $launcherPs1)) { throw "Missing launcher script: $launcherPs1" }

New-Item -ItemType Directory -Path $releasesDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$latest = Get-LatestRelease $releasesDir
if ($null -eq $latest) {
  $latest = [PSCustomObject]@{ Major = 1; Minor = 0; Patch = 0; Text = 'v1.0' }
}

if ([string]::IsNullOrWhiteSpace($ToVersion)) {
  $ToVersion = Next-VersionText $latest
}
Assert-ToVersion $ToVersion

$toReleaseDir = Join-Path $releasesDir $ToVersion
$releasePs1 = Join-Path $toReleaseDir 'Part-Finder.ps1'

if (Test-Path -LiteralPath $toReleaseDir) {
  throw "Release already exists: $toReleaseDir"
}

New-Item -ItemType Directory -Path $toReleaseDir -Force | Out-Null
Copy-Item -LiteralPath $srcPs1 -Destination $releasePs1 -Force

# Update version strings inside the RELEASE PS1 (targeted):
# 1) $script:Version = "vX.Y"
# 2) Header comment line: "# Version: vX.Y ..."
$raw = Get-Content -LiteralPath $releasePs1 -Raw

if ($raw -notmatch '(?m)^#\s*Version:\s*v\d+\.\d+(?:\.\d+)?') {
  throw "Release file missing '# Version:' header line: $releasePs1"
}
if ($raw -notmatch '\$script:Version\s*=') {
  throw "Release file missing `$script:Version assignment: $releasePs1"
}

$raw = $raw -replace '(\$script:Version\s*=\s*")[^"]+(\")', "`$1$ToVersion`$2"
$raw = $raw -replace '(?m)^(#\s*Version:\s*)v\d+\.\d+(?:\.\d+)?', "`$1$ToVersion"

Set-Content -LiteralPath $releasePs1 -Value $raw -Encoding UTF8

# Parse test (hard fail if broken)
[void][ScriptBlock]::Create((Get-Content -LiteralPath $releasePs1 -Raw))

Write-Host "OK: Release created -> releases\\$ToVersion"

# Build coworker install package zip
$tmpPkgRoot = Join-Path $distDir ("PartFinder_{0}" -f $ToVersion)
if (Test-Path -LiteralPath $tmpPkgRoot) { Remove-Item -LiteralPath $tmpPkgRoot -Recurse -Force }
New-Item -ItemType Directory -Path $tmpPkgRoot -Force | Out-Null

# Package layout:
#   <zip>\run-latest.ps1
#   <zip>\vX.Y\Part-Finder.ps1
Copy-Item -LiteralPath $launcherPs1 -Destination (Join-Path $tmpPkgRoot 'run-latest.ps1') -Force
New-Item -ItemType Directory -Path (Join-Path $tmpPkgRoot $ToVersion) -Force | Out-Null
Copy-Item -LiteralPath $releasePs1 -Destination (Join-Path (Join-Path $tmpPkgRoot $ToVersion) 'Part-Finder.ps1') -Force

# Optional: include a settings TEMPLATE (not user settings). Use .example extension so it can be tracked even if *.json is ignored.
$settingsExample = Join-Path $repoRoot 'assets\part_finder_settings.json.example'
if (Test-Path -LiteralPath $settingsExample) {
  Copy-Item -LiteralPath $settingsExample -Destination (Join-Path $tmpPkgRoot $ToVersion 'part_finder_settings.json') -Force
}

$zipPath = Join-Path $distDir ("PartFinder_{0}.zip" -f $ToVersion)
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmpPkgRoot, $zipPath)

Write-Host "OK: Package created -> $zipPath"

# Bump DEV to next minor so it clearly shows 'in development'
$nextDev = Bump-Minor $ToVersion
$devRaw = Get-Content -LiteralPath $srcPs1 -Raw

if ($devRaw -notmatch '(?m)^#\s*Version:\s*v\d+\.\d+(?:\.\d+)?') {
  throw "DEV file missing '# Version:' header line: $srcPs1"
}
if ($devRaw -notmatch '\$script:Version\s*=') {
  throw "DEV file missing `$script:Version assignment: $srcPs1"
}

$devRaw = $devRaw -replace '(\$script:Version\s*=\s*")[^"]+(\")', "`$1$nextDev`$2"
$devRaw = $devRaw -replace '(?m)^(#\s*Version:\s*)v\d+\.\d+(?:\.\d+)?', "`$1$nextDev"

Set-Content -LiteralPath $srcPs1 -Value $devRaw -Encoding UTF8
Write-Host "OK: DEV bumped to $nextDev"

Write-Host "\nNext steps:"
Write-Host "  Write-Host "  git add src/Part-Finder.ps1 releases/$ToVersion""
Write-Host "  git commit -m \"release: $ToVersion\""
Write-Host "  git tag $ToVersion"
Write-Host "  git push && git push origin $ToVersion"
