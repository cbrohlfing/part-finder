param(
    [string]$ToVersion
)

# ==========================================================
# Part Finder - Make Release (Working -> Release)
#
# Source:
#   C:\Scripts\PartFinder\working\Part-Finder.ps1
#   C:\Scripts\PartFinder\working\part_finder_settings.json (optional)
#
# Destination (created):
#   C:\Scripts\PartFinder\vX.Y\
#     Part-Finder.ps1
#     part_finder_settings.json (optional)
#     Run Part Finder vX.Y.lnk
#
# Version rule going forward: vX.Y only.
# Legacy folders vX.Y.Z are still detected as "latest" when
# auto-bumping, but new releases are always vX.Y.
# ==========================================================

$root = "C:\Scripts\PartFinder"
$workingDir = Join-Path $root "working"
$workingPs1 = Join-Path $workingDir "Part-Finder.ps1"
$workingCfg = Join-Path $workingDir "part_finder_settings.json"

function Bump-Minor([string]$vText) {
    if ($vText -notmatch '^v(\d+)\.(\d+)$') { throw "Expected vX.Y, got $vText" }
    $maj = [int]$matches[1]
    $min = [int]$matches[2]
    return ("v{0}.{1}" -f $maj, ($min + 1))
}

function Parse-Version([string]$name) {
    # Accept v1.2 OR v1.2.3 (legacy)
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

function Get-LatestVersionDir([string]$rootPath) {
    $vers = Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction Stop |
    Where-Object { $_.Name -like "v*" } |
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
        throw "ToVersion must be in vX.Y format going forward (example: v1.3). You provided: $v"
    }
}

function New-Shortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$Arguments,
        [string]$WorkingDirectory,
        [string]$Description
    )

    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($ShortcutPath)
    $sc.TargetPath = $TargetPath
    $sc.Arguments = $Arguments
    if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
    if ($Description) { $sc.Description = $Description }
    $sc.Save()
}

if (!(Test-Path -LiteralPath $workingPs1)) {
    throw "Missing working script: $workingPs1"
}

$latest = Get-LatestVersionDir $root
if ($null -eq $latest) {
    # If no versions exist yet, default starting point
    $latest = [PSCustomObject]@{ Major = 1; Minor = 0; Patch = 0; Text = "v1.0" }
}

if ([string]::IsNullOrWhiteSpace($ToVersion)) {
    $ToVersion = Next-VersionText $latest
}
Assert-ToVersion $ToVersion

$toDir = Join-Path $root $ToVersion
$dstPs1 = Join-Path $toDir "Part-Finder.ps1"
$dstCfg = Join-Path $toDir "part_finder_settings.json"

if (Test-Path $toDir) { throw "Target folder already exists: $toDir" }

New-Item -ItemType Directory -Path $toDir -Force | Out-Null
Copy-Item -LiteralPath $workingPs1 -Destination $dstPs1 -Force

# Copy settings forward if present (optional)
if (Test-Path -LiteralPath $workingCfg) {
    Copy-Item -LiteralPath $workingCfg -Destination $dstCfg -Force
}

# Update version strings inside the RELEASE PS1 (targeted):
# 1) $script:Version = "vX.Y"
# 2) Header comment line: "# Version: vX.Y ..."
$raw = Get-Content -LiteralPath $dstPs1 -Raw

if ($raw -notmatch '(?m)^#\s*Version:\s*v\d+\.\d+(?:\.\d+)?') {
    throw "Release file missing '# Version:' header line: $dstPs1"
}

if ($raw -notmatch '\$script:Version\s*=') {
    throw "Release file missing `$script:Version assignment: $dstPs1"
}

# 1) Update $script:Version = "..."
$raw = $raw -replace '(\$script:Version\s*=\s*")[^"]+(")', "`$1$ToVersion`$2"

# 2) Update header comment "# Version: vX.Y ..." (only that line)
# Keeps the rest of the comment (like "(Layout fix + sortable columns)")
$raw = $raw -replace '(?m)^(#\s*Version:\s*)v\d+\.\d+(?:\.\d+)?', "`$1$ToVersion"

Set-Content -LiteralPath $dstPs1 -Value $raw -Encoding UTF8


# Parse test (hard fail if broken)
[void][ScriptBlock]::Create((Get-Content -LiteralPath $dstPs1 -Raw))

"OK: Released WORKING -> $ToVersion"
"OK: Parse passed for $dstPs1"

# Create launcher shortcut in new folder
$psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$lnkPath = Join-Path $toDir "Run Part Finder $ToVersion.lnk"
$args = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Normal -File `"$dstPs1`""

New-Shortcut -ShortcutPath $lnkPath `
    -TargetPath $psExe `
    -Arguments $args `
    -WorkingDirectory $toDir `
    -Description "Launch Part Finder ($ToVersion)"

"OK: Shortcut created: $lnkPath"

# After releasing, bump WORKING to next minor so it shows "in development"
$nextWorking = Bump-Minor $ToVersion

$wraw = Get-Content -LiteralPath $workingPs1 -Raw

if ($wraw -notmatch '(?m)^#\s*Version:\s*v\d+\.\d+(?:\.\d+)?') {
    throw "Working file missing '# Version:' header line: $workingPs1"
}

if ($wraw -notmatch '\$script:Version\s*=') {
    throw "Working file missing `$script:Version assignment: $workingPs1"
}

# 1) Update $script:Version = "..."
$wraw = $wraw -replace '(\$script:Version\s*=\s*")[^"]+(")', "`$1$nextWorking`$2"

# 2) Update header comment "# Version: vX.Y ..."
$wraw = $wraw -replace '(?m)^(#\s*Version:\s*)v\d+\.\d+(?:\.\d+)?', "`$1$nextWorking"

Set-Content -LiteralPath $workingPs1 -Value $wraw -Encoding UTF8

"OK: Working bumped to $nextWorking"
