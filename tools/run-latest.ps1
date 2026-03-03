param(
  [string]$Root = "C:\Scripts\PartFinder"
)

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

$latest = Get-ChildItem -LiteralPath $Root -Directory -ErrorAction Stop |
ForEach-Object { Parse-Version $_.Name } |
Where-Object { $_ -ne $null } |
Sort-Object Major, Minor, Patch |
Select-Object -Last 1

if (-not $latest) {
  throw "No release version folders found in $Root (expected v1.2, v1.2.3, v1.3, etc.)"
}

$verDir = Join-Path $Root $latest.Text
$ps1 = Join-Path $verDir "Part-Finder.ps1"
if (!(Test-Path -LiteralPath $ps1)) { throw "Missing: $ps1" }

Push-Location $verDir
try {
  Write-Host "Running latest RELEASE: $($latest.Text)"
  & $ps1
} finally {
  Pop-Location
}
