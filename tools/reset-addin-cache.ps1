<#
.SYNOPSIS
    Checks your Word add-in catalogs and clears Office's add-in cache.

.DESCRIPTION
    Run this with Word closed when an add-in from your shared folder is missing from the
    SHARED FOLDER list, shows an old version, or keeps disappearing from the ribbon.
    It lists your trusted catalogs and the manifests in them (warning about common problems),
    then empties Office's add-in cache so Word re-reads everything next time it starts.

.EXAMPLE
    ./tools/reset-addin-cache.ps1

.EXAMPLE
    ./tools/reset-addin-cache.ps1 -CheckOnly
#>
param([switch]$CheckOnly)

$report = Join-Path $PSScriptRoot 'addin-check.txt'
try { Start-Transcript -Path $report -Force | Out-Null } catch { }

function Step([string]$t) { Write-Host ""; Write-Host "==> $t" -ForegroundColor Cyan }
function Ok([string]$t)   { Write-Host "    $t" -ForegroundColor Green }
function Note([string]$t) { Write-Host "    $t" -ForegroundColor Yellow }

if (-not $CheckOnly -and (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)) {
    Write-Host "Please close every Word window first (or add -CheckOnly to just run the checks)." -ForegroundColor Yellow
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}

# ------------------------------------------------------------ catalogs
Step "Trusted add-in catalogs"
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
$catalogs = @()
if (Test-Path $key) { $catalogs = @(Get-ChildItem $key | ForEach-Object { Get-ItemProperty $_.PSPath }) }
if ($catalogs.Count -eq 0) {
    Note "No catalogs found. In Word: File > Options > Trust Center > Trust Center Settings >"
    Note "Trusted Add-in Catalogs, add your shared folder's \\COMPUTER\share path and tick Show in Menu."
}

$seenUrls = @{}
$seenIds = @{}
foreach ($c in $catalogs) {
    $url = "$($c.Url)"
    $show = (([int]$c.Flags -band 1) -eq 1)
    if ($show) { $showText = 'yes' } else { $showText = 'NO' }
    Write-Host ""
    Write-Host "    $url   (Show in Menu: $showText)"
    if (-not $show) { Note "  Tick 'Show in Menu' for this catalog, or remove it if you no longer use it." }
    if ($url -notmatch '^\\\\') { Note "  This is not a network path. Use the \\COMPUTER\share path shown on the folder's Sharing tab." }
    if ($seenUrls.ContainsKey($url.ToLower())) { Note "  This catalog is listed twice - remove one of them." }
    $seenUrls[$url.ToLower()] = $true

    if (-not (Test-Path -LiteralPath $url)) { Note "  Word can't reach this folder. Check it is still shared."; continue }
    $files = @(Get-ChildItem -LiteralPath $url -Filter *.xml -File)
    if ($files.Count -eq 0) { Note "  No manifest (.xml) files in this folder." }
    foreach ($file in $files) {
        try {
            [xml]$doc = Get-Content -LiteralPath $file.FullName -Raw
            $id = "$($doc.OfficeApp.Id)".Trim()
            $name = "$($doc.OfficeApp.DisplayName.DefaultValue)"
        } catch {
            Note "  $($file.Name): not a valid XML file - Word will ignore it."
            continue
        }
        if (-not $id) { Note "  $($file.Name): no <Id> found - Word will ignore it."; continue }
        Write-Host "      - $($file.Name): $name  ($id)"
        if ($seenIds.ContainsKey($id)) { Note "    Same Id as $($seenIds[$id]) - Word will only show one of them." }
        else { $seenIds[$id] = $file.Name }
    }
}

# ------------------------------------------------------------ cache
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
if ($CheckOnly) {
    Step "Cache left as it is (-CheckOnly)"
} elseif (Test-Path -LiteralPath $wef) {
    Step "Clearing Office's add-in cache"
    Get-ChildItem -LiteralPath $wef -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Ok "Cleared $wef"
} else {
    Step "No add-in cache to clear"
}

Step "Next"
Write-Host "    1. Open Word."
Write-Host "    2. Home > Add-ins > Advanced (or More Add-ins) > SHARED FOLDER, then click Refresh."
Write-Host "    3. Add each add-in once. They should now stay on the ribbon when Word restarts."
Write-Host ""
Write-Host "    (This report was also saved to tools\addin-check.txt)"
try { Stop-Transcript | Out-Null } catch { }
