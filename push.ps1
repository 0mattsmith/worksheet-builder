<#
.SYNOPSIS
    Publish Worksheet Builder to GitHub and host the add-in on GitHub Pages.

.DESCRIPTION
    First run:  creates the GitHub repository, turns on GitHub Pages, pushes the code and
                makes manifest-hosted.xml (the manifest you add to Word) with your site address.
    Later runs: commits your changes with the message you give and pushes them. GitHub
                re-publishes the add-in website automatically.

.EXAMPLE
    ./push.ps1 "update information"

.EXAMPLE
    ./push.ps1 "first version" -RepoName my-worksheets -Open

.PARAMETER Message
    Commit message describing what changed.
.PARAMETER RepoName
    Name of the GitHub repository (default: worksheet-builder).
.PARAMETER Private
    Create a private repository (GitHub Pages on private repositories needs a paid GitHub plan).
.PARAMETER Force
    Replace whatever is already in the GitHub repository with this folder (use once if you
    previously uploaded files through the GitHub website).
.PARAMETER NoWait
    Do not wait for the website to finish publishing.
.PARAMETER Open
    Open the add-in page in your browser when publishing has finished.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Message = "Update Worksheet Builder",
    [string]$RepoName = "worksheet-builder",
    [switch]$Private,
    [switch]$Force,
    [switch]$NoWait,
    [switch]$Open
)

# Native commands (git, gh) report failure through $LASTEXITCODE, which we check ourselves.
$ErrorActionPreference = 'Continue'
Set-Location -LiteralPath $PSScriptRoot

function Step([string]$text) { Write-Host ""; Write-Host "==> $text" -ForegroundColor Cyan }
function Ok([string]$text)   { Write-Host "    $text" -ForegroundColor Green }
function Note([string]$text) { Write-Host "    $text" -ForegroundColor Yellow }
function Fail([string]$text) { Write-Host ""; Write-Host "ERROR: $text" -ForegroundColor Red; exit 1 }
function Check([string]$what) { if ($LASTEXITCODE -ne 0) { Fail "$what failed (exit code $LASTEXITCODE)." } }
function Write-Utf8([string]$path, [string]$text) {
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

# ---------------------------------------------------------------- 1. Tools and account
Step "Checking git and the GitHub CLI"
foreach ($tool in @('git', 'gh')) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Fail "$tool is not installed or not on your PATH. Install it, open a new PowerShell window and try again."
    }
}
if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'manifest.xml'))) {
    Fail "manifest.xml not found. Put push.ps1 in the worksheet-builder folder (next to manifest.xml) and run it from there."
}
gh auth status *> $null
if ($LASTEXITCODE -ne 0) { Fail "The GitHub CLI is not signed in. Run:  gh auth login" }

$owner = gh api user --jq .login
if ($LASTEXITCODE -ne 0 -or -not $owner) { Fail "Could not read your GitHub account name." }
$owner = "$owner".Trim()
$repo = "$owner/$RepoName"
if ($RepoName -ieq "$owner.github.io") { $siteUrl = "https://$($owner.ToLower()).github.io/" }
else { $siteUrl = "https://$($owner.ToLower()).github.io/$RepoName/" }
Ok "Signed in as $owner"

# ---------------------------------------------------------------- 2. Files
Step "Preparing files"

$gitignore = Join-Path $PSScriptRoot '.gitignore'
if (-not (Test-Path -LiteralPath $gitignore)) {
    Write-Utf8 $gitignore "node_modules/`n*.log`n.DS_Store`nThumbs.db`ndesktop.ini`n"
    Ok "Created .gitignore"
}

# GitHub Actions workflow that publishes the web folder to GitHub Pages on every push.
$wfDir = Join-Path (Join-Path $PSScriptRoot '.github') 'workflows'
$wfFile = Join-Path $wfDir 'pages.yml'
if (-not (Test-Path -LiteralPath $wfFile)) {
    New-Item -ItemType Directory -Force -Path $wfDir | Out-Null
    $workflow = @'
name: Publish add-in to GitHub Pages

on:
  push:
    branches: [main, master]
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: true

jobs:
  deploy:
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/checkout@v5
      - uses: actions/upload-pages-artifact@v5
        with:
          path: web
      - id: deployment
        uses: actions/deploy-pages@v5
'@
    Write-Utf8 $wfFile $workflow
    Ok "Created .github\workflows\pages.yml (publishes the web folder)"
}

# Manifest pointing at the hosted copy. Also published on the site so colleagues can download it.
$manifest = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'manifest.xml'))
$hosted = $manifest.Replace('https://localhost:3000/', $siteUrl)
Write-Utf8 (Join-Path $PSScriptRoot 'manifest-hosted.xml') $hosted
Write-Utf8 (Join-Path (Join-Path $PSScriptRoot 'web') 'manifest.xml') $hosted
Ok "manifest-hosted.xml points at $siteUrl"

# ---------------------------------------------------------------- 3. Commit
Step "Saving your changes"
if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot '.git'))) {
    git init --quiet; Check "git init"
    git symbolic-ref HEAD refs/heads/main
    Ok "Started a new git repository"
}

if (-not (git config user.name)) {
    $displayName = gh api user --jq '.name // .login'
    git config user.name ("$displayName".Trim())
}
if (-not (git config user.email)) {
    $userId = gh api user --jq .id
    git config user.email ("{0}+{1}@users.noreply.github.com" -f "$userId".Trim(), $owner)
}

git add -A; Check "git add"
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
    git commit --quiet -m $Message; Check "git commit"
    Ok "Committed: $Message"
} else {
    Ok "No file changes since the last push"
}
git rev-parse --verify HEAD *> $null
if ($LASTEXITCODE -ne 0) { Fail "There is nothing to push yet." }
$branch = "$(git rev-parse --abbrev-ref HEAD)".Trim()
$headSha = "$(git rev-parse HEAD)".Trim()

# ---------------------------------------------------------------- 4. Repository
Step "Connecting to https://github.com/$repo"
gh repo view $repo *> $null
if ($LASTEXITCODE -ne 0) {
    if ($Private) { $visibility = '--private' } else { $visibility = '--public' }
    gh repo create $repo $visibility --description "Worksheet Builder - a Word add-in for uniform class worksheets"
    Check "Creating the repository"
    Ok "Created the repository"
} else {
    Ok "Repository already exists"
}

$protocol = "$(gh config get git_protocol 2>$null)".Trim()
if ($protocol -eq 'ssh') { $remoteUrl = "git@github.com:$repo.git" }
else {
    $remoteUrl = "https://github.com/$repo.git"
    gh auth setup-git *> $null   # let git use your GitHub CLI sign-in
}
$currentRemote = "$(git remote get-url origin 2>$null)".Trim()
if (-not $currentRemote) {
    git remote add origin $remoteUrl; Check "Adding the GitHub remote"
} elseif ($currentRemote -notmatch [regex]::Escape("$repo") ) {
    Note "Your 'origin' remote points at $currentRemote - switching it to $repo"
    git remote set-url origin $remoteUrl; Check "Updating the GitHub remote"
}

# ---------------------------------------------------------------- 5. GitHub Pages
Step "Turning on GitHub Pages"
$pagesPending = $false
gh api "repos/$repo/pages" *> $null
if ($LASTEXITCODE -ne 0) {
    gh api -X POST "repos/$repo/pages" -f build_type=workflow *> $null
    if ($LASTEXITCODE -eq 0) { Ok "GitHub Pages is on" }
    else { $pagesPending = $true; Note "Will finish this after the first push" }
} else {
    $buildType = "$(gh api "repos/$repo/pages" --jq .build_type)".Trim()
    if ($buildType -ne 'workflow') {
        gh api -X PUT "repos/$repo/pages" -f build_type=workflow *> $null
        Ok "Switched GitHub Pages to publish from the workflow"
    } else {
        Ok "GitHub Pages is already on"
    }
}

# ---------------------------------------------------------------- 6. Push
Step "Pushing to GitHub"
if ($Force) { git push -u --force origin $branch } else { git push -u origin $branch }
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Note "GitHub refused the push. This usually means the repository already has different files,"
    Note "for example ones uploaded through the GitHub website."
    Note "  - To replace them with this folder:   ./push.ps1 `"$Message`" -Force"
    Note "  - Or use a new repository instead:    ./push.ps1 `"$Message`" -RepoName another-name"
    Fail "git push"
}
Ok "Pushed branch $branch"

if ($pagesPending) {
    gh api -X POST "repos/$repo/pages" -f build_type=workflow *> $null
    if ($LASTEXITCODE -eq 0) { Ok "GitHub Pages is on" }
    else {
        if ($Private) { Note "GitHub Pages could not be turned on. Private repositories need a paid GitHub plan for Pages." }
        Fail "Turning on GitHub Pages. Open https://github.com/$repo/settings/pages and choose 'GitHub Actions' as the source."
    }
}

# ---------------------------------------------------------------- 7. Wait for the website
if (-not $NoWait) {
    Step "Publishing the add-in website (usually about a minute)"
    $run = $null
    for ($i = 0; $i -lt 30 -and -not $run; $i++) {
        Start-Sleep -Seconds 3
        $json = gh run list --repo $repo --workflow pages.yml --limit 10 --json databaseId,headSha,status,conclusion 2>$null
        if ($LASTEXITCODE -eq 0 -and $json) {
            $runs = ($json -join "`n") | ConvertFrom-Json
            foreach ($r in $runs) { if ($r.headSha -eq $headSha) { $run = $r; break } }
        }
    }
    if (-not $run) {
        Note "Could not find the publishing run yet. Check progress at https://github.com/$repo/actions"
    } else {
        gh run watch $run.databaseId --repo $repo --exit-status
        if ($LASTEXITCODE -ne 0 -and $pagesPending) {
            Note "The first run started before GitHub Pages was ready - running it again."
            gh run rerun $run.databaseId --repo $repo *> $null
            Start-Sleep -Seconds 5
            gh run watch $run.databaseId --repo $repo --exit-status
        }
        if ($LASTEXITCODE -ne 0) {
            Fail "Publishing did not finish. See https://github.com/$repo/actions for details."
        }
        Ok "Website published"
    }
}

# ---------------------------------------------------------------- 8. Summary
Step "Done"
Write-Host "    Add-in website : ${siteUrl}taskpane.html"
Write-Host "    Manifest file  : $(Join-Path $PSScriptRoot 'manifest-hosted.xml')"
Write-Host "    For colleagues : ${siteUrl}manifest.xml"
Write-Host ""
Write-Host "    First time only: add manifest-hosted.xml to Word (see README, 'Add it to Word')."
Write-Host '    After that just run  ./push.ps1 "what you changed"  - Word loads the new version'
Write-Host "    the next time you open the add-in (allow a few minutes for GitHub's cache)."
if ($Open) { Start-Process "${siteUrl}taskpane.html" }
