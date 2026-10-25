param(
    [Parameter(Mandatory)][string]$Version
)

# Builds a release. Does NOT commit, push or publish anything: it prints the commands to run.
#   1. build.ps1 -Version        -> build\ToolsBySanat.xlam + version.json (sha256, url)
#   2. build-loader.ps1          -> build\ToolsBySanatLoader.xlam
#   3. build-setup.ps1           -> dist\ToolsBySanat-Setup-<version>.exe  (what new users get; Inno Setup wizard)
#   (Install.bat + install.ps1 still work for installing straight from this folder.)

$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw "Version must look like 1.2.3" }
if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Close all Excel windows first.' }

$root = $PSScriptRoot
& (Join-Path $root 'build.ps1') -Version $Version
& (Join-Path $root 'build-loader.ps1')

& (Join-Path $root 'build-setup.ps1')
$setup = Join-Path $root "dist\ToolsBySanat-Setup-$Version.exe"

$cfg = Get-Content (Join-Path $root 'updater.config.json') -Raw | ConvertFrom-Json
Write-Host ''
Write-Host "Built $setup" -ForegroundColor Green
if (-not $cfg.repo) {
    Write-Host 'updater.config.json has no "repo" yet, so version.json has no url and the loader will not check for updates.' -ForegroundColor Yellow
    Write-Host 'Set it to "owner/name" and run this script again before publishing.'
    exit 0
}
Write-Host 'Next (run yourself, after checking the build works):'
Write-Host "  git add -A; git commit -m 'Release v$Version'; git push"
Write-Host "  gh release create v$Version build\ToolsBySanat.xlam $setup --title v$Version --notes 'v$Version'"
Write-Host '  (the .xlam is what the updater downloads; the Setup .exe is what new users run once)'
Write-Host 'Users on silent auto-update get it at their next Excel start after the release exists AND version.json is pushed.'
