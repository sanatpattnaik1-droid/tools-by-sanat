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
# Building uses its own hidden Excel, so an Excel you have open is fine. Only the build\ output files must be free
# (build.ps1 / build-loader.ps1 stop with a clear message if they are locked).
if (Get-Process excel -ErrorAction SilentlyContinue) { Write-Host 'Excel is open: that is fine, the build uses its own hidden copy.' -ForegroundColor DarkGray }

$root = $PSScriptRoot
& (Join-Path $root 'build.ps1') -Version $Version
& (Join-Path $root 'build-loader.ps1')

& (Join-Path $root 'build-setup.ps1')
$setup = Join-Path $root "dist\ToolsBySanat-Setup-$Version.exe"

# checksums people can compare with what they downloaded (also pasted into the release notes)
$sums = Join-Path $root 'dist\SHA256SUMS.txt'
@(
    "$((Get-FileHash $setup -Algorithm SHA256).Hash.ToLower())  ToolsBySanat-Setup-$Version.exe",
    "$((Get-FileHash (Join-Path $root 'build\ToolsBySanat.xlam') -Algorithm SHA256).Hash.ToLower())  ToolsBySanat.xlam"
) | Set-Content $sums -Encoding ascii

$cfg = Get-Content (Join-Path $root 'updater.config.json') -Raw | ConvertFrom-Json
Write-Host ''
Write-Host "Built $setup" -ForegroundColor Green
if (-not $cfg.repo) {
    Write-Host 'updater.config.json has no "repo" yet, so version.json has no url and the loader will not check for updates.' -ForegroundColor Yellow
    Write-Host 'Set it to "owner/name" and run this script again before publishing.'
    exit 0
}
Write-Host ''
Write-Host 'Next:' -ForegroundColor Cyan
Write-Host '  1. YOU sign it (needs your passphrase, so run it yourself):   .\sign-release.ps1'
Write-Host '  2. Check the signature:                                        .\verify-manifest.ps1'
Write-Host "  3. Publish the release FIRST, then push version.json. The repo uses immutable releases, so attach the files"
Write-Host "     to a DRAFT and then publish it (a published release can no longer be changed):"
Write-Host "       gh release create v$Version build\ToolsBySanat.xlam $setup --draft --title v$Version --notes-file <notes incl. dist\SHA256SUMS.txt>"
Write-Host "       gh release edit v$Version --draft=false"
Write-Host "       git add -A; git commit -m 'Release v$Version'; git push"
Write-Host '  (the .xlam is what the updater downloads; the Setup .exe is what new users run once)'
Write-Host 'Unsigned or wrongly signed updates are ignored by every PC that has the signature-checking loader.'
