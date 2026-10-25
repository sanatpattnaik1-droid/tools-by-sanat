# Builds dist\ToolsBySanat-Setup-<version>.exe with Inno Setup from build\*.xlam (run build.ps1 and
# build-loader.ps1 first, or use release.ps1 which does everything). Version comes from version.json.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

$iscc = @("$env:ProgramFiles\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw 'Inno Setup 6 not found. Install it with: winget install JRSoftware.InnoSetup' }
foreach ($f in 'build\ToolsBySanat.xlam', 'build\ToolsBySanatLoader.xlam') { if (-not (Test-Path (Join-Path $root $f))) { throw "$f is missing. Run build.ps1 and build-loader.ps1 first." } }

$version = (Get-Content (Join-Path $root 'version.json') -Raw | ConvertFrom-Json).version
Push-Location $root
try { & $iscc "/DMyAppVersion=$version" 'setup.iss'; if ($LASTEXITCODE -ne 0) { throw "ISCC failed ($LASTEXITCODE)" } } finally { Pop-Location }

$out = Join-Path $root "dist\ToolsBySanat-Setup-$version.exe"
"Built $out ($([int]((Get-Item $out).Length / 1KB)) KB)"
