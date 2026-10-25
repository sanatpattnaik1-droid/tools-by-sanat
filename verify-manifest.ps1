param(
    [string]$VersionFile = (Join-Path $PSScriptRoot 'version.json'),
    [string]$PublicKey
)

# Checks that version.json is correctly signed (uses only the PUBLIC key; safe to run anywhere, also on the published copy).

. (Join-Path $PSScriptRoot 'signing-common.ps1')
if (-not $PublicKey) { $PublicKey = $script:PublicKeyFile }
$m = Get-Content $VersionFile -Raw | ConvertFrom-Json
if (Test-ManifestSignature $m (Get-Content $PublicKey -Raw)) { "OK: version $($m.version) is signed by the owner's key."; exit 0 }
Write-Host "NOT SIGNED CORRECTLY: $VersionFile" -ForegroundColor Red; exit 1
