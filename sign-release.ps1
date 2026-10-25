param(
    [string]$KeyFile,
    [string]$VersionFile = (Join-Path $PSScriptRoot 'version.json'),
    [string]$AddIn = (Join-Path $PSScriptRoot 'build\ToolsBySanat.xlam'),
    [string]$PublicKey,
    [string]$Passphrase         # automated tests only; normally you are asked for it
)

# PER RELEASE: signs version.json with your private key. Run it yourself after release.ps1 and before publishing.
# It refuses to sign if the add-in file no longer matches the hash in version.json, so a stale build is never signed.

. (Join-Path $PSScriptRoot 'signing-common.ps1')
if (-not $KeyFile) { $KeyFile = Join-Path $script:DefaultKeyDir 'update-private.enc' }
if (-not $PublicKey) { $PublicKey = $script:PublicKeyFile }
if (-not (Test-Path $KeyFile)) { throw "Signing key not found: $KeyFile (plug in the USB stick, or run new-signing-key.ps1 once)." }

$m = Get-Content $VersionFile -Raw | ConvertFrom-Json
if (-not $m.version -or -not $m.sha256 -or -not $m.url) { throw 'version.json is missing version, sha256 or url. Run release.ps1 first (and make sure updater.config.json has the repo).' }
if (Test-Path $AddIn) {
    $h = (Get-FileHash $AddIn -Algorithm SHA256).Hash.ToLower()
    if ($h -ne $m.sha256) { throw "The add-in file ($AddIn) does not match the hash in version.json. Rebuild with release.ps1." }
}
Write-Host "About to sign version $($m.version)"
Write-Host "  sha256: $($m.sha256)"
Write-Host "  url   : $($m.url)"

if (-not $Passphrase) { $Passphrase = Read-Secret 'Signing passphrase' }
$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
$rsa.FromXmlString((Unprotect-PrivateKey (Get-Content $KeyFile -Raw) $Passphrase))
$sig = $rsa.SignData([Text.Encoding]::ASCII.GetBytes((Get-ManifestMessage $m)), 'SHA256')

$out = [ordered]@{ version = $m.version; sha256 = $m.sha256; url = $m.url; signature = (ConvertTo-Hex $sig) }
$check = [pscustomobject]$out
if (-not (Test-Path $PublicKey)) { throw "Public key file not found: $PublicKey" }
if (-not (Test-ManifestSignature $check (Get-Content $PublicKey -Raw))) { throw 'The signature does not verify against keys\update-public.xml. This private key does not belong to the public key built into the loader.' }

$json = $out | ConvertTo-Json
$saved = $false
for ($try = 1; $try -le 10 -and -not $saved; $try++) {
    try { [IO.File]::WriteAllText($VersionFile, $json, (New-Object Text.UTF8Encoding($true))); $saved = $true }
    catch { if ($try -eq 10) { throw "Could not save $VersionFile because another program has it open (an editor, Notepad, antivirus or sync tool). Close it and run this again. Nothing was changed." }; Start-Sleep -Milliseconds 700 }
}
Write-Host ''
Write-Host "Signed and verified. version.json now carries the signature for $($m.version)." -ForegroundColor Green
Write-Host 'Tell Claude it is signed; the release can be published now.'
