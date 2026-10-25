param(
    [string]$KeyDir,
    [switch]$Force,
    [string]$Passphrase,        # for automated tests only; normally you are asked for it
    [string]$PublicOut          # for automated tests only
)

# ONE-TIME: creates the key that signs your updates.
#   - the PRIVATE key is encrypted with a passphrase you choose and saved in $KeyDir (default: Documents\ToolsBySanat-signing).
#     Copy that file to a USB stick as a backup and keep the passphrase somewhere safe. If it is lost, no update can
#     be signed any more and every PC needs a new loader.
#   - the PUBLIC key is saved in keys\update-public.xml. It is safe to publish; the loader contains a copy.
# Run it yourself in PowerShell (it asks for the passphrase; nobody else should see it).

. (Join-Path $PSScriptRoot 'signing-common.ps1')
if (-not $KeyDir) { $KeyDir = $script:DefaultKeyDir }
$privFile = Join-Path $KeyDir 'update-private.enc'
$pubFile = if ($PublicOut) { $PublicOut } else { $script:PublicKeyFile }

if ((Test-Path $privFile) -and -not $Force) { throw "A signing key already exists at $privFile. Use -Force only if you really want to replace it (every installed loader would then reject your updates)." }

if (-not $Passphrase) {
    $p1 = Read-Secret 'Choose a passphrase for the signing key (at least 12 characters)'
    $p2 = Read-Secret 'Type it again'
    if ($p1 -ne $p2) { throw 'The two passphrases differ. Nothing was created.' }
    $Passphrase = $p1
}
if ($Passphrase.Length -lt 12) { throw 'Passphrase too short (use at least 12 characters).' }

$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(3072)
$privXml = $rsa.ToXmlString($true); $pubXml = $rsa.ToXmlString($false)

New-Item -ItemType Directory -Force $KeyDir | Out-Null
New-Item -ItemType Directory -Force (Split-Path $pubFile) | Out-Null
Set-Content $privFile (Protect-PrivateKey $privXml $Passphrase) -Encoding ascii
Set-Content $pubFile $pubXml -Encoding ascii -NoNewline

# prove the round trip before telling the user it worked
$back = Unprotect-PrivateKey (Get-Content $privFile -Raw) $Passphrase
if ($back -ne $privXml) { throw 'Self-test failed: the encrypted key does not decrypt to the same key.' }

Write-Host ''
Write-Host 'Signing key created.' -ForegroundColor Green
Write-Host "  Private key (encrypted): $privFile"
Write-Host "  Public key             : $pubFile"
Write-Host ''
Write-Host 'NOW: 1) copy the private key file to a USB stick (backup).  2) write the passphrase down somewhere safe.'
Write-Host '     3) tell Claude the public key is ready (the public key file is safe to share; the private one is not).'
