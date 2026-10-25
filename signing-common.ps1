# Shared helpers for new-signing-key.ps1, sign-release.ps1 and verify-manifest.ps1 (dot-sourced).
# The update signature is RSA-3072 / SHA-256 (PKCS#1 v1.5) over "ToolsBySanat|<version>|<sha256>|<url>".

$ErrorActionPreference = 'Stop'
$script:DefaultKeyDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'ToolsBySanat-signing'
$script:PublicKeyFile = Join-Path $PSScriptRoot 'keys\update-public.xml'

function ConvertTo-Hex([byte[]]$bytes) { -join ($bytes | ForEach-Object { $_.ToString('x2') }) }
function Get-ManifestMessage($m) { "ToolsBySanat|$($m.version)|$($m.sha256)|$($m.url)" }

function Read-Secret([string]$prompt) {
    $s = Read-Host -Prompt $prompt -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s))
}

function New-KeyMaterial([string]$pass, [byte[]]$salt, [int]$iter) {
    $kdf = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($pass, $salt, $iter, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $k = $kdf.GetBytes(64); $kdf.Dispose()
    @{ Enc = $k[0..31]; Mac = $k[32..63] }
}

# Private key XML -> JSON text, encrypted with a key derived from the passphrase (AES-256-CBC + HMAC-SHA256).
function Protect-PrivateKey([string]$xml, [string]$pass) {
    $iter = 300000
    $salt = New-Object byte[] 16; $iv = New-Object byte[] 16
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create(); $rng.GetBytes($salt); $rng.GetBytes($iv)
    $km = New-KeyMaterial $pass $salt $iter
    $aes = [System.Security.Cryptography.Aes]::Create(); $aes.Key = [byte[]]$km.Enc; $aes.IV = $iv
    $plain = [Text.Encoding]::UTF8.GetBytes($xml)
    $cipher = $aes.CreateEncryptor().TransformFinalBlock($plain, 0, $plain.Length)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256 (, [byte[]]$km.Mac)
    $mac = $hmac.ComputeHash([byte[]]($iv + $cipher))
    @{ v = 1; iter = $iter; salt = [Convert]::ToBase64String($salt); iv = [Convert]::ToBase64String($iv); data = [Convert]::ToBase64String($cipher); mac = [Convert]::ToBase64String($mac) } | ConvertTo-Json
}

function Unprotect-PrivateKey([string]$json, [string]$pass) {
    $o = $json | ConvertFrom-Json
    $salt = [Convert]::FromBase64String($o.salt); $iv = [Convert]::FromBase64String($o.iv)
    $cipher = [Convert]::FromBase64String($o.data); $mac = [Convert]::FromBase64String($o.mac)
    $km = New-KeyMaterial $pass $salt ([int]$o.iter)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256 (, [byte[]]$km.Mac)
    $calc = $hmac.ComputeHash([byte[]]($iv + $cipher))
    if ((ConvertTo-Hex $calc) -ne (ConvertTo-Hex $mac)) { throw 'Wrong passphrase (or the key file is damaged).' }
    $aes = [System.Security.Cryptography.Aes]::Create(); $aes.Key = [byte[]]$km.Enc; $aes.IV = $iv
    [Text.Encoding]::UTF8.GetString($aes.CreateDecryptor().TransformFinalBlock($cipher, 0, $cipher.Length))
}

# $true when the manifest's signature is valid for the given public key XML.
function Test-ManifestSignature($m, [string]$publicXml) {
    if (-not $m.signature -or $m.signature -notmatch '^[0-9a-f]+$') { return $false }
    $sig = [byte[]]@(for ($i = 0; $i -lt $m.signature.Length; $i += 2) { [Convert]::ToByte($m.signature.Substring($i, 2), 16) })
    $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
    $rsa.FromXmlString($publicXml)
    $rsa.VerifyData([Text.Encoding]::ASCII.GetBytes((Get-ManifestMessage $m)), 'SHA256', $sig)
}
