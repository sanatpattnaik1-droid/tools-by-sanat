# Diagnoses why "Check for updates" fails on a PC. Safe to run: it only reads (no changes).
# Paste into PowerShell or run:  powershell -ExecutionPolicy Bypass -File diagnose-update.ps1

$url = 'https://raw.githubusercontent.com/sanatpattnaik1-droid/tools-by-sanat/main/version.json'
$os = Get-CimInstance Win32_OperatingSystem
"Windows        : $($os.Caption) $($os.Version)"
"Date and time  : $(Get-Date)"
"Time zone      : $((Get-TimeZone).Id)"
$ie = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
"IE/WinINet TLS : SecureProtocols=$($ie.SecureProtocols)   (2048 = TLS 1.2 on; 0 or blank = system default)"
"Proxy          : enabled=$($ie.ProxyEnable) server=$($ie.ProxyServer) autoconfig=$($ie.AutoConfigURL)"
"Excel version  : $((Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)')"
"Add-in files   : " + ((Get-ChildItem "$env:APPDATA\Microsoft\AddIns" -Filter 'ToolsBySanat*.xlam' -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Name) ($($_.Length))" }) -join ', ')
""
"--- DNS ---"
try { (Resolve-DnsName raw.githubusercontent.com -ErrorAction Stop | Where-Object { $_.IPAddress } | Select-Object -First 2 -ExpandProperty IPAddress) -join ', ' } catch { "DNS failed: $($_.Exception.Message)" }
""
"--- MSXML (the same component the add-in uses) ---"
try {
    $r = New-Object -ComObject MSXML2.XMLHTTP.6.0
    $r.open('GET', $url, $false); $r.send()
    "HTTP status: $($r.status) $($r.statusText)"
    ($r.responseText -replace '\s+', ' ').Substring(0, [Math]::Min(120, $r.responseText.Length))
} catch { "MSXML failed: $($_.Exception.Message)" }
""
"--- .NET (TLS 1.2 forced) ---"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; "HTTP status: $((Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 15).StatusCode)" } catch { ".NET failed: $($_.Exception.Message)" }
""
"--- browser-style check of the download host ---"
try { "github.com: HTTP $((Invoke-WebRequest 'https://github.com/sanatpattnaik1-droid/tools-by-sanat/releases/latest' -UseBasicParsing -TimeoutSec 15).StatusCode)" } catch { "github.com failed: $($_.Exception.Message)" }
