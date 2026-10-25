Attribute VB_Name = "modLoader"
Option Explicit

' Loader for the "Tools By Sanat" add-in. It is installed once and rarely changes.
'
' Folder layout (all under %APPDATA%\Microsoft\AddIns, a folder Excel trusts by default):
'   ToolsBySanat.xlam                     the main add-in, opened by this loader
'   ToolsBySanat_data\pending\            a downloaded + verified update waiting for the next start
'   ToolsBySanat_data\prev\               the version that was replaced (rollback)
'
' At Excel start:  apply a pending update (swap BEFORE the main add-in is opened, because Excel
'                  locks loaded add-ins) -> open the main add-in -> a few seconds later check GitHub
'                  for a newer version, download it, verify the SHA-256 and stage it for next start.
' Every failure is silent and keeps the current version. HTTPS and github hosts only.
'
' Trust model: version.json carries an RSA signature (made offline by the owner) over
'   "ToolsBySanat|<version>|<sha256>|<url>". The loader only accepts an update when that signature verifies
'   against PUBLIC_KEY (built in), the url starts with DOWNLOAD_PREFIX (the official repo's releases) and the
'   downloaded file matches the signed SHA-256. The same checks run again right before the file is swapped in.
'   Without a public key the updater stays off (fails closed).

Private Const VERSION_URL As String = "__VERSION_URL__"
Private Const DOWNLOAD_PREFIX As String = "__DOWNLOAD_PREFIX__"
Private Const PUBLIC_KEY As String = "__PUBLIC_KEY__"
Private Const STRICT_URLS As Boolean = True
Private Const MAIN_NAME As String = "ToolsBySanat.xlam"
Private Const MAX_BYTES As Long = 20971520   ' 20 MB

Private mDetail As String                    ' why the last check failed, shown by "Check for updates"
Private mSoftFail As Boolean                 ' last request got no HTTP answer at all (try the next component)

' ---------------------------------------------------------------- entry points

Public Sub Startup()
    On Error Resume Next
    ApplyPending
    OpenMain
    If UpdatesConfigured() Then
        Application.OnTime Now + TimeSerial(0, 0, 5), "'" & ThisWorkbook.Name & "'!modLoader.BackgroundCheck"
    End If
End Sub

' Silent check, scheduled from Startup.
Public Sub BackgroundCheck()
    On Error Resume Next
    CheckAndStage
End Sub

' "Check for updates" button on the ribbon.
Public Sub CheckNow()
    Dim r As String
    On Error Resume Next
    r = CheckAndStage()
    If r = "off" Then
        MsgBox "Updates are not set up in this installation.", vbInformation, "Tools By Sanat"
    ElseIf r = "none" Then
        MsgBox "You have the latest version (" & InstalledVersion() & ").", vbInformation, "Tools By Sanat"
    ElseIf Left$(r, 7) = "staged:" Then
        MsgBox "Version " & Mid$(r, 8) & " was downloaded." & vbCrLf & _
               "Close all Excel windows and open Excel again to finish the update.", vbInformation, "Tools By Sanat"
    Else
        MsgBox "Could not check for updates." & vbCrLf & vbCrLf & mDetail & vbCrLf & vbCrLf & _
               "Please check your internet connection and try again later.", vbExclamation, "Tools By Sanat"
    End If
End Sub

' ---------------------------------------------------------------- paths

Private Function AddInsDir() As String
    AddInsDir = Environ("APPDATA") & "\Microsoft\AddIns\"
End Function

Private Function DataDir() As String
    DataDir = AddInsDir() & "ToolsBySanat_data\"
End Function

Private Function UpdatesConfigured() As Boolean
    UpdatesConfigured = (Len(VERSION_URL) > 0 And Left$(VERSION_URL, 2) <> "__" And _
                         Left$(DOWNLOAD_PREFIX, 2) <> "__" And Left$(PUBLIC_KEY, 2) <> "__")
End Function

' ---------------------------------------------------------------- opening / swapping

Private Function IsOpen(ByVal nm As String) As Boolean
    Dim wb As Workbook
    For Each wb In Workbooks
        If LCase$(wb.Name) = LCase$(nm) Then IsOpen = True: Exit Function
    Next wb
End Function

Private Sub OpenMain()
    Dim p As String
    On Error Resume Next
    p = AddInsDir() & MAIN_NAME
    If IsOpen(MAIN_NAME) Then Exit Sub
    If Len(Dir(p)) = 0 Then Exit Sub
    Err.Clear
    Workbooks.Open p
    If Err.Number <> 0 And Not IsOpen(MAIN_NAME) Then
        Err.Clear
        If RestorePrevious() Then Workbooks.Open p   ' the new file would not open: roll back
    End If
End Sub

Private Function RestorePrevious() As Boolean
    Dim prev As String, p As String
    On Error GoTo Fail
    prev = DataDir() & "prev\" & MAIN_NAME
    p = AddInsDir() & MAIN_NAME
    If Len(Dir(prev)) = 0 Then Exit Function
    If Len(Dir(p)) > 0 Then Kill p
    FileCopy prev, p
    RestorePrevious = True
    Exit Function
Fail:
    RestorePrevious = False
End Function

Private Sub ApplyPending()
    Dim pend As String, hashFile As String, p As String, prev As String
    pend = DataDir() & "pending\" & MAIN_NAME
    hashFile = DataDir() & "pending\manifest.txt"
    p = AddInsDir() & MAIN_NAME
    prev = DataDir() & "prev\" & MAIN_NAME
    If Len(Dir(pend)) = 0 Then Exit Sub

    On Error GoTo Fail
    ' re-verify: the file and its signed manifest may have been touched since they were staged
    If Not ManifestValid(hashFile, pend) Then
        Kill pend
        If Len(Dir(hashFile)) > 0 Then Kill hashFile
        Exit Sub
    End If

    EnsureDirs "prev"
    If Len(Dir(p)) > 0 Then
        If Len(Dir(prev)) > 0 Then Kill prev
        FileCopy p, prev
        Kill p              ' fails while another Excel window has it open: the update waits for the next start
    End If
    Name pend As p
    If Len(Dir(hashFile)) > 0 Then Kill hashFile
    Exit Sub

Fail:
    ' never leave the user without a main add-in
    On Error Resume Next
    If Len(Dir(p)) = 0 Then
        If Len(Dir(prev)) > 0 Then FileCopy prev, p
    End If
End Sub

' ---------------------------------------------------------------- update check

Private Function InstalledVersion() As String
    On Error Resume Next
    InstalledVersion = Application.Run("'" & MAIN_NAME & "'!modRibbon.AddinVersion")
End Function

' Returns "off" (not configured), "none" (up to date), "staged:<version>" (downloaded, applies at
' next start) or "error".
Private Function CheckAndStage() As String
    Dim body As Variant, json As String, remoteVer As String, url As String, sha As String, sig As String
    Dim localVer As String, pend As String, dl As String

    CheckAndStage = "error"
    mDetail = ""
    If Not UpdatesConfigured() Then CheckAndStage = "off": Exit Function
    On Error GoTo Fail

    If Not HttpGet(VERSION_URL & "?t=" & Format$(Now, "yyyymmddhhnnss"), 10, body) Then mDetail = "Reading the version information failed: " & mDetail: Exit Function
    json = StrConv(body, vbUnicode)
    remoteVer = JsonValue(json, "version")
    url = JsonValue(json, "url")
    sha = LCase$(JsonValue(json, "sha256"))
    sig = JsonValue(json, "signature")
    If Len(remoteVer) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Or sha Like "*[!0-9a-f]*" Then mDetail = "The version information is not in the expected format.": Exit Function
    If Not VerifySignature(remoteVer, sha, url, sig) Then mDetail = mDetail & "The update information is not signed by the owner, so it was ignored.": Exit Function
    If LCase$(Left$(url, Len(DOWNLOAD_PREFIX))) <> LCase$(DOWNLOAD_PREFIX) Then mDetail = "The update address is not the official one, so it was ignored.": Exit Function

    localVer = InstalledVersion()
    If Len(localVer) = 0 Then mDetail = "The installed add-in did not report its version.": Exit Function
    If Not IsNewer(remoteVer, localVer) Then CheckAndStage = "none": Exit Function

    pend = DataDir() & "pending\" & MAIN_NAME
    If Len(Dir(pend)) > 0 Then
        If FileSha256(pend) = sha Then
            WriteText DataDir() & "pending\manifest.txt", remoteVer & vbCrLf & sha & vbCrLf & url & vbCrLf & sig
            CheckAndStage = "staged:" & remoteVer
            Exit Function
        End If
    End If

    EnsureDirs "pending"
    If Not HttpGet(url, 120, body) Then mDetail = "Downloading the update failed: " & mDetail: Exit Function
    If UBound(body) < 1 Or UBound(body) + 1 > MAX_BYTES Then mDetail = "The downloaded file has an unexpected size.": Exit Function
    If body(0) <> 80 Or body(1) <> 75 Then mDetail = "The downloaded file is not an add-in file.": Exit Function   ' .xlam is a zip ("PK")

    dl = DataDir() & "pending\download.tmp"
    If Len(Dir(dl)) > 0 Then Kill dl
    If Not SaveBytes(dl, body) Then mDetail = "Could not save the download to " & dl: Exit Function
    If FileSha256(dl) <> sha Then
        mDetail = "The downloaded file failed the integrity check."
        Kill dl
        Exit Function
    End If

    If Len(Dir(pend)) > 0 Then Kill pend
    Name dl As pend
    WriteText DataDir() & "pending\manifest.txt", remoteVer & vbCrLf & sha & vbCrLf & url & vbCrLf & sig
    CheckAndStage = "staged:" & remoteVer
    Exit Function
Fail:
    mDetail = mDetail & "Unexpected error " & Err.Number & ": " & Err.Description
    CheckAndStage = "error"
End Function

Private Function IsNewer(ByVal remoteVer As String, ByVal installedVer As String) As Boolean
    Dim a() As String, b() As String, i As Long, x As Long, y As Long
    a = Split(remoteVer, ".")
    b = Split(installedVer, ".")
    For i = 0 To 2
        x = 0: y = 0
        If i <= UBound(a) Then x = Val(a(i))
        If i <= UBound(b) Then y = Val(b(i))
        If x > y Then IsNewer = True: Exit Function
        If x < y Then Exit Function
    Next i
End Function

' True if the signature (hex) is a valid RSA/SHA-256 signature, made with the owner's private key, over
' "ToolsBySanat|<version>|<sha256>|<url>". Uses the .NET RSA component that ships with Windows.
Private Function VerifySignature(ByVal ver As String, ByVal sha As String, ByVal url As String, ByVal sigHex As String) As Boolean
    Dim rsa As Object, b() As Byte, d() As Byte, i As Long, msg As String
    On Error GoTo Fail
    sigHex = LCase$(Replace(Replace(Trim$(sigHex), vbCr, ""), vbLf, ""))
    If Len(sigHex) < 128 Or Len(sigHex) Mod 2 <> 0 Or sigHex Like "*[!0-9a-f]*" Then mDetail = mDetail & "No valid signature. ": Exit Function
    ReDim b(0 To Len(sigHex) \ 2 - 1)
    For i = 0 To UBound(b)
        b(i) = CByte("&H" & Mid$(sigHex, i * 2 + 1, 2))
    Next i
    msg = "ToolsBySanat|" & ver & "|" & sha & "|" & url
    d = StrConv(msg, vbFromUnicode)                  ' typed byte arrays: the .NET call rejects an inline Variant
    Set rsa = CreateObject("System.Security.Cryptography.RSACryptoServiceProvider")
    rsa.FromXmlString PUBLIC_KEY
    VerifySignature = CBool(rsa.VerifyData(d, "SHA256", b))
    Exit Function
Fail:
    mDetail = mDetail & "Signature check error " & Err.Number & ": " & Err.Description & ". "
    VerifySignature = False
End Function

' The staged update's manifest (version, sha256, url, signature - one per line) must carry a valid signature,
' an official address, and describe exactly the staged file.
Private Function ManifestValid(ByVal manifestFile As String, ByVal stagedFile As String) As Boolean
    Dim lines() As String
    On Error GoTo Fail
    If Len(Dir(manifestFile)) = 0 Then Exit Function
    lines = Split(Replace(ReadText(manifestFile), vbCr, ""), vbLf)
    If UBound(lines) < 3 Then Exit Function
    If Not VerifySignature(Trim$(lines(0)), Trim$(lines(1)), Trim$(lines(2)), lines(3)) Then Exit Function
    If LCase$(Left$(Trim$(lines(2)), Len(DOWNLOAD_PREFIX))) <> LCase$(DOWNLOAD_PREFIX) Then Exit Function
    ManifestValid = (LCase$(Trim$(lines(1))) = FileSha256(stagedFile))
    Exit Function
Fail:
    ManifestValid = False
End Function

' Minimal reader for {"key": "value"} pairs, enough for version.json.
Private Function JsonValue(ByVal json As String, ByVal key As String) As String
    Dim p As Long, q As Long
    p = InStr(json, """" & key & """")
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, json, ":")
    If p = 0 Then Exit Function
    p = InStr(p, json, """")
    If p = 0 Then Exit Function
    q = InStr(p + 1, json, """")
    If q = 0 Then Exit Function
    JsonValue = Mid$(json, p + 1, q - p - 1)
End Function

' ---------------------------------------------------------------- network

Private Function IsAllowedUrl(ByVal url As String) As Boolean
    Dim host As String, p As Long
    If Not STRICT_URLS Then IsAllowedUrl = True: Exit Function
    If LCase$(Left$(url, 8)) <> "https://" Then Exit Function
    host = Mid$(url, 9)
    p = InStr(host, "/")
    If p > 0 Then host = Left$(host, p - 1)
    host = LCase$(host)
    IsAllowedUrl = (host = "github.com" Or host = "raw.githubusercontent.com")
End Function

' Tries three Windows web components in turn and keeps the first good answer:
'   1. MSXML2.XMLHTTP       uses the system proxy settings, but cannot follow GitHub's cross-site download redirect
'   2. MSXML2.ServerXMLHTTP follows redirects
'   3. WinHttp.WinHttpRequest follows redirects
' Every failure is noted in mDetail so "Check for updates" can say why.
Private Function HttpGet(ByVal url As String, ByVal timeoutSec As Long, ByRef body As Variant) As Boolean
    Dim k As Long
    If Not IsAllowedUrl(url) Then mDetail = mDetail & "address not allowed; ": Exit Function
    For k = 1 To 3
        mSoftFail = False
        If TryGet(k, url, timeoutSec, body) Then HttpGet = True: Exit Function
        If Not mSoftFail Then Exit For      ' network down, timeout or HTTP error: another component will not help
    Next k
End Function

Private Function TryGet(ByVal kind As Long, ByVal url As String, ByVal timeoutSec As Long, ByRef body As Variant) As Boolean
    Dim req As Object, t0 As Single, comp As String, st As Long, px As String
    Select Case kind
        Case 1: comp = "XMLHTTP"
        Case 2: comp = "ServerXMLHTTP"
        Case Else: comp = "WinHttp"
    End Select
    On Error GoTo Fail
    Select Case kind
        Case 1
            Set req = CreateObject("MSXML2.XMLHTTP.6.0")
            req.Open "GET", url, True
            req.setRequestHeader "Cache-Control", "no-cache"
            req.send
            t0 = Timer
            Do While req.readyState <> 4
                DoEvents
                If Timer < t0 Or Timer - t0 > timeoutSec Then
                    req.abort
                    mDetail = mDetail & comp & ": timed out; "
                    Exit Function
                End If
            Loop
        Case 2
            Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            req.setTimeouts 4000, 4000, 10000, timeoutSec * 1000
            px = IeProxy()
            If Len(px) > 0 Then req.setProxy 2, px
            req.Open "GET", url, False
            req.setRequestHeader "Cache-Control", "no-cache"
            req.send
        Case Else
            Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
            req.SetTimeouts 4000, 4000, 10000, timeoutSec * 1000
            px = IeProxy()
            If Len(px) > 0 Then req.SetProxy 2, px
            req.Open "GET", url, False
            req.setRequestHeader "Cache-Control", "no-cache"
            req.send
    End Select
    st = req.Status
    If st <> 200 Then
        mDetail = mDetail & comp & ": HTTP " & st & "; "
        mSoftFail = (st = 0)                ' "no answer" is what XMLHTTP gives for a blocked redirect
        Exit Function
    End If
    body = req.responseBody
    TryGet = True
    Exit Function
Fail:
    mDetail = mDetail & comp & ": " & Err.Description & "; "
End Function

' The proxy that Windows / Internet Options uses, as "host:port", or "" for a direct connection.
' (MSXML2.XMLHTTP reads it by itself; the two fallbacks need it passed in. PAC scripts are not supported.)
Private Function IeProxy() As String
    Dim sh As Object, key As String, s As String, parts() As String, i As Long
    On Error GoTo Done
    Set sh = CreateObject("WScript.Shell")
    key = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
    If CLng(sh.RegRead(key & "ProxyEnable")) <> 1 Then Exit Function
    s = CStr(sh.RegRead(key & "ProxyServer"))
    If InStr(s, "=") = 0 Then
        IeProxy = s
    Else
        parts = Split(s, ";")                       ' "http=a:80;https=b:443"
        For i = 0 To UBound(parts)
            If LCase$(Left$(parts(i), 6)) = "https=" Then IeProxy = Mid$(parts(i), 7): Exit Function
        Next i
    End If
Done:
End Function

' ---------------------------------------------------------------- files

Private Sub EnsureDirs(ByVal subFolder As String)
    Dim d As String
    d = Left$(DataDir(), Len(DataDir()) - 1)
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    d = DataDir() & subFolder
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
End Sub

Private Function SaveBytes(ByVal path As String, body As Variant) As Boolean
    Dim st As Object
    On Error GoTo Fail
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write body
    st.SaveToFile path, 2
    st.Close
    SaveBytes = True
    Exit Function
Fail:
    SaveBytes = False
End Function

Private Function ReadText(ByVal path As String) As String
    Dim f As Integer
    If Len(Dir(path)) = 0 Then Exit Function
    f = FreeFile
    Open path For Input As #f
    If LOF(f) > 0 Then ReadText = Input$(LOF(f), f)
    Close #f
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub

' Lower-case SHA-256 of a file via the Windows certutil tool ("" if it cannot be computed).
Private Function FileSha256(ByVal path As String) As String
    Dim sh As Object, tmp As String, lines() As String, i As Long, s As String, q As String, sys As String
    On Error Resume Next
    q = Chr$(34)
    sys = Environ("SystemRoot") & "\System32\"          ' full paths: do not search the PATH for the tools
    tmp = Environ("TEMP") & "\tbs_hash_" & Format$(Now, "hhnnss") & "_" & CStr(CLng(Timer * 1000)) & ".txt"
    Set sh = CreateObject("WScript.Shell")
    sh.Run q & sys & "cmd.exe" & q & " /c " & q & q & sys & "certutil.exe" & q & " -hashfile " & q & path & q & " SHA256 > " & q & tmp & q & q, 0, True
    lines = Split(ReadText(tmp), vbCrLf)
    Kill tmp
    For i = 0 To UBound(lines)
        s = LCase$(Replace(Trim$(lines(i)), " ", ""))
        If Len(s) = 64 And Not (s Like "*[!0-9a-f]*") Then FileSha256 = s: Exit Function
    Next i
End Function
