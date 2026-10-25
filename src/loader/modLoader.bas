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

Private Const VERSION_URL As String = "__VERSION_URL__"
Private Const STRICT_URLS As Boolean = True
Private Const MAIN_NAME As String = "ToolsBySanat.xlam"
Private Const MAX_BYTES As Long = 20971520   ' 20 MB

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
        MsgBox "Could not check for updates. Please check your internet connection and try again later.", vbExclamation, "Tools By Sanat"
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
    UpdatesConfigured = (Len(VERSION_URL) > 0 And InStr(VERSION_URL, "__") = 0)
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
    hashFile = DataDir() & "pending\sha256.txt"
    p = AddInsDir() & MAIN_NAME
    prev = DataDir() & "prev\" & MAIN_NAME
    If Len(Dir(pend)) = 0 Then Exit Sub

    On Error GoTo Fail
    ' re-verify: the file may have been touched since it was staged
    If LCase$(Trim$(Replace(Replace(ReadText(hashFile), vbCr, ""), vbLf, ""))) <> FileSha256(pend) Then
        Kill pend
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
    Dim body As Variant, json As String, remoteVer As String, url As String, sha As String
    Dim localVer As String, pend As String, dl As String

    CheckAndStage = "error"
    If Not UpdatesConfigured() Then CheckAndStage = "off": Exit Function
    On Error GoTo Fail

    If Not HttpGet(VERSION_URL & "?t=" & Format$(Now, "yyyymmddhhnnss"), 10, body) Then Exit Function
    json = StrConv(body, vbUnicode)
    remoteVer = JsonValue(json, "version")
    url = JsonValue(json, "url")
    sha = LCase$(JsonValue(json, "sha256"))
    If Len(remoteVer) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Or sha Like "*[!0-9a-f]*" Then Exit Function

    localVer = InstalledVersion()
    If Len(localVer) = 0 Then Exit Function
    If Not IsNewer(remoteVer, localVer) Then CheckAndStage = "none": Exit Function

    pend = DataDir() & "pending\" & MAIN_NAME
    If Len(Dir(pend)) > 0 Then
        If FileSha256(pend) = sha Then CheckAndStage = "staged:" & remoteVer: Exit Function
    End If

    EnsureDirs "pending"
    If Not HttpGet(url, 120, body) Then Exit Function
    If UBound(body) < 1 Or UBound(body) + 1 > MAX_BYTES Then Exit Function
    If body(0) <> 80 Or body(1) <> 75 Then Exit Function          ' must be a zip ("PK"): .xlam is one

    dl = DataDir() & "pending\download.tmp"
    If Len(Dir(dl)) > 0 Then Kill dl
    If Not SaveBytes(dl, body) Then Exit Function
    If FileSha256(dl) <> sha Then
        Kill dl
        Exit Function
    End If

    If Len(Dir(pend)) > 0 Then Kill pend
    Name dl As pend
    WriteText DataDir() & "pending\sha256.txt", sha
    CheckAndStage = "staged:" & remoteVer
    Exit Function
Fail:
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

' MSXML (WinINet) is used so the system proxy settings apply. Redirects (GitHub release downloads) are followed.
Private Function HttpGet(ByVal url As String, ByVal timeoutSec As Long, ByRef body As Variant) As Boolean
    Dim req As Object, t0 As Single
    If Not IsAllowedUrl(url) Then Exit Function
    On Error GoTo Fail
    Set req = CreateObject("MSXML2.XMLHTTP.6.0")
    req.Open "GET", url, True
    req.setRequestHeader "Cache-Control", "no-cache"
    req.send
    t0 = Timer
    Do While req.readyState <> 4
        DoEvents
        If Timer < t0 Or Timer - t0 > timeoutSec Then
            req.abort
            Exit Function
        End If
    Loop
    If req.Status <> 200 Then Exit Function
    body = req.responseBody
    HttpGet = True
    Exit Function
Fail:
    HttpGet = False
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
    Dim sh As Object, tmp As String, lines() As String, i As Long, s As String
    On Error Resume Next
    tmp = Environ("TEMP") & "\tbs_hash_" & Format$(Now, "hhnnss") & ".txt"
    Set sh = CreateObject("WScript.Shell")
    sh.Run "cmd /c certutil -hashfile """ & path & """ SHA256 > """ & tmp & """", 0, True
    lines = Split(ReadText(tmp), vbCrLf)
    Kill tmp
    For i = 0 To UBound(lines)
        s = LCase$(Replace(Trim$(lines(i)), " ", ""))
        If Len(s) = 64 And Not (s Like "*[!0-9a-f]*") Then FileSha256 = s: Exit Function
    Next i
End Function
