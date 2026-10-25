Attribute VB_Name = "modSettings"
Option Explicit

' Private data (email addresses, customer names and addresses) is NOT stored in the macros.
' It lives in two CSV files on each PC, in %APPDATA%\Microsoft\AddIns\ToolsBySanat_data\ (next to the add-in;
' Excel could not read new folders created directly under %APPDATA% on the build PC, this one works):
'
'   settings.csv       key,value     e.g. vendor_email.KEYA , asn_email.R801-KHURDA-DC , appointment_cc , requestor_email
'   party_master.csv   gstin,name,state,address1,address2,city,pin    one row per customer
'
' Edit them in Excel (save as CSV). Changes are picked up the next time a macro runs.
' The "Settings" button on the ribbon creates empty templates and opens the folder.

Private mSettings As Object      ' key -> value (case-insensitive)
Private mSettingsStamp As Date
Private mParties As Object       ' GSTIN (upper case) -> Array(name, state, address1, address2, city, pin)
Private mPartiesStamp As Date
Private mWarned As Object        ' files already reported missing in this Excel session

Private Const SETTINGS_FILE As String = "settings.csv"
Private Const PARTY_FILE As String = "party_master.csv"

Public Function SettingsDir() As String
    SettingsDir = Environ("APPDATA") & "\Microsoft\AddIns\ToolsBySanat_data\"
End Function

' Value of one setting, or "" if it is not defined.
Public Function SettingValue(ByVal key As String) As String
    LoadSettings
    If mSettings.Exists(key) Then SettingValue = mSettings(key)
End Function

' One field ("name", "state", "address1", "address2", "city", "pin") of the customer with this GSTIN,
' or "" if the GSTIN is not in party_master.csv.
Public Function PartyField(ByVal gstin As Variant, ByVal field As String) As String
    Dim k As String, idx As Long
    LoadParties
    k = UCase$(Trim$(CStr(gstin)))
    If Len(k) = 0 Then Exit Function
    If Not mParties.Exists(k) Then Exit Function
    Select Case LCase$(field)
        Case "name": idx = 0
        Case "state": idx = 1
        Case "address1": idx = 2
        Case "address2": idx = 3
        Case "city": idx = 4
        Case "pin": idx = 5
        Case Else: Exit Function
    End Select
    PartyField = mParties(k)(idx)
End Function

' Ribbon "Settings" button: create empty templates if the files are missing, then open the folder.
Public Sub OpenSettingsFolder()
    Dim d As String
    d = SettingsDir()
    On Error GoTo Fail
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    If Len(Dir(d & SETTINGS_FILE)) = 0 Then
        WriteTemplate d & SETTINGS_FILE, "key,value" & vbCrLf & _
            "vendor_email.EXAMPLE,someone@example.com;" & vbCrLf & _
            "asn_email.EXAMPLE-DC,asn@example.com" & vbCrLf & _
            "appointment_cc,you@example.com;" & vbCrLf & _
            "requestor_email,you@example.com"
    End If
    If Len(Dir(d & PARTY_FILE)) = 0 Then
        WriteTemplate d & PARTY_FILE, "gstin,name,state,address1,address2,city,pin" & vbCrLf & _
            "00AAAAA0000A0Z0,EXAMPLE TRADERS,Orissa,Address line 1,Address line 2,City,000000"
    End If
    Shell "explorer.exe """ & Left$(d, Len(d) - 1) & """", vbNormalFocus
    Exit Sub
Fail:
    MsgBox "Could not open the settings folder:" & vbCrLf & d, vbExclamation, "Tools By Sanat"
End Sub

' ---------------------------------------------------------------- loading

Private Sub LoadSettings()
    Dim path As String, stamp As Date, data As Variant, i As Long, k As String
    path = SettingsDir() & SETTINGS_FILE
    stamp = FileStamp(path)
    If Not mSettings Is Nothing And stamp = mSettingsStamp Then Exit Sub

    Set mSettings = CreateObject("Scripting.Dictionary")
    mSettings.CompareMode = 1          ' text compare: keys are case-insensitive
    mSettingsStamp = stamp
    data = ReadCsv(path)
    If IsEmpty(data) Then WarnMissing SETTINGS_FILE: Exit Sub
    If UBound(data, 2) < 2 Then Exit Sub
    For i = 2 To UBound(data, 1)
        k = Trim$(data(i, 1))
        If Len(k) > 0 Then mSettings(k) = Trim$(data(i, 2))
    Next i
End Sub

Private Sub LoadParties()
    Dim path As String, stamp As Date, data As Variant, i As Long, j As Long, k As String
    Dim col As Object, names As Variant, f(0 To 5) As String
    path = SettingsDir() & PARTY_FILE
    stamp = FileStamp(path)
    If Not mParties Is Nothing And stamp = mPartiesStamp Then Exit Sub

    Set mParties = CreateObject("Scripting.Dictionary")
    mPartiesStamp = stamp
    data = ReadCsv(path)
    If IsEmpty(data) Then WarnMissing PARTY_FILE: Exit Sub

    ' columns are found by header name, so their order in the file does not matter
    Set col = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(data, 2)
        col(LCase$(Trim$(data(1, j)))) = j
    Next j
    names = Array("name", "state", "address1", "address2", "city", "pin")
    If Not col.Exists("gstin") Then Exit Sub
    For i = 2 To UBound(data, 1)
        k = UCase$(Trim$(data(i, col("gstin"))))
        If Len(k) > 0 Then
            For j = 0 To 5
                f(j) = ""
                If col.Exists(names(j)) Then f(j) = data(i, col(names(j)))
            Next j
            mParties(k) = Array(f(0), f(1), f(2), f(3), f(4), f(5))
        End If
    Next i
End Sub

Private Function FileStamp(ByVal path As String) As Date
    On Error Resume Next
    If Len(Dir(path)) > 0 Then FileStamp = FileDateTime(path)
End Function

Private Sub WarnMissing(ByVal fileName As String)
    If mWarned Is Nothing Then Set mWarned = CreateObject("Scripting.Dictionary")
    If mWarned.Exists(fileName) Then Exit Sub
    mWarned(fileName) = True
    MsgBox "The settings file '" & fileName & "' was not found in:" & vbCrLf & SettingsDir() & vbCrLf & vbCrLf & _
           "Email addresses / customer details will be left blank." & vbCrLf & _
           "Click Settings on the Tools By Sanat tab to create the file, then fill it in.", vbExclamation, "Tools By Sanat"
End Sub

Private Sub WriteTemplate(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub

' Reads a UTF-8 / ANSI CSV (quotes, embedded commas, quotes, line breaks) into a 1-based 2D string array.
' Returns Empty if the file is missing or empty.
Private Function ReadCsv(ByVal path As String) As Variant
    Dim stm As Object, txt As String
    If Len(Dir(path)) = 0 Then Exit Function
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    txt = stm.ReadText
    stm.Close
    If Len(txt) > 0 Then If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)
    If Len(txt) = 0 Then Exit Function

    Dim recs As Collection, fields As Collection
    Dim i As Long, n As Long, ch As String, fld As String, inQ As Boolean, nCols As Long, c As Long
    Set recs = New Collection: Set fields = New Collection
    n = Len(txt)
    For i = 1 To n
        ch = Mid$(txt, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(txt, i + 1, 1) = """" Then
                    fld = fld & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                fld = fld & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = "," Then
            fields.Add fld: fld = ""
        ElseIf ch = vbLf Then
            fields.Add fld: fld = ""
            recs.Add fields: Set fields = New Collection
        ElseIf ch <> vbCr Then
            fld = fld & ch
        End If
    Next i
    If Len(fld) > 0 Or fields.Count > 0 Then fields.Add fld: recs.Add fields
    If recs.Count = 0 Then Exit Function

    For c = 1 To recs.Count
        If recs(c).Count > nCols Then nCols = recs(c).Count
    Next c
    Dim out() As String, r As Long, rec As Collection
    ReDim out(1 To recs.Count, 1 To nCols)
    For r = 1 To recs.Count
        Set rec = recs(r)
        For c = 1 To rec.Count
            out(r, c) = rec(c)
        Next c
    Next r
    ReadCsv = out
End Function
