Attribute VB_Name = "Flipkart_PO_formatter"
Option Explicit
'------------------------------------------------------------------
' Flipkart PO formatter
' Reads the raw PO export from the ACTIVE sheet of the active workbook and builds a new workbook in
' the Output.xlsx layout: filtered, 17 columns, sorted by PO ID.
'
' Install : Excel > Alt+F11 > File > Import File... > FormatPO.bas
' Run     : open the Input file, leave it as the active window, Alt+F8 > FormatPO > Run
' Save as : .xlsm if you want to keep the macro inside the workbook.
'------------------------------------------------------------------

' ---- Settings (edit here) ----
Private Const KEEP_STATUS As String = "APPROVED"      ' comma separated, "" = all
Private Const KEEP_FROM_OU As String = "VS54377"      ' comma separated, "" = all
Private Const ONLY_PENDING As Boolean = False         ' True = keep only Pending Quantity > 0
Private Const BRAND_RENAME As String = "Choki Choki=Malkist,KOPIKO=Malkist"   ' Old=New,Old=New

' Output header | input header | type (s=text, d=date, n=number)
Private Const LAYOUT As String = _
    "PO ID|PO ID|s;Order Date|Order Date|d;Expiry Date|Expiry Date|d;PO Status|PO Status|s;" & _
    "PO Item Status|PO Item Status|s;Brand|Brand|s;PO Item ID|PO Item ID|s;Product ID|Product ID|s;Title|Title|s;" & _
    "mrp|supplier_mrp_amount|n;Ordered Quantity|Ordered Quantity|n;Approved Quantity|Approved Quantity|n;" & _
    "Cancelled Quantity|Cancelled Quantity|n;Pending Quantity|Pending Quantity|n;" & _
    "Total Goods Amount|Total Goods Amount|n;Total Tax Amount|Total Tax Amount|n;" & _
    "Total Amount|Total Amount|n;supplier_app_amount|supplier_app_amount|n"

Public Sub Flipkart_PO_formatter()
    Dim wsIn As Worksheet
    Dim oldCalc As XlCalculation, oldScr As Boolean, oldAlerts As Boolean

    If ActiveWorkbook Is Nothing Then MsgBox "Open the Input file first.", vbExclamation, "Format PO": Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then MsgBox "Select the Input worksheet first.", vbExclamation, "Format PO": Exit Sub
    Set wsIn = ActiveSheet

    oldScr = Application.ScreenUpdating: oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    ' ---- read input into memory ----
    Dim lastR As Long, lastC As Long, data As Variant
    lastR = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).row
    lastC = wsIn.Cells(1, wsIn.Columns.Count).End(xlToLeft).Column
    If lastR < 2 Then Err.Raise vbObjectError + 1, , "Input file has no data rows."
    data = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(lastR, lastC)).Value2

    ' ---- map columns by header name ----
    Dim spec() As String, nOut As Long, i As Long, j As Long, p() As String
    spec = Split(LAYOUT, ";"): nOut = UBound(spec) + 1
    Dim hdrOut() As String, typ() As String, src() As Long
    ReDim hdrOut(1 To nOut): ReDim typ(1 To nOut): ReDim src(1 To nOut)
    Dim missing As String
    For i = 1 To nOut
        p = Split(spec(i - 1), "|")
        hdrOut(i) = p(0): typ(i) = p(2): src(i) = ColOf(data, lastC, p(1))
        If src(i) = 0 Then missing = missing & vbLf & " - " & p(1)
    Next i
    If Len(missing) > 0 Then Err.Raise vbObjectError + 2, , "Input is missing column(s):" & missing
    Dim cStatus As Long, cOu As Long, cPend As Long
    cStatus = ColOf(data, lastC, "PO Status")
    cOu = ColOf(data, lastC, "From OU ID")
    cPend = ColOf(data, lastC, "Pending Quantity")
    If Len(KEEP_FROM_OU) > 0 And cOu = 0 Then Err.Raise vbObjectError + 3, , "Input has no 'From OU ID' column."

    Dim stList As String, ouList As String
    stList = "," & UCase$(Replace(KEEP_STATUS, " ", "")) & ","
    ouList = "," & UCase$(Replace(KEEP_FROM_OU, " ", "")) & ","

    ' ---- filter + convert ----
    Dim outArr() As Variant, n As Long, r As Long, keep As Boolean, v As Variant
    ReDim outArr(1 To lastR - 1, 1 To nOut + 1)
    For r = 2 To lastR
        keep = True
        If Len(KEEP_STATUS) > 0 Then keep = (InStr(stList, "," & UCase$(Trim$(CStr(data(r, cStatus)))) & ",") > 0)
        If keep And Len(KEEP_FROM_OU) > 0 Then keep = (InStr(ouList, "," & UCase$(Trim$(CStr(data(r, cOu)))) & ",") > 0)
        If keep And ONLY_PENDING Then keep = (ToNum(data(r, cPend)) > 0)
        If keep Then
            n = n + 1
            For j = 1 To nOut
                v = data(r, src(j))
                Select Case typ(j)
                    Case "d": outArr(n, j) = ToDay(v)
                    Case "n": outArr(n, j) = ToNum(v)
                    Case Else
                        If hdrOut(j) = "Brand" Then v = RenameBrand(CStr(v))
                        outArr(n, j) = CStr(v)
                End Select
            Next j
            outArr(n, nOut + 1) = n                       ' original sequence for stable sort
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 4, , "No rows matched the filters (status / From OU ID)."

    ' ---- write output workbook ----
    Dim wbOut As Workbook, ws As Worksheet
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbOut.Worksheets(1): ws.name = "Output"
    For j = 1 To nOut: ws.Cells(1, j).Value = hdrOut(j): Next j

    For j = 1 To nOut      ' text columns first, so IDs are never auto-converted
        If typ(j) = "s" Then ws.Range(ws.Cells(2, j), ws.Cells(n + 1, j)).NumberFormat = "@"
    Next j
    Dim blk() As Variant
    ReDim blk(1 To n, 1 To nOut + 1)
    For i = 1 To n: For j = 1 To nOut + 1: blk(i, j) = outArr(i, j): Next j: Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, nOut + 1)).Value = blk

    ' sort by PO ID then original order
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)), order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(2, nOut + 1), ws.Cells(n + 1, nOut + 1)), order:=xlAscending
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, nOut + 1))
        .Header = xlYes
        .Apply
    End With
    ws.Columns(nOut + 1).Delete

    ' formatting
    For j = 1 To nOut
        If typ(j) = "d" Then ws.Range(ws.Cells(2, j), ws.Cells(n + 1, j)).NumberFormat = "dd-mm-yy"
    Next j
    ws.Rows(1).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, nOut)).AutoFilter
    ws.Columns.AutoFit
    If ws.Columns(9).ColumnWidth > 70 Then ws.Columns(9).ColumnWidth = 70
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScr
    MsgBox n & " rows written to a new workbook (" & wbOut.name & ")." & vbLf & _
           "Save it as Output.xlsx.", vbInformation, "Format PO"
    Exit Sub

Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScr
    Application.DisplayAlerts = oldAlerts
    MsgBox "Format PO failed:" & vbLf & msg, vbExclamation, "Format PO"
End Sub

' ---- helpers ----
Private Function ColOf(data As Variant, lastC As Long, ByVal name As String) As Long
    Dim c As Long
    For c = 1 To lastC
        If LCase$(Trim$(CStr(data(1, c)))) = LCase$(name) Then ColOf = c: Exit Function
    Next c
End Function

Private Function ToNum(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        v = Trim$(Replace(v, ",", ""))
        If v = "" Or UCase$(v) = "NA" Then Exit Function
        ToNum = Val(v)                      ' Val is locale independent ("." decimal)
    Else
        ToNum = CDbl(v)
    End If
End Function

Private Function ToDay(ByVal v As Variant) As Variant
    ' Returns an integer Excel date serial, or Empty if unreadable
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(v)
        If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
            ToDay = CLng(DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))))
        ElseIf IsDate(s) Then
            ToDay = CLng(Int(CDate(s)))
        End If
    Else
        ToDay = CLng(Int(CDbl(v)))
    End If
End Function

Private Function RenameBrand(ByVal b As String) As String
    Dim pairs() As String, k As Long, kv() As String
    RenameBrand = b
    If Len(BRAND_RENAME) = 0 Then Exit Function
    pairs = Split(BRAND_RENAME, ",")
    For k = 0 To UBound(pairs)
        kv = Split(pairs(k), "=")
        If UBound(kv) = 1 Then
            If LCase$(Trim$(b)) = LCase$(Trim$(kv(0))) Then RenameBrand = Trim$(kv(1)): Exit Function
        End If
    Next k
End Function


