Attribute VB_Name = "BlinkitPOFormatter"
Option Explicit

' Blinkit PO formatter
' Turns the Blinkit PO export that is open in the active sheet (header in row 1) into a new workbook:
'   Sheet2 - PivotTable: manufacturer_name > po_number > order_date > expiry_date,
'            Sum of remaining_quantity, Sum of total_amount, Grand Total
'   Sheet1 - data: only the chosen po_state (default Unscheduled), mrp moved after upc,
'            po_number / upc as text, dates as real dates (dd-mm-yy), frozen header, filter, auto-fit
' The active workbook is not changed and the new workbook is left open and unsaved.
'
' Works on any workbook; no references needed (late binding only).

Private Const DEFAULT_STATE As String = "Unscheduled"
Private Const DATE_FMT As String = "dd-mm-yy"
Private Const TITLE As String = "Blinkit PO formatter"

Public Sub FormatBlinkitPO()
    Dim data As Variant, stateFilter As String, answer As String, n As Long

    data = ReadActiveSheet()
    If IsEmpty(data) Then
        MsgBox "Open the Blinkit PO export and select its sheet first (header in row 1, data below).", vbExclamation, TITLE
        Exit Sub
    End If
    If IndexOfCol(data, "po_number") = 0 Then
        MsgBox "The active sheet does not look like a Blinkit PO export (no po_number column in row 1).", vbExclamation, TITLE
        Exit Sub
    End If

    answer = InputBox("Keep only rows with this po_state." & vbCrLf & _
                      "Leave empty to keep all rows.", TITLE, DEFAULT_STATE)
    If StrPtr(answer) = 0 Then Exit Sub   ' Cancel
    stateFilter = Trim$(answer)

    n = RunFormat(data, stateFilter)
    If n > 0 Then MsgBox n & " rows written to a new workbook (not saved).", vbInformation, TITLE
End Sub

' Does the work without any dialogs. Returns rows written (0 = nothing matched or failure).
' The new workbook is left open and unsaved.
Public Function RunFormat(ByVal data As Variant, ByVal stateFilter As String) As Long
    Dim outRows As Long
    Dim wb As Workbook, wsData As Worksheet, wsPivot As Worksheet
    Dim prevCalc As XlCalculation

    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set wsData = wb.Worksheets(1)
    wsData.name = "Sheet1"
    Set wsPivot = wb.Worksheets.Add(Before:=wsData)
    wsPivot.name = "Sheet2"

    outRows = WriteData(wsData, data, stateFilter)
    If outRows = 0 Then
        Application.DisplayAlerts = False
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.Calculation = prevCalc
        Application.ScreenUpdating = True
        MsgBox "No rows with po_state = """ & stateFilter & """ were found.", vbExclamation, TITLE
        Exit Function
    End If

    FormatDataSheet wsData, outRows
    BuildPivot wb, wsData, wsPivot, outRows

    wsData.Activate
    wsData.Range("A2").Select
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    RunFormat = outRows
    Exit Function

Fail:
    Dim msg As String
    msg = Err.Description
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Failed: " & msg, vbCritical, TITLE
End Function

' ---------------------------------------------------------------- data

' Column order of the formatted sheet (same as Output.xlsx). Unknown columns are appended.
Private Function ColumnOrder() As Variant
    ColumnOrder = Array("po_number", "facility_name", "manufacturer_name", "entity_vendor_legal_name", _
        "vendor_name", "order_date", "appointment_date", "expiry_date", "po_state", "item_id", "name", _
        "uom_text", "upc", "mrp", "units_ordered", "remaining_quantity", "landing_rate", "cost_price", _
        "margin_percentage", "cess_value", "sgst_value", "igst_value", "cgst_value", "tax_value", "total_amount")
End Function

Private Function IsIn(ByVal name As String, ByVal list As String) As Boolean
    IsIn = InStr(1, "|" & list & "|", "|" & name & "|", vbBinaryCompare) > 0
End Function

' Returns number of data rows written (excluding header) to ws starting at A1.
Private Function WriteData(ws As Worksheet, data As Variant, ByVal stateFilter As String) As Long
    Dim hdr() As String, nSrc As Long, i As Long, j As Long, k As Long
    Dim order As Variant, srcIdx() As Long, nOut As Long, stateIdx As Long
    Dim names() As String

    nSrc = UBound(data, 2)
    ReDim hdr(1 To nSrc)
    For j = 1 To nSrc: hdr(j) = Trim$(CStr(data(1, j))): Next j

    ' output column list = known columns present in file, in fixed order, then extras
    order = ColumnOrder()
    ReDim names(1 To nSrc)
    ReDim srcIdx(1 To nSrc)
    For k = LBound(order) To UBound(order)
        j = IndexOf(hdr, CStr(order(k)))
        If j > 0 Then nOut = nOut + 1: names(nOut) = hdr(j): srcIdx(nOut) = j
    Next k
    For j = 1 To nSrc
        If Not IsIn(hdr(j), Join(SliceStr(names, nOut), "|")) Then
            nOut = nOut + 1: names(nOut) = hdr(j): srcIdx(nOut) = j
        End If
    Next j

    stateIdx = IndexOf(hdr, "po_state")

    ' text formats must be set before values are written so IDs keep every digit
    Dim textCols As String, dateCols As String, textOnly As String
    textCols = "po_number|upc"
    dateCols = "order_date|appointment_date|expiry_date"
    textOnly = "facility_name|manufacturer_name|entity_vendor_legal_name|vendor_name|po_state|name|uom_text"
    For k = 1 To nOut
        If IsIn(names(k), textCols) Then ws.Range(ws.Cells(1, k), ws.Cells(UBound(data, 1), k)).NumberFormat = "@"
        If IsIn(names(k), dateCols) Then ws.Range(ws.Cells(1, k), ws.Cells(UBound(data, 1), k)).NumberFormat = DATE_FMT
    Next k

    Dim outArr() As Variant, r As Long, src As String
    ReDim outArr(1 To UBound(data, 1), 1 To nOut)
    For k = 1 To nOut: outArr(1, k) = SafeText(names(k)): Next k
    r = 1
    For i = 2 To UBound(data, 1)
        If Len(stateFilter) > 0 And stateIdx > 0 Then
            If CStr(data(i, stateIdx)) <> stateFilter Then GoTo NextRow
        End If
        r = r + 1
        For k = 1 To nOut
            src = CStr(data(i, srcIdx(k)))
            If Len(src) = 0 Then
                ' leave blank
            ElseIf IsIn(names(k), textCols) Then
                outArr(r, k) = src                       ' these columns are formatted as text
            ElseIf IsIn(names(k), textOnly) Then
                outArr(r, k) = SafeText(src)             ' text from the file must not become a formula
            ElseIf IsIn(names(k), dateCols) Then
                outArr(r, k) = Guard(ParseDate(src, names(k) = "appointment_date"))
            Else
                outArr(r, k) = Guard(ParseNumber(src))
            End If
        Next k
NextRow:
    Next i

    If r > 1 Then ws.Range("A1").Resize(r, nOut).Value = SliceRows(outArr, r, nOut)
    If r < UBound(data, 1) Then ws.Range(ws.Cells(r + 1, 1), ws.Cells(UBound(data, 1), nOut)).Clear
    WriteData = r - 1
End Function

' A value that could not be parsed as a date / number comes back as the original text: make it safe too.
Private Function Guard(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then Guard = SafeText(CStr(v)) Else Guard = v
End Function

Private Function SliceRows(arr As Variant, ByVal nRows As Long, ByVal nCols As Long) As Variant
    Dim o() As Variant, i As Long, j As Long
    ReDim o(1 To nRows, 1 To nCols)
    For i = 1 To nRows: For j = 1 To nCols: o(i, j) = arr(i, j): Next j: Next i
    SliceRows = o
End Function

Private Function SliceStr(arr() As String, ByVal n As Long) As Variant
    Dim o() As String, i As Long
    If n = 0 Then SliceStr = Array(""): Exit Function
    ReDim o(0 To n - 1)
    For i = 1 To n: o(i - 1) = arr(i): Next i
    SliceStr = o
End Function

Private Function IndexOf(arr() As String, ByVal s As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = s Then IndexOf = i: Exit Function
    Next i
End Function

' "2026-10-14 00:00:00 +0000 UTC" -> Date (time dropped unless keepTime and non-midnight)
Private Function ParseDate(ByVal s As String, ByVal keepTime As Boolean) As Variant
    Dim d As Date, hh As Long, mi As Long, ss As Long
    If Len(s) < 10 Or Mid$(s, 5, 1) <> "-" Then ParseDate = s: Exit Function
    d = DateSerial(CLng(Mid$(s, 1, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
    If keepTime And Len(s) >= 19 Then
        d = d + TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
    End If
    ParseDate = d
End Function

' Locale independent (always '.' as decimal separator)
Private Function ParseNumber(ByVal s As String) As Variant
    If s Like "*[!0-9.eE+-]*" Then ParseNumber = s Else ParseNumber = Val(s)
End Function

Private Sub FormatDataSheet(ws As Worksheet, ByVal nRows As Long)
    Dim lastCol As Long, c As Long, hdr As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hdr = CStr(ws.Cells(1, c).Value)
        If hdr = "appointment_date" Then ws.Range(ws.Cells(2, c), ws.Cells(nRows + 1, c)).NumberFormat = "dd-mm-yy hh:mm"
    Next c
    ws.Range(ws.Cells(1, 1), ws.Cells(nRows + 1, lastCol)).Columns.AutoFit
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' ---------------------------------------------------------------- pivot

Private Sub BuildPivot(wb As Workbook, wsData As Worksheet, wsPivot As Worksheet, ByVal nRows As Long)
    Dim lastCol As Long, pc As PivotCache, pt As PivotTable, f As Variant, pf As PivotField
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, _
                SourceData:=wsData.Range(wsData.Cells(1, 1), wsData.Cells(nRows + 1, lastCol)))
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="PivotTable1")

    pt.ManualUpdate = True
    pt.RowAxisLayout xlTabularRow
    pt.HasAutoFormat = False
    pt.TableStyle2 = "PivotStyleLight16"
    pt.ColumnGrand = True
    pt.RowGrand = False

    Dim n As Long
    For Each f In Array("manufacturer_name", "po_number", "order_date", "expiry_date")
        n = n + 1
        Set pf = pt.PivotFields(CStr(f))
        pf.Orientation = xlRowField
        pf.Position = n
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        pf.RepeatLabels = False
        If f = "order_date" Or f = "expiry_date" Then pf.NumberFormat = DATE_FMT
    Next f

    pt.AddDataField pt.PivotFields("remaining_quantity"), "Sum of remaining_quantity", xlSum
    pt.AddDataField pt.PivotFields("total_amount"), "Sum of total_amount", xlSum
    pt.ManualUpdate = False
    On Error Resume Next
    pt.HideValuesRow = True
    On Error GoTo 0

    wsPivot.Columns("A:F").AutoFit
    pt.TableRange1.Borders.LineStyle = xlContinuous
End Sub

' ---------------------------------------------------------------- source sheet

' Reads the used range of the active sheet (header in row 1) into a 1-based 2D array of strings,
' in the same shape the formatter expects. Returns Empty if there is no usable data.
Private Function ReadActiveSheet() As Variant
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, v As Variant
    Dim out() As String, r As Long, c As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    Set ws = ActiveSheet
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With
    If lastRow < 2 Or lastCol < 1 Then Exit Function

    v = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    ReDim out(1 To lastRow, 1 To lastCol)
    For r = 1 To lastRow
        For c = 1 To lastCol
            out(r, c) = CellText(v(r, c))
        Next c
    Next r
    ReadActiveSheet = out
End Function

' Cell value as text. Dates become "yyyy-mm-dd hh:nn:ss" (what ParseDate expects); whole numbers lose
' any exponent so long IDs (po_number, upc) keep every digit; decimals always use ".".
Private Function CellText(ByVal x As Variant) As String
    Select Case VarType(x)
        Case vbEmpty, vbNull, vbError
            CellText = ""
        Case vbDate
            CellText = Format$(x, "yyyy-mm-dd hh:nn:ss")
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            If x = Int(x) And Abs(x) < 1E+15 Then
                CellText = Format$(x, "0")
            Else
                CellText = Trim$(Str$(x))
            End If
        Case Else
            CellText = CStr(x)
    End Select
End Function

' Column number of a header name in row 1 of the data array, 0 if missing.
Private Function IndexOfCol(data As Variant, ByVal name As String) As Long
    Dim j As Long
    For j = 1 To UBound(data, 2)
        If Trim$(CStr(data(1, j))) = name Then IndexOfCol = j: Exit Function
    Next j
End Function
