Attribute VB_Name = "Wisdom_Stkreport"
Sub Wisdom_Stkreport()

    Dim wsSrc As Worksheet, wsPivot As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim pCache As PivotCache
    Dim pTable As PivotTable

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '==============================
    ' SOURCE SHEET CLEANUP
    '==============================
    Set wsSrc = ActiveSheet

    ' Delete first 4 rows
    wsSrc.Rows("1:4").Delete

    ' Remove unwanted columns
    wsSrc.Columns("B:C").Delete
    wsSrc.Columns("I:L").Delete

    ' Fix header
    wsSrc.Range("G1").Value = "Case"
    wsSrc.Range("H1").Value = "Nos"

    ' Remove merged cells
    wsSrc.Range("G1:H1").UnMerge

    ' Last row / column
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    ' Filter "Total :"
    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).AutoFilter _
        Field:=4, Criteria1:="Total :"

    '==============================
    ' CREATE PIVOT SHEET
    '==============================
    Set wsPivot = Sheets.Add
    wsPivot.name = "Stock_Pivot"

    Set pCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)) _
    )

    Set pTable = pCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="Pivot_Stock_Report" _
    )

    '==============================
    ' PIVOT SETTINGS
    '==============================
    With pTable
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ColumnGrand = True
        .RowGrand = True
    End With

    '==============================
    ' ROW FIELDS
    '==============================
    With pTable.PivotFields("Description")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array(False, False, False, False, False, False, _
                           False, False, False, False, False, False)
    End With

    '==============================
    ' DATA FIELDS
    '==============================
    pTable.AddDataField pTable.PivotFields("MRP"), "Sum of MRP", xlSum
    pTable.AddDataField pTable.PivotFields("Case"), "Sum of Case", xlSum
    pTable.AddDataField pTable.PivotFields("Nos"), "Sum of Nos", xlSum

    '==============================
    ' FINAL TOUCH
    '==============================
    wsPivot.Columns.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Stock Pivot Report Created Successfully!", vbInformation

End Sub


