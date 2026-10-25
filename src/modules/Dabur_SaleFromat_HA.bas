Attribute VB_Name = "Dabur_SaleFromat_HA"
Sub Dabur_SaleFromat_HA()
Attribute Dabur_SaleFromat_HA.VB_ProcData.VB_Invoke_Func = " \n14"
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "DMS E Invoice Format")
   If answer = vbYes Then
If Range("A3").Value = "Bill Number" And Range("B3").Value = "Bill Date" And Range("C3").Value = "PO Number" And Range("D3").Value = "PO Date" And Range("E3").Value = "Salesman Code" And Range("F3").Value = "Salesman Name" And Range("N3").Value = "MRP" And Range("O3").Value = "LSP Rate" And Range("P3").Value = "Selling Rate" And Range("U3").Value = "Sales Value" And Range("Y3").Value = "Total Discount" And Range("Z3").Value = "Total Tax" And Range("AA3").Value = "Net Amount" Then
Application.ScreenUpdating = False
If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Range("AU1").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
        Range("AU1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
         Range("AU1").Select
    Selection.NumberFormat = ";;;"
    
    
    Columns("D:D").Select
    Range("D4").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Range("E3").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Range("F3").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Range("G3").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Range("H3").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Range("J3").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Range("K3").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Range("L3").Activate
    Selection.Insert Shift:=xlToRight
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "Taxable"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[2]-RC[1],2)"
    Selection.End(xlDown).Select
    Range("L5000").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.FillDown
    Selection.Copy
    Columns("L:L").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*2"
    Range("P6").Select
    Selection.End(xlDown).Select
    Range("P5000").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("P2:P5000").Select
    Range("P5000").Activate
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A5000").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("E5000").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("R5000").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R5000C15", Version:=8).CreatePivotTable TableDestination:= _
        "Sheet2!R3C1", TableName:="PivotTable1", DefaultVersion:=8
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Bill Number")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("PO Number")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Retailer Name")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Product Name")
        .Orientation = xlRowField
        .Position = 4
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("MRP")
        .Orientation = xlRowField
        .Position = 5
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Sales Qty"), "Sum of Sales Qty", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Taxable"), "Sum of Taxable", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Net Amount"), "Sum of Net Amount", xlSum
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Bill Number").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Bill Date").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("PO Number").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Retailer Name").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Product Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("MRP").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Selling Rate").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Sales Qty").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Free Qty").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Sales Value").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Total Discount").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Taxable").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Total Tax").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Net Amount").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("CGST Tax").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A5").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Bill Number").Subtotals = _
        Array(True, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotSelect "'Bill Number'[All;Total]", _
        xlDataAndLabel, True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A4").Select
    ActiveWindow.FreezePanes = True
    ActiveWorkbook.ShowPivotTableFieldList = False
        Application.ScreenUpdating = True
        
        Dim Filename As String
path = Sheets("Sheet1").Range("Z1")
Filename = "Dabur_Salefromat_HA"
dt = Format(CStr(Now), "dd-mmm-yyhhmmss")
ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
        
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
    MsgBox "File Exported To.. " & path & Filename & dt & ""
    Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Exit Sub
End If
End Sub
