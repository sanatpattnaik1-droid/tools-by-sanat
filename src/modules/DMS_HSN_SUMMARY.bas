Attribute VB_Name = "DMS_HSN_SUMMARY"
Sub DMS_HSN_SUMMARY()
Attribute DMS_HSN_SUMMARY.VB_ProcData.VB_Invoke_Func = " \n14"
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "DMS E Invoice Format")
   If answer = vbYes Then
If Range("A1").Value = "Invoice Date" And Range("B1").Value = "Distributor Code" And Range("C1").Value = "Distributor Name" And Range("D1").Value = "Salesman Name" And Range("E1").Value = "Route Name" And Range("F1").Value = "Invoice Number" And Range("G1").Value = "distr customer code" And Range("H1").Value = "cmp customer code" And Range("I1").Value = "Retailer code" And Range("J1").Value = "Retailer Name" And Range("K1").Value = "Retailer GSTIN" And Range("L1").Value = "Retailer Created Date" And Range("M1").Value = "Retailer Channel" And Range("N1").Value = "Retailer Sub Channel" And Range("O1").Value = "Retailer Group" And Range("P1").Value = "Retailer Class" And Range("Q1").Value = "Source" And Range("R1").Value = "Delivery Date" And Range("S1").Value = "HSN Code" And Range("T1").Value = "HSN Name" And Range("U1").Value = "Regional Brand Code" And Range("V1").Value = "Regional Brand Name" And Range("W1").Value = "Sub Brand Code" And Range("X1").Value = "Sub Brand Name" Then
If Range("Y1").Value = "Product Code" And Range("Z1").Value = "Product Name" And Range("AA1").Value = "Product Batch Code" And Range("AB1").Value = "Sell Rate" And Range("AC1").Value = "MRP" And Range("AD1").Value = "Invoice Quantity" And Range("AE1").Value = "Invoice Quantity In KG" And Range("AF1").Value = "Offer Quantity" And Range("AG1").Value = "Offer Quantity In KG" And Range("AH1").Value = "Gross Amount" And Range("AI1").Value = "CGST %" And Range("AJ1").Value = "CGST" And Range("AK1").Value = "SGST/UGST %" And Range("AL1").Value = "SGST/UGST" And Range("AM1").Value = "IGST %" And Range("AN1").Value = "IGST" And Range("AO1").Value = "CESS %" And Range("AP1").Value = "CESS Value" And Range("AQ1").Value = "TCS %" And Range("AR1").Value = "TCS Amount" And Range("AS1").Value = "Scheme Discount" And Range("AT1").Value = "Distributor Discount" And Range("AU1").Value = "Cash Discount" And Range("AW1").Value = "Status" And Range("C2").Value = "MAX INTERNATIONAL" Then
Application.ScreenUpdating = False
   Dim rng As Range
    Set rng = ActiveSheet.UsedRange
    For i = rng.Cells.Count To 1 Step -1
        If rng.Item(i).Value = "Cancel" Then
            rng.Item(i).EntireRow.Delete
             End If
    Next i
If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
     Range("F1").CurrentRegion.Select
Selection.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
Cells.Select
    Cells.EntireColumn.AutoFit
    
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("H:H").Select
    Range("H2").Activate
    Selection.Insert Shift:=xlToRight
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Taxable"
    Range("H2").Select
     Application.CutCopyMode = False
    
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-1]-RC[11]-RC[12]-RC[13],2)"
    Range("H2:H10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-2]*RC[-1]%,2)"
    Range("J2:J10000").Select
    Selection.FillDown
    Selection.Copy
    Range("L2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-4]*RC[-1]%,2)"
    Range("L2:L10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-6]*RC[-1]%,2)"
    Range("N2:N10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("P2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-8]*RC[-1]%,2)"
    Range("P2:P10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Q:Q").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    
   Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-9]+RC[-7]+RC[-5]+RC[-3]+RC[-1],2)"
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    Range("H1:H10000").Select
    Selection.End(xlUp).Select
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Rate"
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]+RC[3]"
    Range("H2:H10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToLeft).Select
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToLeft).Select
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    
        Range("A1:M10000").Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet0!R1C1:R10000C13", Version:=7).CreatePivotTable TableDestination:= _
        "Sheet1!R3C1", TableName:="PivotTable2", DefaultVersion:=7
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2")
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
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("HSN Code")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice Date").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Distributor Name"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice Number").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("HSN Code").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Regional Brand Name"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Invoice Quantity"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Taxable").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Rate").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CGST").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("SGST/UGST").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("IGST").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("CESS Value").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Net Amount").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Invoice Quantity"), "Sum of Invoice Quantity", _
        xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Net Amount"), "Sum of Net Amount", xlSum
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Rate")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Taxable"), "Sum of Taxable", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("IGST"), "Sum of IGST", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("CGST"), "Sum of CGST", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("SGST/UGST"), "Sum of SGST/UGST", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("CESS Value"), "Sum of CESS Value", xlSum
    ActiveWorkbook.ShowPivotTableFieldList = False
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("2:2").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Summary For HSN-"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "No. of HSN"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "HSN"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "UQC"
    Columns("D:D").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Total Quantity"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "Total Value"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "Taxable Value"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "Integrated Tax Amount"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "Central Tax Amount"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "State/UT Tax Amount"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "Cess Amount"
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "NOS-NUMBERS"
    Range("C5:C10000").Select
    Selection.FillDown
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Total Value"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Total Taxable Value"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Total Integrated Tax"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Total Central Tax"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "Total State/UT Tax"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "Total Cess"
    Columns("A:A").Select
    Range("A3").Activate
    Selection.Replace What:="GRAND TOTAL", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
         Range("A3").Activate
    Selection.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A3").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=COUNTA(R[2]C:R[10000]C)"
    Range("A3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A3").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("E3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[110000]C)"
    Range("G3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[10000]C)"
    Range("H3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[10000]C)"
    Range("I3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[10000]C)"
    Range("J3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[10000]C)"
    Range("K3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[2]C:R[10000]C)"
    Range("E3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(RC[2]:RC[6])"
    
    Range("A4:K4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A2:K2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Rows("3:3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
     ActiveCell.FormulaR1C1 = "=Sheet0!R[2]C[-1]"
    Range("B1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.NumberFormat = "mmmm-yy"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""mmM-yyyy"")"
    Range("C1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]&RC[1]"
    Range("B1").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.ClearContents
Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]"
    Range("A1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A1").Select

    Application.GoTo Reference:="R20000C1"
    Range("A20000").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",R[-1999]C),FIND(""["",CELL(""filename"",R[-1999]C))-1)"
    Range("A20000").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("A2").Select
    Range("A1:AK20000").Copy
    Workbooks.Add
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A2").Select
     Dim path As String
Dim Filename As String
path = Range("A20000")
Filename = Range("A1")
dt = Format(CStr(Now), "mmyyhhmmss")
ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx", FileFormat:=51
    Range("A20000").Select
    Selection.ClearContents
    Application.ScreenUpdating = True
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
    MsgBox "File Exported To.. " & path & Filename & dt & ""
    Range("A2").Select
    ActiveWorkbook.Close SaveChanges:=True
    Exit Sub
    End If
Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Else1: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Exit Sub
End Sub
