Attribute VB_Name = "Wipro_Ecom_Report"
Sub Macro1()

'Wipro_Ecom_Report

    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.AutoFilter
    Columns("C:C").Select
    Range("C8").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Columns("J:J").Select
    Range("J16").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Range("E2").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("D:H").Select
    Range("D2").Activate
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Columns("B:B").Select
    Range("B10").Activate
    Selection.Replace What:="/24-25", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Columns("I:I").Select
    Range("I15").Activate
    Selection.Insert Shift:=xlToRight
    Columns("H:H").Select
    Range("H14").Activate
    Selection.TextToColumns Destination:=Range("H1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=",", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Selection.End(xlUp).Select
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "FSN"
    Columns("M:M").Select
    Range("M12").Activate
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("P1").Select
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    Range("O1").Select
    ActiveSheet.Paste
    Range("N1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Discount%"
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Rows("1:1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort.SortFields _
        .Add2 Key:=Range("A1"), SortOn:=xlSortOnValues, order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort.SortFields _
        .Add2 Key:=Range("B1"), SortOn:=xlSortOnValues, order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sale Bill Register (96)").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B2").Select
End Sub


