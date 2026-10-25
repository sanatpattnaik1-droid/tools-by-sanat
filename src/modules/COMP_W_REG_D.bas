Attribute VB_Name = "COMP_W_REG_D"
Sub COMP_W_REG_D()
Attribute COMP_W_REG_D.VB_ProcData.VB_Invoke_Func = " \n14"
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "Actoin")
   If answer = vbYes Then
 If Range("A1").Value = "Company Name  :" Then
 Application.ScreenUpdating = False
  Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Cells.Select
    Application.CutCopyMode = False
    With Selection.Font
        .name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:AI").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Item"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Invc N"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Invc No"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Pack Size"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Party Name"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "M.R.P"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Case"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Nos"
    Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("E:E").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
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
    Columns("F:F").EntireColumn.AutoFit
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 9), Array(2, 9), Array(3, 9), Array(4, 1), Array(5, 9)), _
        TrailingMinusNumbers:=True
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-2]"
    Range("E3").Select
    Selection.End(xlDown).Select
    Range("E5653").Select
    Range("J2").Select
    Selection.Cut
    Range("J1").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("J1:J5653")
    Range("J1:J5653").Select
    Selection.Cut
    Range("J2").Select
    ActiveSheet.Paste
    Application.ScreenUpdating = True
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Exit Sub
End If
End Sub
