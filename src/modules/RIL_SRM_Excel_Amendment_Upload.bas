Attribute VB_Name = "RIL_SRM_Excel_Amendment_Upload"
Sub RIL_Excel_Amendment_Upload()
    Application.ScreenUpdating = False
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:I").Select
    Selection.Cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:K").Select
    Selection.Cut
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight
    Columns("L:O").Select
    Selection.EntireColumn.Hidden = True
    Range("P2").Select
    ActiveCell.FormulaR1C1 = _
    "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
    Range("P2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Range("P2").Select
    Selection.NumberFormat = ";;;"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("C2:P2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Range("A2").Select
    Workbooks.Add
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "PO_NO"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "LINE_ITEM"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "MRP"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "CASE_LOT"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "EAN"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "ARTICLE_CODE"
    Range("A1:F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
        
    End With
    
    Dim Filename As String
    path = Range("H2")
    Filename = "Amendment_RIL_Upload"
    dt = Format(CStr(Now), "ddmmyyyyhhmmss")
    ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
    Range("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
    MsgBox "File Exported To.. " & path & Filename & dt & ""
    ActiveWorkbook.Close SaveChanges:=True
    ActiveWorkbook.Close SaveChanges:=False
    
    
End Sub
