Attribute VB_Name = "RIL_SRM_Excel_Amendment_Email"
Sub RIL__Excel_Amendment_Email()
    If Range("D2").Value = "" Then
        MsgBox "Po Expire Date Not Avaialeble Please Enter PO Expire Date and Try Again", vbInformationm, "Error"
    Else
        Application.ScreenUpdating = False
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns("F:F").Select
        Selection.Cut
        Columns("D:D").Select
        Selection.Insert Shift:=xlToRight
        Columns("E:E").Select
        Range("E3").Activate
        Selection.Delete Shift:=xlToLeft
        Columns("G:G").Select
        Selection.Cut
        Columns("E:E").Select
        Selection.Insert Shift:=xlToRight
        Columns("J:J").Select
        Selection.Cut
        Columns("G:G").Select
        Selection.Insert Shift:=xlToRight
        Columns("H:H").Select
        Selection.Insert Shift:=xlToRight
        Columns("L:N").Select
        Selection.Cut
        Columns("I:I").Select
        Selection.Insert Shift:=xlToRight
        Columns("L:L").Select
        Selection.Insert Shift:=xlToRight
        Columns("N:N").Select
        Selection.Insert Shift:=xlToRight
        Columns("O:O").Select
        Selection.ClearContents
        ActiveWindow.SmallScroll ToRight:=5
        Columns("Q:Q").Select
        Selection.Delete Shift:=xlToLeft
        Range("Q2").Select
        ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
        Range("Q2").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("Q2").Select
        Selection.NumberFormat = ";;;"
        Range("A2").Select
        Range("C2:Q10000").Copy
        Workbooks.Add
        Range("D2").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "Date"
        Range("B1").Select
        ActiveCell.FormulaR1C1 = "Rrfmcg"
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "Requestor Email Id"
        Range("A2").Select
        ActiveCell.FormulaR1C1 = "=TODAY()"
        Range("B2").Select
        ActiveCell.FormulaR1C1 = "Rrfmcg"
        Range("C2").Select
        ActiveCell.FormulaR1C1 = SettingValue("requestor_email")
        Range("D1").Select
        ActiveCell.FormulaR1C1 = "Po No"
        Range("E1").Select
        ActiveCell.FormulaR1C1 = "Po Line Item"
        Range("F1").Select
        ActiveCell.FormulaR1C1 = "Article Code"
        Range("G1").Select
        ActiveCell.FormulaR1C1 = "Del Date(Dd.Mm.Yyyy)"
        Range("H1").Select
        ActiveCell.FormulaR1C1 = "Old Mrp"
        Range("I1").Select
        ActiveCell.FormulaR1C1 = "New Mrp"
        Range("J1").Select
        ActiveCell.FormulaR1C1 = "Req Units In One Case"
        Range("K1").Select
        ActiveCell.FormulaR1C1 = "Nos Of Case"
        Range("L1").Select
        ActiveCell.FormulaR1C1 = "Total Qty"
        Range("M1").Select
        ActiveCell.FormulaR1C1 = "Gross Wt. of Case Lot"
        Range("N1").Select
        ActiveCell.FormulaR1C1 = "Old EAN"
        Range("O1").Select
        ActiveCell.FormulaR1C1 = "New EAN"
        Range("P1").Select
        ActiveCell.FormulaR1C1 = "Del Compl flag"
        Columns("N:N").Select
        Selection.Copy
        Columns("O:O").Select
        Selection.Insert Shift:=xlToRight
        Range("O1").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "New EAN"
        Columns("P:P").Select
        Selection.Delete Shift:=xlToLeft
        Range("A1:P1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
        Selection.Font.Bold = True
        Range("A2:C10000").Select
        Selection.FillDown
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Columns("D:D").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Delete
        Dim Filename As String
        path = Range("R2")
        Filename = "Amendment_RIL_Email"
        dt = Format(CStr(Now), "ddmmyyyyhhmmss")
        ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
        Range("R:R").Select
        Selection.Delete Shift:=xlToLeft
        Range("A1").Select
        ActiveWorkbook.Save
        Application.ScreenUpdating = True
        MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
        MsgBox "File Exported To.. " & path & Filename & dt & ""
        ActiveWorkbook.Close SaveChanges:=True
    End If
End Sub
