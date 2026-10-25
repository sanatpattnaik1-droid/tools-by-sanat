Attribute VB_Name = "IRN_UPLOAD"
Sub IRN_UPLOAD()
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "HA IRN Upload Format")
   If answer = vbYes Then
   If Range("A1").Value = "Sl. No" And Range("B1").Value = "Ack No" And Range("C1").Value = "Ack Date" And Range("D1").Value = "Doc No" And Range("E1").Value = "Doc Date" And Range("F1").Value = "Doc Type" And Range("G1").Value = "Inv Value." And Range("H1").Value = "Recipient GSTIN" And Range("I1").Value = "Status" And Range("J1").Value = "IRN" And Range("K1").Value = "SignedQrCOde" Then
   Application.ScreenUpdating = False
 Range("L1").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
        Range("L1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
          Range("L1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        End With
        Range("A1").Select
       Range("A1:L600").Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I4").Select
    Selection.Copy
    Application.CutCopyMode = False
    With Range("A2:A" & Range("B" & Rows.Count).End(xlUp).row)
    .Cells(1, 1).Value = 1
    .DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=1, Trend:=False
  End With
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Cut
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Sl. No"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "IRN"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Ack No"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Ack Date"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Doc No"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Doc Typ"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Doc Date"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Inv Value."
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Recipient GSTIN"
    Columns("J:J").Select
    Selection.ClearContents
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Status"
    Range("J2").Select
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Signed QR Code"
    Range("A2").Select
    Range("A2").Select
    Dim path As String
Dim Filename As String
path = Range("L1")
Filename = "Upload_"
dt = Format(CStr(Now), "ddmmyyyyhhmmss")
ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
    Range("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
    MsgBox "File Exported To.. " & path & Filename & dt & ""
    ActiveWorkbook.Close SaveChanges:=True
    Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
    End If
Exit Sub
End If
End Sub

