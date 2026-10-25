Attribute VB_Name = "RIL_PO_TO_EXCEL_PROCESS"
Sub RIL_PO_TO_EXCEL_PROCESS1()
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "RELIANCE PO TO EXCEL")
   If answer = vbYes Then
   If Range("B1").Value = "Article No. HSN/SAC Code" Then
       Application.ScreenUpdating = False
Range("B1").Select
    ActiveCell.FormulaR1C1 = "Article No."
   Range("A1:M10000").UnMerge
    Columns("E:E").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
   Dim rng As Range
    Set rng = ActiveSheet.UsedRange
    For i = rng.Cells.Count To 1 Step -1
        If rng.Item(i).Value = "Article No. HSN/SAC Code" Then
            rng.Item(i).EntireRow.Delete
             End If
    Next i
    
    Columns("I:I").Select
    Range("I10000").Activate
    Selection.Insert Shift:=xlToRight
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("I2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("I2:I10000").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]&RC[-1]"
    Range("I2:I10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToRight
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[1]C[-1]"
    Range("F2:F10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("D1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$K$109").AutoFilter Field:=3, Criteria1:="<>"
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.End(xlUp).Select
    Selection.AutoFilter
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Dim Shp As Shape
For Each Shp In ActiveSheet.Shapes
    Shp.Delete
Next Shp

    Range("G3:G10000").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("G3:G10000").Select
    Selection.Cut
    Range("G2").Select
    ActiveSheet.Paste
    Range("D1").Select
    Selection.AutoFilter
    Columns("D:D").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete

    Range("J:J,L:L,M:M").Select
    Range("M1").Activate
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Po No"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(CELL(""filename""),SEARCH(""["",CELL(""filename""))+1, SEARCH(""]"",CELL(""filename""))-SEARCH(""["",CELL(""filename""))-6)"
    Range("C2:C100").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").Select
    Range("A40").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "DC_CODE"
    Columns("H:H").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "PURCH_ORDER_DATE"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "EXPECTED_DATE"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "SUPPLIER_CODE"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "SUPPLIER_NAME"
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Cut
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "PURCH_ORDER_NUMBER"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "PURCH_ORDER_LINE"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "EAN_NO"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "      "
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "PRODUCT_DESCRIPTION"
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "VENDOR_ITEM_CODE"
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "ORDERING_UNIT"
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "QTY_PER_PACK_OF_ORDERING_UOM"
    Range("M2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[2]/RC[1]"
    Range("M2:M300").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("M:M").Select
    Selection.Replace What:="#DIV/0!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
        Range("N1").Select
    ActiveCell.FormulaR1C1 = "NO_OF_PACKS_ORDERED"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "TOTAL_QUANTITY"
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "CLOSING_STOCK"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "MRP_PER_UNIT"
    Columns("R:R").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "LANDING_COST_INCL_TAX_PER_UNIT"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "SALES_TAX"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "COST_PRICE_INCL_TAX_PER_PO_OU"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "TAX_PER"
    Range("U1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
      Cells.Select
    Selection.Style = "Normal"
    With Selection.Font
        .name = "Calibri"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Size = 11
    Selection.RowHeight = 15
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
 Range("X1").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
        Range("X1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Company"
 Dim LstRw As Long, x
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 9)) = "CHANDRIKA" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 5)) = "GIFFY" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 8)) = "HYGIENIX" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 2)) = "JO" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 8)) = "MAXKLEEN" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 8)) = "SAFEWASH" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 7)) = "SANTOOR" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 8)) = "SOFTOUCH" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 5)) = "WIPRO" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 7)) = "YARDLEY" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 5)) = "YRDLY" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 LstRw = Cells(Rows.Count, "K").End(xlUp).row
 For x = 1 To LstRw
 If UCase(Left(Cells(x, "K"), 2)) = "WP" Then
 Cells(x, "J") = "WIPRO"
 End If
 Next
 Rows("1:500").Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Store_Name"
    Columns("N:N").Select
    Selection.Cut
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Dim lastRow As Long
    lastRow = Range("E" & Rows.Count).End(xlUp).row
    For i = 2 To lastRow
        If Range("A" & i).Value = "FR66" Then
Range("B" & i).Value = "CUTTACK-FR66"
End If
If Range("A" & i).Value = "FR67" Then
Range("B" & i).Value = "PATRAPADA-FR67"
End If
If Range("A" & i).Value = "FR63" Then
Range("B" & i).Value = "ESPLANADE MALL-FR63"
End If
If Range("A" & i).Value = "FR64" Then
Range("B" & i).Value = "KHARVEL NAGAR-FR64"
End If
If Range("A" & i).Value = "FR62" Then
Range("B" & i).Value = "NAYAPALLI-FR62"
End If
If Range("A" & i).Value = "FR65" Then
Range("B" & i).Value = "PATIA-FR65"
End If
If Range("A" & i).Value = "FR68" Then
Range("B" & i).Value = "SYMPHONY MALL-FR68"
End If
If Range("A" & i).Value = "FRIC" Then
Range("B" & i).Value = "JAJPUR-FRIC"
End If
If Len(PartyField(Range("A" & i).Value, "name")) > 0 Then
Range("B" & i).Value = PartyField(Range("A" & i).Value, "name")
End If
If Range("A" & i).Value = "FRFM" Then
Range("B" & i).Value = "JHARSUGUDA-FRFM"
    End If
If Range("A" & i).Value = "FRGG" Then
Range("B" & i).Value = "KEONJHAR-FRGG"
End If
If Range("A" & i).Value = "FR69" Then
Range("B" & i).Value = "PURI-FR69"
End If
If Range("A" & i).Value = "FRFN" Then
Range("B" & i).Value = "ROURKELA-FRFN"
End If
If Range("A" & i).Value = "FRFD" Then
Range("B" & i).Value = "SAMBALPUR-FRFD"
End If
If Range("A" & i).Value = "R801" Then
Range("B" & i).Value = "RELIANCE-DC"
 End If
    Next i
    
   Dim Filename As String
path = Range("P1")
Filename = "Reliance_"
dt = Format(CStr(Now), "ddmmyyyyhhmmss")
ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
    Range("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
    MsgBox "HIGHLIGHT THE SKU(S) NEED TO BE AMENDMENT.", vbInformation, "Status"
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Exit Sub
End If
End Sub
