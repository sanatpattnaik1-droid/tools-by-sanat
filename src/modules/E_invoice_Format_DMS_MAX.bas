Attribute VB_Name = "E_invoice_Format_DMS_MAX"
Sub E_invoice_Format_DMS_MAX()
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
Columns("AQ:AR").Select
    Selection.Delete Shift:=xlToLeft
If ActiveSheet.AutoFilterMode Then
     ActiveSheet.AutoFilterMode = False
     End If
     Range("F1").CurrentRegion.Select
Selection.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Supply Type Code *"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Document Type *"
    Columns("H:H").Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Cut
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Buyer Trade Name"
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Buyer POS *"
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Buyer Addr1 *"
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Buyer Addr2"
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Buyer Location *"
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Buyer Pin Code"
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Buyer State *"
    Columns("N:N").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Sl.No. *"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(R2C3:RC[-11],RC[-11])"
    Range("N2:N10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("N5").Select
    Application.CutCopyMode = False
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("Q:Q").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("R:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("S:S").Select
    Selection.Delete Shift:=xlToLeft
    Range("W1").Select
    Columns("O:O").Select
    Selection.Cut
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Is_Service *"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "No"
    Range("Q2:Q10000").Select
    Selection.FillDown
    Selection.End(xlUp).Select
    Columns("Q:Q").Select
    Selection.Cut
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Free Quantity"
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Unit *"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "Pieces"
    Range("T2:T10000").Select
    Selection.FillDown
    Selection.End(xlUp).Select
    Columns("U:U").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Unit Price *"
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Taxable value *"
    Columns("X:X").Select
    Selection.Insert Shift:=xlToRight
    Range("W2").Select
    Application.CutCopyMode = False
    
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-1]-RC[10]-RC[11]-RC[12],2)"
    Range("W2:W10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Selection.Copy
        Range("V2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("U2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[2]/RC[-3],2)"
    Range("U2:U10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
         Columns("U:U").Select
    Selection.Replace What:="#DIV/0!", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
      Range("Z2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-3]*RC[-1]%,2)"
    Range("Z2:Z10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.Copy
        
        Range("AB2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("X1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "GST Rate"
    Range("X2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[1]+RC[3]"
    Range("X2:X10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Z2").Select
    Application.CutCopyMode = False
    Columns("Y:Y").Select
    Selection.Delete Shift:=xlToLeft
    Columns("Z:Z").Select
    Selection.Delete Shift:=xlToLeft
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = "SGST"
    
    
  Range("AB2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-5]*RC[-1]%,2)"
    Range("AB2:AB10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AD2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-7]*RC[-1]%,2)"
    Range("AD2:AD10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    Columns("AE:AE").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("AF:AF").Select
    Selection.Delete Shift:=xlToLeft
    Range("AE1").Select
    ActiveCell.FormulaR1C1 = "Item Total"
    
    
 Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-8]+RC[-6]+RC[-5]+RC[-3]+RC[-1],2)"
    Range("AE2:AE10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        
        
    Range("AF1").Select
    ActiveCell.FormulaR1C1 = "Total Taxable value*"
    Range("AG1").Select
    ActiveCell.FormulaR1C1 = "Sgst Amt"
    Range("AH1").Select
    ActiveCell.FormulaR1C1 = "Cgst Amt"
    Range("AI1").Select
    ActiveCell.FormulaR1C1 = "Igst Amt"
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = "Cess"
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "Round off"
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "Total Invoice value *"
    Range("AF2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-29]:R10000C[-29],RC[-29],R2C[-9]:R10000C[-9])"
    Range("AF2:AF10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AG2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-30]:R10000C[-30],RC[-30],R2C[-7]:R10000C[-7])"
    Range("AG2:AG10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AH2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-31]:R10000C[-31],RC[-31],R2C[-9]:R10000C[-9])"
    Range("AH2:AH10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AI2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-32]:R10000C[-32],RC[-32],R2C[-7]:R10000C[-7])"
    Range("AI2:AI10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AJ2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-33]:R10000C[-33],RC[-33],R2C[-6]:R10000C[-6])"
    Range("AJ2:AJ10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AL2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-35]:R10000C[-35],RC[-35],R2C[-7]:R10000C[-7])"
    Range("AL2:AL10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False_
    
    
     Range("AK2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(ROUND(RC[1],0)-RC[1],2)"
    Range("AK2:AK10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = "=ROUND(SUM(RC[-6]:RC[-1]),0)"
    Range("AL2:AL10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.End(xlUp).Select
    

    
    Range("A1:AL1").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    
        Dim lastRow As Long
    lastRow = Range("E" & Rows.Count).End(xlUp).row
    Dim partyKey As String
    For i = 2 To lastRow
        partyKey = Trim$(CStr(Range("E" & i).Value))
        If Len(PartyField(partyKey, "name")) > 0 Then
            Range("G" & i).Value = PartyField(partyKey, "name")
            Range("H" & i).Value = PartyField(partyKey, "state")
            Range("I" & i).Value = PartyField(partyKey, "address1")
            Range("J" & i).Value = PartyField(partyKey, "address2")
            Range("K" & i).Value = PartyField(partyKey, "city")
            Range("L" & i).Value = PartyField(partyKey, "pin")
            Range("M" & i).Value = PartyField(partyKey, "state")
        End If
    Next i
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "B2B"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Tax Invoice"
    Range("A2:B10000").Select
    Selection.FillDown
Columns("C:C").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Range("A2").Select
    Application.GoTo Reference:="R20000C1"
    Range("A20000").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(CELL(""filename"",R[-1999]C),FIND(""["",CELL(""filename"",R[-1999]C))-1)"
    Range("A20000").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("A2").Select
    Range("A1:AL20000").Copy
    Workbooks.Add
    ActiveSheet.Paste
     Dim path As String
Dim Filename As String
path = Range("A20000")
Filename = "E_Invoice_Data_"
dt = Format(CStr(Now), "ddmmyyyyhhmmss")
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
Exit Sub
End If
Else1: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
Exit Sub
End If
End Sub
