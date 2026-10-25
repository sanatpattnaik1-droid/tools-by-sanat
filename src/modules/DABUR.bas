Attribute VB_Name = "DABUR"
Sub DABUR_E_INVOICE()
Attribute DABUR_E_INVOICE.VB_ProcData.VB_Invoke_Func = " \n14"

Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "DMS E Invoice Format")
   If answer = vbYes Then
   
   If Range("A1").Value = "BillWise ProductWise Sales Report" And Range("E3").Value = "Salesman Code" And Range("D3").Value = "PO Date" Then
   
Application.ScreenUpdating = False
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "B2B"
    Range("A2:A10000").Select
    Selection.FillDown
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "No"
    Range("D2:D10000").Select
    Selection.FillDown
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Tax Invoice"
    Range("E2:E10000").Select
    Selection.FillDown
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToRight
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Columns("AH:AH").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("AI:AI").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToRight
    Columns("AG:AG").Select
    Selection.Cut
    Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight
    Range("AG2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(R2C6:RC[-27],RC[-27])"
    Range("AG2:AG10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     Range("AI2").Select
    ActiveCell.FormulaR1C1 = "No"
    Range("AI2:AI10000").Select
    Selection.FillDown
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight
    Columns("AN:AN").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToRight
    
    Range("AN2").Select
    ActiveCell.FormulaR1C1 = "PIECES"
    Range("AN2:AN10000").Select
    Selection.FillDown
    
    Columns("AO:AO").Select
    Selection.Insert Shift:=xlToRight
    Range("AO2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[1]/RC[-3],2)"
    Range("AO2:AO10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Replace What:="#DIV/0!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "Unit Price"
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("AP2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[5]-RC[7]-RC[9]-RC[11]-RC[13],2)"
    Range("AP2:AP10000").Select
    Application.CutCopyMode = False
    Range("AP2:AP10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AS2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("AT2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[2]+RC[4]++RC[6]"
    Range("AT2:AT10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("AV:AV").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AW:AW").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AV:AW").Select
    Selection.Cut
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight
    Columns("AW:AW").Select
    Selection.Insert Shift:=xlToRight
    Range("AW2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[-2],2)"
    Range("AW2:AW10000").Select
    Selection.FillDown
    Selection.Copy
    Range("AU2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AV2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("AY:AY").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("AW2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(RC[2],2)"
    Range("AW2:AW10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("AZ:BA").Select
    Selection.Cut
    Columns("AX:AX").Select
    Selection.Insert Shift:=xlToRight
    Columns("AZ:AZ").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Columns("BF:BF").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Range("BE1").Select
    ActiveCell.FormulaR1C1 = "Item Total"
    Range("BI1").Select
    ActiveCell.FormulaR1C1 = "Total Taxable"
    Range("BJ1").Select
    ActiveCell.FormulaR1C1 = "SGST"
    Range("BK1").Select
    ActiveCell.FormulaR1C1 = "CGST"
    Range("BL1").Select
    ActiveCell.FormulaR1C1 = "IGST"
    Range("BM1").Select
    ActiveCell.FormulaR1C1 = "CESS"
    Range("BN1").Select
    ActiveCell.FormulaR1C1 = "STATE CESS"
    Range("BO1").Select
    ActiveCell.FormulaR1C1 = "DISS"
    Range("BP1").Select
    ActiveCell.FormulaR1C1 = "OTHER"
    Range("BQ1").Select
    ActiveCell.FormulaR1C1 = "ROUND OFF"
    Range("BR1").Select
    ActiveCell.FormulaR1C1 = "Total Invoice"
    Range("BF2").Select
Range("BI2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(R2C[-55]:R10000C[-55],RC[-55],R2C[-16]:R10000C[-16])"
    Range("BI2:BI10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BJ2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=TRUNC(SUMIF(R2C[-56]:R10000C[-56],RC[-56],R2C[-15]:R10000C[-15]),2)"
    Range("BJ2:BJ10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BK2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("BL2").Select
    ActiveCell.FormulaR1C1 = _
        "=TRUNC(SUMIF(R2C[-58]:R10000C[-58],RC[-58],R2C[-15]:R10000C[-15]),2)"
    Range("BL2:BL10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BR2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=TRUNC(SUMIF(R2C[-64]:R10000C[-64],RC[-64],R2C[-13]:R10000C[-13]),2)"
    Range("BR2:BR10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("BQ2").Select
    ActiveCell.FormulaR1C1 = "=TRUNC(ROUND(RC[1],0)-RC[1],2)"
    Range("BQ2:BQ10000").Select
    Selection.FillDown
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    
    Columns("F:F").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
    Rows("2:10000").Select
    Selection.Copy
    Workbooks.Add
    Range("A2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        Range("A1").Select
ActiveCell.FormulaR1C1 = "Supply Type Code *"
Range("B1").Select
ActiveCell.FormulaR1C1 = "Reverse Charge"
Range("C1").Select
ActiveCell.FormulaR1C1 = "e-Comm GSTIN"
Range("D1").Select
ActiveCell.FormulaR1C1 = "Igst On Intra"
Range("E1").Select
ActiveCell.FormulaR1C1 = "Document Type *"
Range("F1").Select
ActiveCell.FormulaR1C1 = "Document Number *"
Range("G1").Select
ActiveCell.FormulaR1C1 = "Document Date (DD/MM/YYYY) *"
Range("H1").Select
ActiveCell.FormulaR1C1 = "Buyer GSTIN *"
Range("I1").Select
ActiveCell.FormulaR1C1 = "Buyer Legal Name *"
Range("J1").Select
ActiveCell.FormulaR1C1 = "Buyer Trade Name "
Range("K1").Select
ActiveCell.FormulaR1C1 = "Buyer POS *"
Range("L1").Select
ActiveCell.FormulaR1C1 = "Buyer Addr1 *"
Range("M1").Select
ActiveCell.FormulaR1C1 = "Buyer Addr2"
Range("N1").Select
ActiveCell.FormulaR1C1 = "Buyer Location *"
Range("O1").Select
ActiveCell.FormulaR1C1 = "Buyer Pin Code"
Range("P1").Select
ActiveCell.FormulaR1C1 = "Buyer State *"
Range("Q1").Select
ActiveCell.FormulaR1C1 = "Buyer Phone Number"
Range("R1").Select
ActiveCell.FormulaR1C1 = "Buyer Email Id"
Range("S1").Select
ActiveCell.FormulaR1C1 = "Dispatch Name "
Range("T1").Select
ActiveCell.FormulaR1C1 = "Dispatch Addr1 "
Range("U1").Select
ActiveCell.FormulaR1C1 = "Dispatch Addr2"
Range("V1").Select
ActiveCell.FormulaR1C1 = "Dispatch Location "
Range("W1").Select
ActiveCell.FormulaR1C1 = "Dispatch Pin Code "
Range("X1").Select
ActiveCell.FormulaR1C1 = "Dispatch State "
Range("Y1").Select
ActiveCell.FormulaR1C1 = "Shipping GSTIN "
Range("Z1").Select
ActiveCell.FormulaR1C1 = "Shipping Legal Name "
Range("AA1").Select
ActiveCell.FormulaR1C1 = "Shipping Trade Name "
Range("AB1").Select
ActiveCell.FormulaR1C1 = "Shipping Addr1 "
Range("AC1").Select
ActiveCell.FormulaR1C1 = "Shipping Addr2"
Range("AD1").Select
ActiveCell.FormulaR1C1 = "Shipping Location "
Range("AE1").Select
ActiveCell.FormulaR1C1 = "Shipping Pin Code "
Range("AF1").Select
ActiveCell.FormulaR1C1 = "Shipping State "
Range("AG1").Select
ActiveCell.FormulaR1C1 = "Sl.No. *"
Range("AH1").Select
ActiveCell.FormulaR1C1 = "Product Description"
Range("AI1").Select
ActiveCell.FormulaR1C1 = "Is_Service *"
Range("AJ1").Select
ActiveCell.FormulaR1C1 = "HSN code *"
Range("AK1").Select
ActiveCell.FormulaR1C1 = "Bar code "
Range("AL1").Select
ActiveCell.FormulaR1C1 = "Quantity *"
Range("AM1").Select
ActiveCell.FormulaR1C1 = "Free Quantity"
Range("AN1").Select
ActiveCell.FormulaR1C1 = "Unit *"
Range("AO1").Select
ActiveCell.FormulaR1C1 = "Unit Price *"
Range("AP1").Select
ActiveCell.FormulaR1C1 = "Gross Amount *"
Range("AQ1").Select
ActiveCell.FormulaR1C1 = "Discount"
Range("AR1").Select
ActiveCell.FormulaR1C1 = "Pre Tax Value"
Range("AS1").Select
ActiveCell.FormulaR1C1 = "Taxable value *"
Range("AT1").Select
ActiveCell.FormulaR1C1 = "GST Rate (%) *"
Range("AU1").Select
ActiveCell.FormulaR1C1 = "Sgst Amt(Rs)"
Range("AV1").Select
ActiveCell.FormulaR1C1 = "Cgst Amt (Rs)"
Range("AW1").Select
ActiveCell.FormulaR1C1 = "Igst Amt (Rs)"
Range("AX1").Select
ActiveCell.FormulaR1C1 = "Cess Rate (%)"
Range("AY1").Select
ActiveCell.FormulaR1C1 = "Cess Amt Adval (Rs)"
Range("AZ1").Select
ActiveCell.FormulaR1C1 = "Cess Non Adval Amt (Rs)"
Range("BA1").Select
ActiveCell.FormulaR1C1 = "State Cess Rate (%)"
Range("BB1").Select
ActiveCell.FormulaR1C1 = "State Cess Adval Amt (Rs)"
Range("BC1").Select
ActiveCell.FormulaR1C1 = "State Cess Non-Adval Amt (Rs)"
Range("BD1").Select
ActiveCell.FormulaR1C1 = "Other Charges  "
Range("BE1").Select
ActiveCell.FormulaR1C1 = "Item Total *"
Range("BF1").Select
ActiveCell.FormulaR1C1 = "Batch Name"
Range("BG1").Select
ActiveCell.FormulaR1C1 = "Batch Expiry Dt"
Range("BH1").Select
ActiveCell.FormulaR1C1 = "Warranty Dt"
Range("BI1").Select
ActiveCell.FormulaR1C1 = "Total Taxable value *"
Range("BJ1").Select
ActiveCell.FormulaR1C1 = "Sgst Amt"
Range("BK1").Select
ActiveCell.FormulaR1C1 = "Cgst Amt"
Range("BL1").Select
ActiveCell.FormulaR1C1 = "Igst Amt"
Range("BM1").Select
ActiveCell.FormulaR1C1 = "Cess Amt"
Range("BN1").Select
ActiveCell.FormulaR1C1 = "State Cess Amt"
Range("BO1").Select
ActiveCell.FormulaR1C1 = "Discount"
Range("BP1").Select
ActiveCell.FormulaR1C1 = "Other charges"
Range("BQ1").Select
ActiveCell.FormulaR1C1 = "Round off"
Range("BR1").Select
ActiveCell.FormulaR1C1 = "Total Invoice value *"
Range("BS1").Select
ActiveCell.FormulaR1C1 = "Shipping Bill No"
Range("BT1").Select
ActiveCell.FormulaR1C1 = "Shipping Bill Dt"
Range("BU1").Select
ActiveCell.FormulaR1C1 = "Port"
Range("BV1").Select
ActiveCell.FormulaR1C1 = "Refund claim"
Range("BW1").Select
ActiveCell.FormulaR1C1 = "Foreign Currency "
Range("BX1").Select
ActiveCell.FormulaR1C1 = "Country Code "
Range("BY1").Select
ActiveCell.FormulaR1C1 = "Export Duty Amount"
Range("BZ1").Select
ActiveCell.FormulaR1C1 = "Trans ID"
Range("CA1").Select
ActiveCell.FormulaR1C1 = "Trans Name"
Range("CB1").Select
ActiveCell.FormulaR1C1 = "Trans Mode "
Range("CC1").Select
ActiveCell.FormulaR1C1 = "Distance "
Range("CD1").Select
ActiveCell.FormulaR1C1 = "Trans Doc No."
Range("CE1").Select
ActiveCell.FormulaR1C1 = "Trans Doc Date"
Range("CF1").Select
ActiveCell.FormulaR1C1 = "Vehicle No."
Range("CG1").Select
ActiveCell.FormulaR1C1 = "Vehicle Type"
Range("CH1").Select
ActiveCell.FormulaR1C1 = "Error List"
Application.ScreenUpdating = True



Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Else1: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
End If
Exit Sub
End Sub
