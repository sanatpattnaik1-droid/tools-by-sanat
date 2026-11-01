Attribute VB_Name = "modGuard"
Option Explicit
Option Compare Text

' Keeps a ribbon tool from running on the wrong file (for example the Flipkart tool on a Blinkit export).
' Before a ribbon button runs its macro, modRibbon.RunMacro asks CanRun(tag). The ACTIVE sheet must carry the
' column headings the tool was built for; if it does not, the user gets a plain explanation and nothing runs.
' Tools that work on any sheet (Fill Down, Convert Cells, ...) are simply not listed here.
'
' Signature text: items separated by ";"
'   A5=PARTY_NAME     cell A5 must say PARTY_NAME (Trim, ignores case; a "*" in the text acts as a wildcard)
'   @1:po_number      the heading po_number must appear somewhere in row 1
' A tool may have several alternative signatures (several Reg lines with the same tag): one must match.
' strict=False means the headings are a best guess: the user is asked "Run anyway?" instead of being stopped.

Private Const TITLE As String = "Tools By Sanat"
Private Const MAX_SHOWN As Long = 4

Private mCount As Long
Private mTag() As String
Private mName() As String
Private mStrict() As Boolean
Private mSig() As String

' True when the active sheet is the file this ribbon tool was built for (or the tool has no check).
' Otherwise explains why and returns False: the caller must not run the tool.
Public Function CanRun(ByVal tag As String) As Boolean
    Dim ws As Worksheet, bad As Long, detail As String, toolName As String, strict As Boolean
    Dim msg As String, other As String

    If Not Describe(tag, toolName, strict) Then CanRun = True: Exit Function

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open the file you want to process first, then click the button.", vbExclamation, TITLE
        Exit Function
    End If
    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Select the worksheet you want to process first, then click the button.", vbExclamation, TITLE
        Exit Function
    End If
    Set ws = ActiveSheet

    bad = Compare(ws, tag, detail)
    If bad = 0 Then CanRun = True: Exit Function

    msg = "This sheet does not look like the file for """ & toolName & """:" & detail
    If bad > MAX_SHOWN Then msg = msg & vbLf & " - ...and " & (bad - MAX_SHOWN) & " more"
    msg = msg & vbLf & vbLf & "Sheet: [" & ws.Parent.Name & "] " & ws.Name
    other = LooksLike(ws, tag)
    If Len(other) > 0 Then msg = msg & vbLf & "It looks like the file for: " & other

    If strict Then
        MsgBox msg & vbLf & vbLf & "Nothing was changed.", vbExclamation, TITLE
    Else
        CanRun = (MsgBox(msg & vbLf & vbLf & "Run anyway?", vbExclamation + vbYesNo + vbDefaultButton2, TITLE) = vbYes)
    End If
End Function

' Silent version for tests: does this sheet satisfy one of the tool's signatures?
Public Function Matches(ws As Worksheet, ByVal tag As String) As Boolean
    Dim detail As String
    Matches = (Compare(ws, tag, detail) = 0)
End Function

Private Function Describe(ByVal tag As String, ByRef toolName As String, ByRef strict As Boolean) As Boolean
    Dim i As Long
    LoadTable
    For i = 1 To mCount
        If StrComp(mTag(i), tag, vbTextCompare) = 0 Then
            toolName = mName(i): strict = mStrict(i): Describe = True: Exit Function
        End If
    Next i
End Function

' Failed checks of the best-fitting signature for this tool (0 = matches, -1 = tool has no signature).
' detail gets the first few problems, one per line.
Private Function Compare(ws As Worksheet, ByVal tag As String, ByRef detail As String) As Long
    Dim i As Long, bad As Long, d As String
    Compare = -1
    LoadTable
    For i = 1 To mCount
        If StrComp(mTag(i), tag, vbTextCompare) = 0 Then
            bad = Failures(ws, mSig(i), d)
            If Compare < 0 Or bad < Compare Then Compare = bad: detail = d
            If bad = 0 Then Exit For
        End If
    Next i
End Function

Private Function Failures(ws As Worksheet, ByVal sig As String, ByRef detail As String) As Long
    Dim items() As String, i As Long, it As String, p As Long, want As String, got As String, row As Long
    Dim ok As Boolean, line As String
    detail = ""
    items = Split(sig, ";")
    For i = LBound(items) To UBound(items)
        it = items(i)
        If Left$(it, 1) = "@" Then
            p = InStr(it, ":"): row = CLng(Mid$(it, 2, p - 2)): want = Mid$(it, p + 1)
            ok = RowHas(ws, row, want)
            line = " - heading """ & want & """ not found in row " & row
        Else
            p = InStr(it, "="): want = Mid$(it, p + 1)
            got = CellText(ws, Left$(it, p - 1))
            ok = Same(got, want)
            If Len(got) = 0 Then
                line = " - cell " & Left$(it, p - 1) & " should say """ & want & """ but is empty"
            Else
                line = " - cell " & Left$(it, p - 1) & " should say """ & want & """ but says """ & Left$(got, 30) & """"
            End If
        End If
        If Not ok Then
            Failures = Failures + 1
            If Failures <= MAX_SHOWN Then detail = detail & vbLf & line
        End If
    Next i
End Function

' Names of the other tools whose strict signature this sheet satisfies.
Private Function LooksLike(ws As Worksheet, ByVal exceptTag As String) As String
    Dim i As Long, d As String, names As String
    For i = 1 To mCount
        If mStrict(i) And StrComp(mTag(i), exceptTag, vbTextCompare) <> 0 Then
            If Failures(ws, mSig(i), d) = 0 Then
                If InStr(1, "; " & names & ";", "; " & mName(i) & ";", vbTextCompare) = 0 Then
                    If Len(names) > 0 Then names = names & "; "
                    names = names & mName(i)
                End If
            End If
        End If
    Next i
    LooksLike = names
End Function

Private Function CellText(ws As Worksheet, ByVal addr As String) As String
    Dim v As Variant
    On Error Resume Next
    v = ws.Range(addr).Value2
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

Private Function RowHas(ws As Worksheet, ByVal row As Long, ByVal want As String) As Boolean
    Dim lastC As Long, c As Long, v As Variant
    lastC = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    If lastC = 1 Then
        RowHas = Same(CellText(ws, "A" & row), want)
        Exit Function
    End If
    v = ws.Range(ws.Cells(row, 1), ws.Cells(row, lastC)).Value2
    For c = 1 To lastC
        If Not IsError(v(1, c)) Then
            If Same(Trim$(CStr(v(1, c))), want) Then RowHas = True: Exit Function
        End If
    Next c
End Function

Private Function Same(ByVal got As String, ByVal want As String) As Boolean
    If InStr(want, "*") > 0 Then
        Same = (got Like want)
    Else
        Same = (StrComp(got, want, vbTextCompare) = 0)
    End If
End Function

Private Sub Reg(ByVal tag As String, ByVal toolName As String, ByVal strict As Boolean, ByVal sig As String)
    mCount = mCount + 1
    ReDim Preserve mTag(1 To mCount): ReDim Preserve mName(1 To mCount)
    ReDim Preserve mStrict(1 To mCount): ReDim Preserve mSig(1 To mCount)
    mTag(mCount) = tag: mName(mCount) = toolName: mStrict(mCount) = strict: mSig(mCount) = sig
End Sub

' Tag = the ribbon button's tag (Module.Sub). The headings come from each macro's own column check.
Private Sub LoadTable()
    If mCount > 0 Then Exit Sub
    Reg "COMP_W_REG_D.COMP_W_REG_D", "Company Sale Register", True, "A1=Company Name  :"
    Reg "Sale_Receiving_Summary.Receiving_Summary", "Receiving Summary", True, "A3=Van;B3=Salesman"
    Reg "Collection_Outstanding_Report1.Collection_Outstanding_Report", "Collection Outstanding", True, "A5=PARTY_NAME;B5=INVOICE_NO;C5=INVOICE_DATE;D5=INVOICE_AMOUNT;E5=OUTS_DAYS;F5=COLLECTION_AMOUNT;G5=OUTSTANDING" & _
        "_AMOUNT"
    Reg "Van_Wise_Outstanding_Report.Van_Wise_Outstanding_Report", "Van Wise Outstanding", True, "A5=VAN_NAME;B5=TYP;C5=COMPANY_NAME;D5=PARTY_;E5=INVOICE_DATE;F5=INVOICE_NO;G5=INVOICE_AMNT"
    Reg "IRN_UPLOAD.IRN_UPLOAD", "IRN Upload Format", True, "A1=Sl. No;B1=Ack No;C1=Ack Date;D1=Doc No;E1=Doc Date;F1=Doc Type;G1=Inv Value.;H1=Recipient GSTIN;I1=Status;J" & _
        "1=IRN;K1=SignedQrCOde"
    Reg "E_Invoice_Validation.E_Invoice_Validation", "E-Invoice Validation", True, "A5=SUPPLY_TYPE_CODE"
    Reg "Flipkart_PO_formatter.Flipkart_PO_formatter", "Flipkart PO Format", True, "@1:PO ID;@1:PO Status;@1:PO Item Status;@1:PO Item ID;@1:Product ID;@1:supplier_mrp_amount;@1:Ordered Quantity" & _
        ";@1:Total Goods Amount;@1:supplier_app_amount"
    Reg "BlinkitPOFormatter.FormatBlinkitPO", "Blinkit PO Format", True, "@1:po_number;@1:facility_name;@1:manufacturer_name;@1:po_state;@1:upc;@1:order_date;@1:expiry_date;@1:remainin" & _
        "g_quantity;@1:total_amount"
    Reg "Marico_HideUnwantedColumns.Marico_HideUnwantedColumns", "Marico Hide Columns", True, "@1:Bill Date;@1:Bill Number;@1:Merged Order PO Reference;@1:Retailer Name;@1:Product Code;@1:Final Amount"
    Reg "DABUR.DABUR_E_INVOICE", "Dabur E-Invoice", True, "A1=BillWise ProductWise Sales Report;E3=Salesman Code;D3=PO Date"
    Reg "Dabur_SaleFromat_HA.Dabur_SaleFromat_HA", "Dabur Sale Format", True, "A3=Bill Number;B3=Bill Date;C3=PO Number;D3=PO Date;E3=Salesman Code;F3=Salesman Name;N3=MRP;O3=LSP Rate;P3=Se" & _
        "lling Rate;U3=Sales Value;Y3=Total Discount;Z3=Total Tax;AA3=Net Amount"
    Reg "Wisdom_Stkreport.Wisdom_Stkreport", "Wisdom Stock Report", False, "@5:Description;@5:MRP"
    Reg "DMS_REP.DMS_REP", "Mavic Report > Customer Wise Sale", True, "A1=Invoice Date;B1=Distributor Code;C1=Distributor Name;D1=Salesman Name;E1=Route Name;F1=Invoice Number;G1=di" & _
        "str customer code;H1=cmp customer code;I1=Retailer code;J1=Retailer Name;K1=Retailer GSTIN;L1=Retailer Created" & _
        " Date;M1=Retailer Channel;N1=Retailer Sub Channel;O1=Retailer Group;P1=Retailer Class;Q1=Source;R1=Delivery Da" & _
        "te;S1=HSN Code;T1=HSN Name;U1=Regional Brand Code;V1=Regional Brand Name;AW1=Status"
    Reg "DMS_HSN_SUMMARY.DMS_HSN_SUMMARY", "Mavic Report > HSN Summary", True, "A1=Invoice Date;B1=Distributor Code;C1=Distributor Name;D1=Salesman Name;E1=Route Name;F1=Invoice Number;G1=di" & _
        "str customer code;H1=cmp customer code;I1=Retailer code;J1=Retailer Name;K1=Retailer GSTIN;L1=Retailer Created" & _
        " Date;M1=Retailer Channel;N1=Retailer Sub Channel;O1=Retailer Group;P1=Retailer Class;Q1=Source;R1=Delivery Da" & _
        "te;S1=HSN Code;T1=HSN Name;U1=Regional Brand Code;V1=Regional Brand Name;W1=Sub Brand Code;X1=Sub Brand Name;Y" & _
        "1=Product Code;Z1=Product Name;AA1=Product Batch Code;AB1=Sell Rate;AC1=MRP;AD1=Invoice Quantity;AE1=Invoice Q" & _
        "uantity In KG;AF1=Offer Quantity;AG1=Offer Quantity In KG;AH1=Gross Amount;AI1=CGST %;AJ1=CGST;AK1=SGST/UGST %" & _
        ";AL1=SGST/UGST;AM1=IGST %;AN1=IGST;AO1=CESS %;AP1=CESS Value;AQ1=TCS %;AR1=TCS Amount;AS1=Scheme Discount;AT1=" & _
        "Distributor Discount;AU1=Cash Discount;AW1=Status;C2=MAX INTERNATIONAL"
    Reg "DMS_STK.DMS_STK", "Mavic Report > Stock Statement", True, "A1=Zone;B1=State;C1=District;D1=Taluk;E1=Market;F1=ZSM;G1=Cluster Manager;H1=ASM;I1=ASE;J1=SDO;K1=Distributor " & _
        "Code;L1=Distributor Name;M1=Distributor Br Code;N1=Distributor Br Name;O1=Godown Name;P1=HSN Code;Q1=Product C" & _
        "ode;R1=Product Description;S1=Batch;T1=Expiry Date;U1=MRP;V1=PurchPrice Without Tax;W1=Report UOM Code;X1=Sale" & _
        "able Stock;Y1=Report UOM Saleable Qty;AH1=Total ReportUOMQty"
    Reg "E_invoice_Format_DMS_MAX.E_invoice_Format_DMS_MAX", "Mavic Report > E-Invoice Format MAX", True, "A1=Invoice Date;B1=Distributor Code;C1=Distributor Name;D1=Salesman Name;E1=Route Name;F1=Invoice Number;G1=di" & _
        "str customer code;H1=cmp customer code;I1=Retailer code;J1=Retailer Name;K1=Retailer GSTIN;L1=Retailer Created" & _
        " Date;M1=Retailer Channel;N1=Retailer Sub Channel;O1=Retailer Group;P1=Retailer Class;Q1=Source;R1=Delivery Da" & _
        "te;S1=HSN Code;T1=HSN Name;U1=Regional Brand Code;V1=Regional Brand Name;W1=Sub Brand Code;X1=Sub Brand Name;Y" & _
        "1=Product Code;Z1=Product Name;AA1=Product Batch Code;AB1=Sell Rate;AC1=MRP;AD1=Invoice Quantity;AE1=Invoice Q" & _
        "uantity In KG;AF1=Offer Quantity;AG1=Offer Quantity In KG;AH1=Gross Amount;AI1=CGST %;AJ1=CGST;AK1=SGST/UGST %" & _
        ";AL1=SGST/UGST;AM1=IGST %;AN1=IGST;AO1=CESS %;AP1=CESS Value;AQ1=TCS %;AR1=TCS Amount;AS1=Scheme Discount;AT1=" & _
        "Distributor Discount;AU1=Cash Discount;AW1=Status;C2=MAX INTERNATIONAL"
    Reg "RIL_SRM_Excel_Filter_Data.RIL_SRM_Excel_Filter_Data", "SRM PO Excel", True, "A1=DC_CODE;B1=PURCH_ORDER_NUMBER;C1=PURCH_ORDER_DATE;D1=EXPECTED_DATE;E1=SUPPLIER_CODE;F1=SUPPLIER NAME;G1=PUR" & _
        "CH_ORDER_LINE;H1=EAN_NO;I1=PRODUCT_CODE;J1=PRODUCT_DESCRIPTION;K1=VENDOR_ITEM_CODE;L1=ORDERING_UNIT;M1=QTY_PER" & _
        "_PACK_OF_ORDERING_UOM;N1=NO_OF_PACKS_ORDERED;O1=TOTAL_QUANTITY"
    Reg "Print_Setup_RIL.Print_Setup", "Print Setup", False, "A1=DC_CODE;B1=Store_Name"
    Reg "RIL_PO_TO_EXCEL.RIL_PO_TO_EXCEL", "RIL PO to Excel", True, "B1=Article No. HSN/SAC Code"
    Reg "RIL_PO_TO_EXCEL.RIL_PO_TO_EXCEL", "RIL PO to Excel", True, "A1=PURCHASE ORDER*;B2=Article No. HSN/SAC Code"
    Reg "RIL_Process_For_Amendment.RIL_Process_For_Amendment", "Process Amendment", True, "A1=DC_CODE;B1=Store_Name;C1=PURCH_ORDER_NUMBER;D1=PURCH_ORDER_DATE;E1=EXPECTED_DATE;F1=PURCH_ORDER_LINE;G1=EAN" & _
        "_NO;H1=PRODUCT_CODE;I1=Company;J1=PRODUCT_DESCRIPTION;K1=MRP_PER_UNIT;L1=QTY_PER_PACK_OF_ORDERING_UOM;M1=NO_OF" & _
        "_PACKS_ORDERED"
    Reg "RIL_Insert_Brand_Name_ASN.RIL_Insert_Brand_Name_ASN", "Insert Brand (ASN)", True, "A1=Invoice No.;B1=RR PO No.;C1=Invoice Date;D1=Delivery Date"
    Reg "RIL_Appointment_Auto.RIL_Appointment", "RIL Appointment", True, "A1=DC_CODE;B1=PURCH_ORDER_NUMBER;C1=PURCH_ORDER_DATE;D1=EXPECTED_DATE;E1=SUPPLIER_CODE;F1=SUPPLIER_NAME;G1=PUR" & _
        "CH_ORDER_LINE;H1=EAN_NO;I1=PRODUCT_CODE;J1=PRODUCT_DESCRIPTION;K1=VENDOR_ITEM_CODE;L1=ORDERING_UNIT;M1=QTY_PER" & _
        "_PACK_OF_ORDERING_UOM;N1=NO_OF_PACKS_ORDERED;O1=TOTAL_QUANTITY"
End Sub