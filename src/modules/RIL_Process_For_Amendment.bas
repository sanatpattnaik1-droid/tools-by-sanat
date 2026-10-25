Attribute VB_Name = "RIL_Process_For_Amendment"
Sub RIL_Process_For_Amendment()
    answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "HA IRN Upload Format")
    If answer = vbYes Then
        If Range("A1").Value = "DC_CODE" And Range("B1").Value = "Store_Name" And Range("C1").Value = "PURCH_ORDER_NUMBER" And Range("D1").Value = "PURCH_ORDER_DATE" And Range("E1").Value = "EXPECTED_DATE" And Range("F1").Value = "PURCH_ORDER_LINE" And Range("G1").Value = "EAN_NO" And Range("H1").Value = "PRODUCT_CODE" And Range("I1").Value = "Company" And Range("J1").Value = "PRODUCT_DESCRIPTION" And Range("K1").Value = "MRP_PER_UNIT" And Range("L1").Value = "QTY_PER_PACK_OF_ORDERING_UOM" And Range("M1").Value = "NO_OF_PACKS_ORDERED" Then
            Application.ScreenUpdating = False
            
            If ActiveSheet.AutoFilterMode Then
                ActiveSheet.AutoFilterMode = False
            End If
            ActiveWorkbook.Save
            Range("A2:M300").Select
            Dim high As Range
            Dim c As Range
            Set high = Selection
            For Each c In high
                If c.Interior.Pattern <> xlNone Then
                    Run "RIL_CLEAN_DATA1"
                    Exit Sub
                End If
            Next c
            MsgBox ("PLEASE HIGHLIGHTED SKU(S) WHICH NEED TO AMENDMENT AND TRY AGAIN!")
            Range("A1").Select
            Application.ScreenUpdating = True
            Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        End If
        If answer = vbNo Then Exit Sub
    End If
    
    
End Sub
