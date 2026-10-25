Attribute VB_Name = "DMS_REP"
Sub DMS_REP()
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "DMS Customer Wise Sale Report")
   If answer = vbYes Then
If Range("A1").Value = "Invoice Date" And Range("B1").Value = "Distributor Code" And Range("C1").Value = "Distributor Name" And Range("D1").Value = "Salesman Name" And Range("E1").Value = "Route Name" And Range("F1").Value = "Invoice Number" And Range("G1").Value = "distr customer code" And Range("H1").Value = "cmp customer code" And Range("I1").Value = "Retailer code" And Range("J1").Value = "Retailer Name" And Range("K1").Value = "Retailer GSTIN" And Range("L1").Value = "Retailer Created Date" And Range("M1").Value = "Retailer Channel" And Range("N1").Value = "Retailer Sub Channel" And Range("O1").Value = "Retailer Group" And Range("P1").Value = "Retailer Class" And Range("Q1").Value = "Source" And Range("R1").Value = "Delivery Date" And Range("S1").Value = "HSN Code" And Range("T1").Value = "HSN Name" And Range("U1").Value = "Regional Brand Code" And Range("V1").Value = "Regional Brand Name" And Range("AW1").Value = "Status" Then
Application.ScreenUpdating = False
Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("O:O").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToLeft).Select
         Dim rng As Range
    Set rng = ActiveSheet.UsedRange
    For i = rng.Cells.Count To 1 Step -1
        If rng.Item(i).Value = "Cancel" Then
            rng.Item(i).EntireRow.Delete
             End If
             Next i
    Range("C1").CurrentRegion.Select
    Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        Range("A1").Select
        Application.ScreenUpdating = True
        MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
        
   Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        
    End If
          
Exit Sub
End If
End Sub
