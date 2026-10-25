Attribute VB_Name = "DMS_STK"
Sub DMS_STK()
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "DMS Stock")
   If answer = vbYes Then
If Range("A1").Value = "Zone" And Range("B1").Value = "State" And Range("C1").Value = "District" And Range("D1").Value = "Taluk" And Range("E1").Value = "Market" And Range("F1").Value = "ZSM" And Range("G1").Value = "Cluster Manager" And Range("H1").Value = "ASM" And Range("I1").Value = "ASE" And Range("J1").Value = "SDO" And Range("K1").Value = "Distributor Code" And Range("L1").Value = "Distributor Name" And Range("M1").Value = "Distributor Br Code" And Range("N1").Value = "Distributor Br Name" And Range("O1").Value = "Godown Name" And Range("P1").Value = "HSN Code" And Range("Q1").Value = "Product Code" And Range("R1").Value = "Product Description" And Range("S1").Value = "Batch" And Range("T1").Value = "Expiry Date" And Range("U1").Value = "MRP" And Range("V1").Value = "PurchPrice Without Tax" And Range("W1").Value = "Report UOM Code" And Range("X1").Value = "Saleable Stock" And Range("Y1").Value = "Report UOM Saleable Qty" And Range("AH1").Value = "Total ReportUOMQty" Then
    Application.ScreenUpdating = False
    Columns("A:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:P").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Total Qty In KG"
    Range("H2").Select
Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String

  SrcData = ActiveSheet.name & "!" & Columns("A:H").Address(ReferenceStyle:=xlR1C1)

  Set sht = Sheets.Add
  StartPvt = sht.name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)
     Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)
    Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")
With ActiveSheet.PivotTables("PivotTable1").PivotFields("MRP")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Product Description")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Qty"), "Sum of Total Qty", xlSum
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Total Qty In KG"), "Sum of Total Qty In KG", xlSum
        Application.ScreenUpdating = True
        MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
 
       Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        
    End If
          
Exit Sub
End If
End Sub

