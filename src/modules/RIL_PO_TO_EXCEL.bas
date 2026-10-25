Attribute VB_Name = "RIL_PO_TO_EXCEL"
Sub RIL_PO_TO_EXCEL()
If Range("A1:A1") Like "PURCHASE ORDER*" Then
Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Run "RIL_PO_TO_EXCEL_PROCESS1"
    Else
    Run "RIL_PO_TO_EXCEL_PROCESS1"
    End If
End Sub

