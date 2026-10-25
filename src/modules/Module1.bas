Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Range("C6").Select
    Columns("C:C").ColumnWidth = 11.14
    Columns("D:D").ColumnWidth = 10.29
    Columns("E:E").ColumnWidth = 12.14
    Columns("F:F").ColumnWidth = 4.71
    Columns("H:I").Select
    Range("I1").Activate
    Selection.EntireColumn.Hidden = True
    Columns("K:K").ColumnWidth = 6
    Columns("L:L").ColumnWidth = 3.43
    Columns("M:M").ColumnWidth = 4.29
    Columns("N:N").ColumnWidth = 4
    Columns("O:O").ColumnWidth = 5
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").ColumnWidth = 6
End Sub
