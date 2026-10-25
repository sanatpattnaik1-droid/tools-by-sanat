Attribute VB_Name = "Fill_Down"
Sub FillDownBlanks()
Dim c As Range
For Each c In Selection.Cells
   If c.Value = "" Then
    c.FillDown
    End If
Next c
End Sub
