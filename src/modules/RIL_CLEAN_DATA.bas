Attribute VB_Name = "RIL_CLEAN_DATA"
Sub RIL_CLEAN_DATA1()
Dim answer As Integer
 Dim Cel As Range, rng As Range, uRng As Range, c As Long, r As Long
    Set rng = ActiveSheet.UsedRange
    Set uRng = rng.Offset(rng.Rows.Count).Resize(1, 1)
    For r = 2 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            Set Cel = rng.Cells(r, c)
            If Not Cel.EntireRow.Interior.ColorIndex = 19 Then
                Set uRng = Union(uRng, Cel)
                Exit For
            End If
        Next c
    Next r
    uRng.EntireRow.Delete
    UserForm1.Show
End Sub

    
    
