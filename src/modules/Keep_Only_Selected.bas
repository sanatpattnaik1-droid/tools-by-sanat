Attribute VB_Name = "Keep_Only_Selected"
Sub Keep_Only_Selected()
Dim answer As Integer
answer = MsgBox("This Process Will Delete the Row Which is Not Highlighted Do You Want to Continue?", vbQuestion + vbYesNo + vbDefaultButton1, "HA IRN Upload Format")
   If answer = vbYes Then
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
    Else
    End If
End Sub
