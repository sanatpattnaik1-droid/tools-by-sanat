Attribute VB_Name = "CombineHeaderAndClean"
Sub CombineHeaderAndClean()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastCol As Long
    Dim c7 As Long, c8 As Long, c9 As Long
    
    c7 = ws.Cells(7, ws.Columns.Count).End(xlToLeft).Column
    c8 = ws.Cells(8, ws.Columns.Count).End(xlToLeft).Column
    c9 = ws.Cells(9, ws.Columns.Count).End(xlToLeft).Column
    lastCol = Application.WorksheetFunction.Max(c7, c8, c9)
    
    Dim c As Long, r As Long
    Dim combinedText As String, v As String

    ' No fixed headings to check, so show what is about to be deleted and ask first.
    If lastCol = 0 Or Application.WorksheetFunction.CountA(ws.Rows("7:9")) = 0 Then
        MsgBox "Rows 7 to 9 of this sheet are empty, so there is no 3-line header to combine." & vbLf & _
               "Nothing was changed.", vbExclamation, "Combine Header + Clean"
        Exit Sub
    End If
    If MsgBox("This will delete rows 1 to 6 and join the header lines in rows 7, 8 and 9 into one row." & vbLf & _
              "It cannot be undone." & vbLf & vbLf & _
              "Sheet: [" & ws.Parent.Name & "] " & ws.Name & vbLf & _
              "Row 7 starts with: " & Left$(Trim$(CStr(ws.Cells(7, 1).Text)) & " " & Trim$(CStr(ws.Cells(7, 2).Text)), 60) & vbLf & vbLf & _
              "Is this the right file?", vbQuestion + vbYesNo + vbDefaultButton2, "Combine Header + Clean") <> vbYes Then Exit Sub

    ' 1. Insert a blank row right after row 9
    ws.Rows(10).Insert Shift:=xlDown
    
    ' 2. Build concatenated header text (row7 + row8 + row9) into new row 10
    For c = 1 To lastCol
        combinedText = ""
        For r = 7 To 9
            v = Trim(ws.Cells(r, c).Value)
            If v <> "" Then
                If combinedText <> "" Then combinedText = combinedText & " "
                combinedText = combinedText & v
            End If
        Next r
        ws.Cells(10, c).Value = combinedText   ' written directly as a value (no formula, so no separate paste-special needed)
    Next c
    
    ' 3. Delete the old 3-line header rows (7,8,9)
    ws.Rows("7:9").Delete Shift:=xlUp
    
    ' 4. Delete the top info rows (1 to 6)
    ws.Rows("1:6").Delete Shift:=xlUp
    
    MsgBox "Done! Header combined and cleaned up.", vbInformation

End Sub
