Attribute VB_Name = "Hands_on_Trade_PO_Clean"
Sub CleanPOData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String
    Dim dateVal As Date
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ============================================================
    ' STEP 1: Convert PO Number in Column A to Text format
    ' ============================================================
    For i = 2 To lastRow
        cellVal = CStr(ws.Cells(i, "A").Value)
        ws.Cells(i, "A").NumberFormat = "@"
        ws.Cells(i, "A").Value = cellVal
    Next i
    
    ' ============================================================
    ' STEP 2: Clean dates in Column F and G (format: 2026-02-26 05:50:19+00:00)
    ' ============================================================
    Dim colFG As Variant
    colFG = Array("F", "G")
    Dim col As Variant
    
    For Each col In colFG
        For i = 2 To lastRow
            cellVal = Trim(CStr(ws.Cells(i, col).Value))
            If cellVal <> "" Then
                ' Remove timezone offset like +00:00
                If InStr(cellVal, "+") > 0 Then
                    cellVal = Left(cellVal, InStr(cellVal, "+") - 1)
                End If
                cellVal = Trim(cellVal)
                ' Keep only date part (before space)
                If InStr(cellVal, " ") > 0 Then
                    cellVal = Left(cellVal, InStr(cellVal, " ") - 1)
                End If
                ' cellVal is now like: 2026-02-26
                On Error Resume Next
                dateVal = CDate(cellVal)
                If Err.Number = 0 Then
                    ws.Cells(i, col).NumberFormat = "DD/MM/YYYY"
                    ws.Cells(i, col).Value = dateVal
                End If
                On Error GoTo 0
            End If
        Next i
    Next col
    
    ' ============================================================
    ' STEP 3: Clean dates in Column H (format: 2026-03-08T18:29:59Z)
    ' ============================================================
    For i = 2 To lastRow
        cellVal = Trim(CStr(ws.Cells(i, "H").Value))
        If cellVal <> "" Then
            ' Remove the Z at the end
            If Right(cellVal, 1) = "Z" Then
                cellVal = Left(cellVal, Len(cellVal) - 1)
            End If
            ' Split on "T" to separate date and time
            If InStr(cellVal, "T") > 0 Then
                cellVal = Left(cellVal, InStr(cellVal, "T") - 1)
            End If
            ' cellVal is now like: 2026-03-08
            ' Parse YYYY-MM-DD manually
            Dim dateParts() As String
            dateParts = Split(cellVal, "-")
            If UBound(dateParts) = 2 Then
                Dim yyyy As Integer, mm As Integer, dd As Integer
                yyyy = CInt(dateParts(0))
                mm = CInt(dateParts(1))
                dd = CInt(dateParts(2))
                On Error Resume Next
                dateVal = DateSerial(yyyy, mm, dd)
                If Err.Number = 0 Then
                    ws.Cells(i, "H").NumberFormat = "DD/MM/YYYY"
                    ws.Cells(i, "H").Value = dateVal
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' ============================================================
    ' STEP 4: Move MRP column (Y) to Column N
    ' ============================================================
    ' Insert new blank column at N (shifts existing N onwards to right)
    ws.Columns("N").Insert Shift:=xlToRight
    
    ' After insert, old column Y is now Z — copy it to N
    Dim newLastRow As Long
    newLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ws.Range("Z1:Z" & newLastRow).Copy Destination:=ws.Range("N1")
    
    ' Delete original MRP column (Z)
    ws.Columns("Z").Delete
    
    ' Format Column N as number with 2 decimals
    ws.Range("N2:N" & newLastRow).NumberFormat = "0.00"
    
    MsgBox "Data cleaning complete!" & vbCrLf & _
           "1. PO Numbers converted to text in Column A" & vbCrLf & _
           "2. Dates cleaned in Columns F, G (DD/MM/YYYY)" & vbCrLf & _
           "3. Dates cleaned in Column H - ISO 8601 format handled (DD/MM/YYYY)" & vbCrLf & _
           "4. MRP column moved from Y to N", vbInformation
End Sub

