Attribute VB_Name = "Marico_HideUnwantedColumns"
Public Sub Marico_HideUnwantedColumns()
    Dim keep As Variant
    keep = Array( _
        "Bill Date", "Bill Number", "Merged Order PO Reference", "Status", _
        "Retailer Name", "SubChannelName", "Brand", "Product Code", _
        "Product Name", "Batch Code", "New MRP", "Sales Qty", "Offer Qty", _
        "Total Qty", "Gross Amt", "Sales Ex Tax", "Total Tax Amount", _
        "NetAmount", "Final Amount")

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim hdr As String, found As Long
    Application.ScreenUpdating = False

    For c = 1 To lastCol
        hdr = Trim$(CStr(ws.Cells(1, c).Value))
        If IsInList(hdr, keep) Then
            ws.Columns(c).Hidden = False
            found = found + 1
        Else
            ws.Columns(c).Hidden = True
        End If
    Next c

    Application.ScreenUpdating = True

    If found <> UBound(keep) + 1 Then
        MsgBox "Done, but only " & found & " of " & (UBound(keep) + 1) & _
               " expected columns were found. Check the header names in row 1.", vbExclamation
    Else
        MsgBox "Done. " & found & " columns kept, the rest hidden.", vbInformation
    End If
End Sub

' Optional: unhide everything again.
Public Sub ShowAllColumns()
    ActiveSheet.Columns.Hidden = False
End Sub

Private Function IsInList(ByVal s As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(s, arr(i), vbTextCompare) = 0 Then
            IsInList = True
            Exit Function
        End If
    Next i
End Function


