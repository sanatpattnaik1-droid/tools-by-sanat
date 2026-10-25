Attribute VB_Name = "RandomDistributeItems"
Sub DistributeItemsWithLineItemLimit()
    Dim ws As Worksheet, outWs As Worksheet
    Dim lastRow As Long, row As Long, i As Long, p As Long
    Dim totalQty As Long, numPeople As Long: numPeople = 20
    Dim activePeople As Long, totalAssigned As Long
    Dim rndVals() As Double, randomParts() As Long
    Dim sumRnd As Double, indexList As Collection, idx As Long
    Dim maxLineItemsPerPerson As Long: maxLineItemsPerPerson = 20
    Dim maxPerPerson As Long: maxPerPerson = 20
    Dim personLineCount() As Long
   
    Randomize ' Initialize random number generator
   
    Set ws = ThisWorkbook.Sheets("Sheet1")
   
    ' Prepare output sheet
    On Error Resume Next
    Set outWs = ThisWorkbook.Sheets("DistributedOutput")
    If outWs Is Nothing Then
        Set outWs = ThisWorkbook.Sheets.Add(After:=ws)
        outWs.name = "DistributedOutput"
    Else
        outWs.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers
    outWs.Cells(1, 1).Value = "Item"
    For i = 1 To numPeople
        outWs.Cells(1, i + 1).Value = "Person " & i
    Next i

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
   
    ' Initialize line item count per person
    ReDim personLineCount(1 To numPeople)
    For p = 1 To numPeople
        personLineCount(p) = 0
    Next p

    For row = 2 To lastRow
        outWs.Cells(row, 1).Value = ws.Cells(row, 1).Value
        totalQty = ws.Cells(row, 2).Value
       
        ReDim rndVals(1 To numPeople)
        ReDim randomParts(1 To numPeople)
       
        For i = 1 To numPeople
            randomParts(i) = 0
        Next i
       
        ' Determine how many people to distribute to
        Set indexList = New Collection
       
        If totalQty > 25 Then
            activePeople = WorksheetFunction.RandBetween(6, 12)
        ElseIf totalQty < 10 Then
            activePeople = WorksheetFunction.RandBetween(2, 3)
        Else
            activePeople = numPeople
        End If
       
        ' Select random people with available line item slots
        Dim attemptCount As Long: attemptCount = 0
        Do While indexList.Count < activePeople And attemptCount < 1000
            idx = WorksheetFunction.RandBetween(1, numPeople)
            If personLineCount(idx) < maxLineItemsPerPerson Then
                On Error Resume Next
                indexList.Add idx, CStr(idx)
                On Error GoTo 0
            End If
            attemptCount = attemptCount + 1
            If attemptCount >= 1000 Then Exit Do ' break to avoid infinite loop
        Loop
       
        ' Adjust activePeople if less assigned due to line item limits
        activePeople = indexList.Count
       
        If activePeople = 0 Then
            MsgBox "No eligible people found for item '" & ws.Cells(row, 1).Value & "'. Please adjust max line item limit.", vbCritical
            Exit Sub
        End If
       
        ' Increase line item count for assigned people
        For i = 1 To activePeople
            idx = indexList.Item(i)
            personLineCount(idx) = personLineCount(idx) + 1
        Next i
       
        ' Generate random weights
        sumRnd = 0
        For i = 1 To activePeople
            rndVals(i) = Rnd
            sumRnd = sumRnd + rndVals(i)
        Next i
       
        ' Distribute quantities proportionally with maxPerPerson cap
        totalAssigned = 0
        For i = 1 To activePeople
            idx = indexList.Item(i)
            Dim portion As Long
            portion = Application.WorksheetFunction.RoundDown(rndVals(i) / sumRnd * totalQty, 0)
            If portion > maxPerPerson Then portion = maxPerPerson
            randomParts(idx) = portion
            totalAssigned = totalAssigned + portion
        Next i
       
        ' Distribute leftover quantities one by one safely
        Dim safetyCounter As Long: safetyCounter = 0
        Do While totalAssigned < totalQty And safetyCounter < 10000
            Dim randPersonIdx As Long
            randPersonIdx = WorksheetFunction.RandBetween(1, activePeople)
            idx = indexList.Item(randPersonIdx)
            If randomParts(idx) < maxPerPerson Then
                randomParts(idx) = randomParts(idx) + 1
                totalAssigned = totalAssigned + 1
            End If
            safetyCounter = safetyCounter + 1
        Loop
       
        If totalAssigned < totalQty Then
            MsgBox "Could not distribute all quantity for item '" & ws.Cells(row, 1).Value & "'. Increase maxPerPerson or maxLineItemsPerPerson.", vbExclamation
        End If
       
        ' Output results
        For i = 1 To numPeople
            outWs.Cells(row, i + 1).Value = randomParts(i)
        Next i
    Next row
   
    MsgBox "Distribution complete!", vbInformation
End Sub

