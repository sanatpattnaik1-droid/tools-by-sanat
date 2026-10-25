Attribute VB_Name = "ConvertTextToNumber"
Sub ConvertTextToNumber()
 
    With Range("D2:F100")
     .NumberFormat = "General"
     .Value = .Value
    End With

End Sub
