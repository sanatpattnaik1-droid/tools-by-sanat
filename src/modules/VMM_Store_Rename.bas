Attribute VB_Name = "VMM_Store_Rename"
Sub VMM_Store_Rename()

Range("C1").Select
    ActiveCell.FormulaR1C1 = "Old Store Name"
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "New Store Name"
  Dim lastRow As Long
    lastRow = Range("B" & Rows.Count).End(xlUp).row
    For i = 2 To lastRow
        If Range("B" & i).Value = "HO01" Then
Range("C" & i).Value = "CUTTACK(7)"
End If
If Range("B" & i).Value = "HO03" Then
Range("C" & i).Value = "BERHAMPUR"
End If
If Range("B" & i).Value = "HO04" Then
Range("C" & i).Value = "JHARSUGUDA"
End If
If Range("B" & i).Value = "HO05" Then
Range("C" & i).Value = "BARIPADA"
End If
If Range("B" & i).Value = "HO06" Then
Range("C" & i).Value = "BHADRAK"
End If
If Range("B" & i).Value = "HO01" Then
Range("C" & i).Value = "CUTTACK"
End If
If Range("B" & i).Value = "HO07" Then
Range("C" & i).Value = "BHUBANESWAR-3-JD VIHAR"
End If
If Range("B" & i).Value = "HO08" Then
Range("C" & i).Value = "ROURKELA(12)"
End If
If Range("B" & i).Value = "HO09" Then
Range("C" & i).Value = "PURI(13)"
End If
If Range("B" & i).Value = "HO10" Then
Range("C" & i).Value = "BHUBNESWAR-4-PATRAPARA"
End If
If Range("B" & i).Value = "HO11" Then
Range("C" & i).Value = "BHUBANESWAR-5-RASULGARH"
End If
If Range("B" & i).Value = "HO12" Then
Range("C" & i).Value = "BALANGIR-SWOSTI TOWER"
End If
If Range("B" & i).Value = "HO13" Then
Range("C" & i).Value = "BARGARH"
End If
If Range("B" & i).Value = "HO14" Then
Range("C" & i).Value = "BALASORE"
End If
If Range("B" & i).Value = "HO15" Then
Range("C" & i).Value = "SAMBHALPUR"
End If
If Range("B" & i).Value = "HO19" Then
Range("C" & i).Value = "KENDRAPADA"
End If
If Range("B" & i).Value = "HO20" Then
Range("C" & i).Value = "CUTTACK-2 CDA"
End If
If Range("B" & i).Value = "HO17" Then
Range("C" & i).Value = "BHAWANIPATNA"
End If
If Range("B" & i).Value = "HO21" Then
Range("C" & i).Value = "KEONJHAR"
End If
If Range("B" & i).Value = "HO23" Then
Range("C" & i).Value = "ANUGUL"
End If
If Range("B" & i).Value = "HO22" Then
Range("C" & i).Value = "RAYAGADA"
End If
If Range("B" & i).Value = "HO27" Then
Range("C" & i).Value = "BHANJANAGAR"
End If
If Range("B" & i).Value = "HO28" Then
Range("C" & i).Value = "NOWRANGPUR"
End If
If Range("B" & i).Value = "HO25" Then
Range("C" & i).Value = "BHUBANESHWAR LEWIS ROAD(20)"
End If
If Range("B" & i).Value = "HO30" Then
Range("C" & i).Value = "JAJPUR ROAD"
End If
If Range("B" & i).Value = "HO24" Then
Range("C" & i).Value = "SEMILIGUDA"
End If
If Range("B" & i).Value = "HO32" Then
Range("C" & i).Value = "TITILAGARH"
End If
If Range("B" & i).Value = "HO29" Then
Range("C" & i).Value = "PARALEKHEMUNDI-HO29"
End If
If Range("B" & i).Value = "HO26" Then
Range("C" & i).Value = "SUNDERGARH-HO26"
End If
If Range("B" & i).Value = "HO31" Then
Range("C" & i).Value = "JAGATPUR-HO31(JAGATPUR-H)"
End If
If Range("B" & i).Value = "HO35" Then
Range("C" & i).Value = "JAGATSINGHPUR"




End If
    Next i
End Sub
