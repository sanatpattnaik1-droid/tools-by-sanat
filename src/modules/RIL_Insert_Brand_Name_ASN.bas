Attribute VB_Name = "RIL_Insert_Brand_Name_ASN"
Sub RIL_Insert_Brand_Name_ASN()
If Range("A1").Value = "Invoice No." And Range("B1").Value = "RR PO No." And Range("C1").Value = "Invoice Date" And Range("D1").Value = "Delivery Date" Then
            Application.ScreenUpdating = False

    Columns("W:W").Select
    Selection.Cut
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "Company"
    Dim LstRw As Long, x
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 9)) = "CHANDRIKA" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "GIFFY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "HYGIENIX" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 2)) = "JO" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "MAXKLEEN" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "SAFEWASH" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "SANTOOR" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "SOFTOUCH" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "WIPRO" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "YARDLEY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "YRDLY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 2)) = "WP" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "YARDLY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "HYGNIX" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "YRDLEY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "DOY" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 9)) = "GLUCOVITA" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 10)) = "ENCHANTEUR" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "MXKLN" Then
                    Cells(x, "K") = "WIPRO"
                End If
                            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "GIFY" Then
                    Cells(x, "K") = "WIPRO"
                End If
                
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "SNTR" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "ZELSSY" Then
                    Cells(x, "K") = "ZELSSY"
                End If
            Next
            
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "VICCO" Then
                    Cells(x, "K") = "VICCO"
                End If
            Next
            
            
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "5050" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "BRIT" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "BRITANIA" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 9)) = "BRITANNIA" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "BRITNIA" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "LITTLE" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "MILK" Then
                    Cells(x, "K") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "DABUR" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "DBR" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "ODONIL" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "REAL" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "VATIKA" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "FEM" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "MESWAK" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "ODNIL" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "ODOM" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "ODOMOS" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "OXY" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "MESWAK" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "ODNIL" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "ODOM" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "ODOMOS" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "OXY" Then
                    Cells(x, "K") = "DABUR"
                End If
            Next
            
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "FERRERO" Then
                    Cells(x, "K") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "KINDER" Then
                    Cells(x, "K") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "NUTELLA" Then
                    Cells(x, "K") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "MDH" Then
                    Cells(x, "K") = "MDH"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 2)) = "PR" Then
                    Cells(x, "K") = "PRIYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "PRIYA" Then
                    Cells(x, "K") = "PRIYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "PEDIGREE" Then
                    Cells(x, "K") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "WHISKAS" Then
                    Cells(x, "K") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 3)) = "PED" Then
                    Cells(x, "K") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "KEYA" Then
                    Cells(x, "K") = "KEYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "SFWSH" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "OCEAN" Then
                    Cells(x, "K") = "OCEAN"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "MAYORA" Then
                    Cells(x, "K") = "MALKIST"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "MALKIST" Then
                    Cells(x, "K") = "MALKIST"
                End If
            Next
            LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 5)) = "KHADI" Then
                    Cells(x, "K") = "KHADI NATURAL"
                End If
            Next
             LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 8)) = "FLORALIS" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
             LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "YL FLOR" Then
                    Cells(x, "K") = "WIPRO"
                End If
            Next
             LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 4)) = "GRDN" Then
                    Cells(x, "K") = "WIPRO"
                End If

            Next
             LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 7)) = "SUNPURE" Then
                    Cells(x, "K") = "SUNPURE"
                End If
            Next
             LstRw = Cells(Rows.Count, "L").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "L"), 6)) = "KOPIKO" Then
                    Cells(x, "K") = "MALKIST"
                End If
            Next
               LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "FLOOR" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "YDLN" Then
                    Cells(x, "J") = "WIPRO"
                End If
             Next
                 
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "DAVIDOFF" Then
                    Cells(x, "J") = "UDI"
                End If
             Next
                 
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "TWINING" Then
                    Cells(x, "J") = "UDI"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "HARIBOL" Then
                    Cells(x, "J") = "HARIBOL"
                End If
            Next
                 
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "SANTOR" Then
                    Cells(x, "J") = "Wipro"
                End If
            Next
        

    Columns("M:M").Select
    Range("M2").Activate
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.End(xlUp).Select
    ActiveCell.FormulaR1C1 = "MRP_PER_UNIT"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "QTY_PER_PACK_OF_ORDERING_UOM"
    Columns("P:V").Select
    Selection.EntireColumn.Hidden = True
    Range("Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("W2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]/RC[-9]"
    Range("W2:W150").Select
    Selection.FillDown
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Application.ScreenUpdating = True
    MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
            Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        End If
        Exit Sub
End Sub


