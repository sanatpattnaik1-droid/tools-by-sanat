Attribute VB_Name = "RIL_SRM_Excel_Filter_Data"
Sub RIL_SRM_Excel_Filter_Data()
    Dim answer As Integer
    answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "HA IRN Upload Format")
    If answer = vbYes Then
        If Range("A1").Value = "DC_CODE" And Range("B1").Value = "PURCH_ORDER_NUMBER" And Range("C1").Value = "PURCH_ORDER_DATE" And Range("D1").Value = "EXPECTED_DATE" And Range("E1").Value = "SUPPLIER_CODE" And Range("F1").Value = "SUPPLIER NAME" And Range("G1").Value = "PURCH_ORDER_LINE" And Range("H1").Value = "EAN_NO" And Range("I1").Value = "PRODUCT_CODE" And Range("J1").Value = "PRODUCT_DESCRIPTION" And Range("K1").Value = "VENDOR_ITEM_CODE" And Range("L1").Value = "ORDERING_UNIT" And Range("M1").Value = "QTY_PER_PACK_OF_ORDERING_UOM" And Range("N1").Value = "NO_OF_PACKS_ORDERED" And Range("O1").Value = "TOTAL_QUANTITY" Then
            Application.ScreenUpdating = False
            If ActiveSheet.AutoFilterMode Then
                ActiveSheet.AutoFilterMode = False
            End If
            Range("Y1").Select
            ActiveCell.FormulaR1C1 = _
            "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
            Range("Y1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Range("Y1").Select
            Selection.NumberFormat = ";;;"
            Columns("J:J").Select
            Selection.Insert Shift:=xlToRight
            Range("J1").Select
            ActiveCell.FormulaR1C1 = "Company"
            Dim LstRw As Long, x
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 9)) = "CHANDRIKA" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "GIFFY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "HYGIENIX" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 2)) = "JO" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "MAXKLEEN" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "SAFEWASH" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "SANTOOR" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "SOFTOUCH" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "WIPRO" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "YARDLEY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "YRDLY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 2)) = "WP" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "YARDLY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "HYGNIX" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "YRDLEY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "DOY" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 9)) = "GLUCOVITA" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 10)) = "ENCHANTEUR" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "MXKLN" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "SFWSH" Then
                    Cells(x, "J") = "WIPRO"
                End If
                
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "GIFY" Then
                    Cells(x, "J") = "WIPRO"
                End If
                
                                                Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "SNTR" Then
                    Cells(x, "J") = "WIPRO"
                End If
                
    
                
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "ZELSSY" Then
                    Cells(x, "J") = "ZELSSY"
                End If
            Next
            
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "VICCO" Then
                    Cells(x, "J") = "VICCO"
                End If
            Next
            
            
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "5050" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "BRIT" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "BRITANIA" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 9)) = "BRITANNIA" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "BRITNIA" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "LITTLE" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "MILK" Then
                    Cells(x, "J") = "BRITANNIA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "DABUR" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "DBR" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "ODONIL" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "REAL" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "VATIKA" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "FEM" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "MESWAK" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "ODNIL" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "ODOM" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "ODOMOS" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "OXY" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "MESWAK" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "ODNIL" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "ODOM" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "ODOMOS" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "OXY" Then
                    Cells(x, "J") = "DABUR"
                End If
            Next
            
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "FERRERO" Then
                    Cells(x, "J") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "KINDER" Then
                    Cells(x, "J") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "NUTELLA" Then
                    Cells(x, "J") = "FERRERO"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "MDH" Then
                    Cells(x, "J") = "MDH"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 2)) = "PR" Then
                    Cells(x, "J") = "PRIYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "PRIYA" Then
                    Cells(x, "J") = "PRIYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "PEDIGREE" Then
                    Cells(x, "J") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "WHISKAS" Then
                    Cells(x, "J") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "PED" Then
                    Cells(x, "J") = "PEDIGREE"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "KEYA" Then
                    Cells(x, "J") = "KEYA"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "OCEAN" Then
                    Cells(x, "J") = "OCEAN"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "MAYORA" Then
                    Cells(x, "J") = "MALKIST"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "MALKIST" Then
                    Cells(x, "J") = "MALKIST"
                End If
            Next
                LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "KHADI" Then
                    Cells(x, "J") = "KHADI NATURAL"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 8)) = "FLORALIS" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "YL FLOR" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 4)) = "GRDN" Then
                    Cells(x, "J") = "WIPRO"
                End If

            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "SUNPURE" Then
                    Cells(x, "J") = "SUNPURE"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "KOPIKO" Then
                    Cells(x, "J") = "MALKIST"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "LOACKER" Then
                    Cells(x, "J") = "UDI"
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
                If UCase(Left(Cells(x, "K"), 6)) = "BACTER" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            

             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "TABASCO" Then
                    Cells(x, "J") = "UDI"
                End If
            Next

             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "BARILLA" Then
                    Cells(x, "J") = "UDI"
                End If
            Next
     
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 7)) = "Barilla" Then
                    Cells(x, "J") = "UDI"
                End If
            Next

     
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 5)) = "LOTUS" Then
                    Cells(x, "J") = "UDI"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "SANTOR" Then
                    Cells(x, "J") = "Wipro"
                End If
            Next
             LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 3)) = "OCN" Then
                    Cells(x, "J") = "OCEAN"
                End If
            Next
            LstRw = Cells(Rows.Count, "K").End(xlUp).row
            For x = 1 To LstRw
                If UCase(Left(Cells(x, "K"), 6)) = "O'CEAN" Then
                    Cells(x, "J") = "OCEAN"
                End If
            Next
            
        
        
            Range("J1").Select
            ActiveCell.FormulaR1C1 = "Company"
            
            Rows("1:10000").Copy
            Workbooks.Add
            ActiveSheet.Paste
            Columns("E:E").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Columns("J:J").Select
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Columns("M:M").Select
            Selection.Delete Shift:=xlToLeft
            Columns("N:N").Select
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Columns("O:O").Select
            Selection.Delete Shift:=xlToLeft
            Columns("B:B").Select
            Selection.Insert Shift:=xlToRight
            Range("B1").Select
            ActiveCell.FormulaR1C1 = "Store_Name"
            Columns("N:N").Select
            Selection.Cut
            Columns("K:K").Select
            Selection.Insert Shift:=xlToRight
            Dim lastRow As Long
            lastRow = Range("E" & Rows.Count).End(xlUp).row
            For i = 2 To lastRow
                If Range("A" & i).Value = "FR66" Then
                    Range("B" & i).Value = "CUTTACK-FR66"
                End If
                If Range("A" & i).Value = "FR67" Then
                    Range("B" & i).Value = "PATRAPADA-FR67"
                End If
                If Range("A" & i).Value = "FR63" Then
                    Range("B" & i).Value = "ESPLANADE MALL-FR63"
                End If
                If Range("A" & i).Value = "FR64" Then
                    Range("B" & i).Value = "KHARVEL NAGAR-FR64"
                End If
                If Range("A" & i).Value = "FR62" Then
                    Range("B" & i).Value = "NAYAPALLI-FR62"
                End If
                If Range("A" & i).Value = "FR65" Then
                    Range("B" & i).Value = "PATIA-FR65"
                End If
                If Range("A" & i).Value = "FR68" Then
                    Range("B" & i).Value = "SYMPHONY MALL-FR68"
                End If
                If Range("A" & i).Value = "FRIC" Then
                    Range("B" & i).Value = "JAJPUR-FRIC"
                End If
                If Len(PartyField(Range("A" & i).Value, "name")) > 0 Then
                    Range("B" & i).Value = PartyField(Range("A" & i).Value, "name")
                End If
                If Range("A" & i).Value = "FRFM" Then
                    Range("B" & i).Value = "JHARSUGUDA-FRFM"
                End If
                If Range("A" & i).Value = "FRGG" Then
                    Range("B" & i).Value = "KEONJHAR-FRGG"
                End If
                If Range("A" & i).Value = "FR69" Then
                    Range("B" & i).Value = "PURI-FR69"
                End If
                If Range("A" & i).Value = "FRFN" Then
                    Range("B" & i).Value = "ROURKELA-FRFN"
                End If
                If Range("A" & i).Value = "SLMR" Then
                    Range("B" & i).Value = "MUKUNDPRASAD -DC"
                End If
                If Range("A" & i).Value = "R801" Then
                    Range("B" & i).Value = "RELIANCE-DC"
                End If
            Next i
            Dim Filename As String
            path = Range("S1")
            Filename = "Reliance_"
            dt = Format(CStr(Now), "dd_mmmm_yyyyhhmmss")
            ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
            Range("S:S").Select
            Selection.Delete Shift:=xlToLeft
            ActiveWorkbook.Save
            Cells.Select
            Cells.EntireColumn.AutoFit
            Range("A1").Select
            Application.ScreenUpdating = True
            MsgBox "Data Proceed Completed Successfully.", vbInformation, "Status"
            MsgBox "HIGHLIGHT THE SKU(S) NEED TO BE AMENDMENT", vbInformation, "Status"
            MsgBox "File Exported To.. " & path & Filename & dt & ""
            Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        End If
        Exit Sub
    End If
End Sub
