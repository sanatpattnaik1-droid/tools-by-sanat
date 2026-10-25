Attribute VB_Name = "RIL_Appointment_Auto"
Sub RIL_Appointment()
    Dim answer As Integer
    answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "HA IRN Upload Format")
    If answer = vbYes Then
        If Range("A1").Value = "DC_CODE" And Range("B1").Value = "PURCH_ORDER_NUMBER" And Range("C1").Value = "PURCH_ORDER_DATE" And Range("D1").Value = "EXPECTED_DATE" And Range("E1").Value = "SUPPLIER_CODE" And Range("F1").Value = "SUPPLIER_NAME" And Range("G1").Value = "PURCH_ORDER_LINE" And Range("H1").Value = "EAN_NO" And Range("I1").Value = "PRODUCT_CODE" And Range("J1").Value = "PRODUCT_DESCRIPTION" And Range("K1").Value = "VENDOR_ITEM_CODE" And Range("L1").Value = "ORDERING_UNIT" And Range("M1").Value = "QTY_PER_PACK_OF_ORDERING_UOM" And Range("N1").Value = "NO_OF_PACKS_ORDERED" And Range("O1").Value = "TOTAL_QUANTITY" Then
            Application.ScreenUpdating = False
            If ActiveSheet.AutoFilterMode Then
                ActiveSheet.AutoFilterMode = False
            End If
            Range("X1").Select
            ActiveCell.FormulaR1C1 = _
            "=LEFT(CELL(""filename"",RC[1]),FIND(""["",CELL(""filename"",RC[1]))-1)"
            Range("X1").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Range("X1").Select
            Selection.NumberFormat = ";;;"
            Application.ScreenUpdating = False
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
                If UCase(Left(Cells(x, "K"), 5)) = "MXKLN" Then
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
                If UCase(Left(Cells(x, "K"), 5)) = "SFWSH" Then
                    Cells(x, "J") = "Wipro"
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
                    Cells(x, "J") = "Wipro"
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
                If UCase(Left(Cells(x, "K"), 8)) = "DAVIDOFF" Then
                    Cells(x, "J") = "DAVIDOFF"
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
                If UCase(Left(Cells(x, "K"), 7)) = "HARIBOL" Then
                    Cells(x, "J") = "HARIBOL"
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
                If UCase(Left(Cells(x, "K"), 6)) = "BACTER" Then
                    Cells(x, "J") = "WIPRO"
                End If
            Next
            
            
            
            Range("J1").Select
            ActiveCell.FormulaR1C1 = "Company"
            
            
            Range("A1:Y4000").Copy
            
            
            Workbooks.Add
            ActiveSheet.Paste
            Application.CutCopyMode = False
            Application.CutCopyMode = False
            Sheets.Add
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Sheet1!R1C1:R4000C22", Version:=6).CreatePivotTable TableDestination:= _
            "Sheet2!R3C1", TableName:="PivotTable1", DefaultVersion:=6
            Sheets("Sheet2").Select
            Cells(3, 1).Select
            With ActiveSheet.PivotTables("PivotTable1")
                .ColumnGrand = True
                .HasAutoFormat = True
                .DisplayErrorString = False
                .DisplayNullString = True
                .EnableDrilldown = True
                .ErrorString = ""
                .MergeLabels = False
                .NullString = ""
                .PageFieldOrder = 2
                .PageFieldWrapCount = 0
                .PreserveFormatting = True
                .RowGrand = True
                .SaveData = True
                .PrintTitles = False
                .RepeatItemsOnEachPrintedPage = True
                .TotalsAnnotation = False
                .CompactRowIndent = 1
                .InGridDropZones = False
                .DisplayFieldCaptions = True
                .DisplayMemberPropertyTooltips = False
                .DisplayContextTooltips = True
                .ShowDrillIndicators = True
                .PrintDrillIndicators = False
                .AllowMultipleFilters = False
                .SortUsingCustomLists = True
                .FieldListSortAscending = False
                .ShowValuesRow = False
                .CalculatedMembersInFilters = False
                .RowAxisLayout xlCompactRow
            End With
            With ActiveSheet.PivotTables("PivotTable1").PivotCache
                .RefreshOnFileOpen = False
                .MissingItemsLimit = xlMissingItemsDefault
            End With
            ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
            With ActiveSheet.PivotTables("PivotTable1").PivotFields("DC_CODE")
                .Orientation = xlRowField
                .Position = 1
            End With
            With ActiveSheet.PivotTables("PivotTable1").PivotFields("PURCH_ORDER_NUMBER")
                .Orientation = xlRowField
                .Position = 2
            End With
            With ActiveSheet.PivotTables("PivotTable1").PivotFields("PURCH_ORDER_DATE")
                .Orientation = xlRowField
                .Position = 3
            End With
            With ActiveSheet.PivotTables("PivotTable1").PivotFields("EXPECTED_DATE")
                .Orientation = xlRowField
                .Position = 4
            End With
            ActiveSheet.PivotTables("PivotTable1").PivotFields("DC_CODE").Subtotals = Array _
            (False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("PURCH_ORDER_NUMBER"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("PURCH_ORDER_DATE"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("EXPECTED_DATE").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields("SUPPLIER_CODE").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields("SUPPLIER_NAME").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields("PURCH_ORDER_LINE"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("EAN_NO").Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("PRODUCT_CODE").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("Company").Subtotals = Array _
            (False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("PRODUCT_DESCRIPTION"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("VENDOR_ITEM_CODE"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("ORDERING_UNIT").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields( _
            "QTY_PER_PACK_OF_ORDERING_UOM").Subtotals = Array(False, False, False, False, False _
            , False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("NO_OF_PACKS_ORDERED"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("TOTAL_QUANTITY").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields("CLOSING_STOCK").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )
            ActiveSheet.PivotTables("PivotTable1").PivotFields("MRP_PER_UNIT").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields( _
            "LANDING_COST_INCL_TAX_PER_UNIT").Subtotals = Array(False, False, False, False, _
            False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("SALES_TAX").Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields( _
            "COST_PRICE_INCL_TAX_PER_PO_OU").Subtotals = Array(False, False, False, False, _
            False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").PivotFields("TAX_PER").Subtotals = Array _
            (False, False, False, False, False, False, False, False, False, False, False, False)
            ActiveSheet.PivotTables("PivotTable1").RowAxisLayout xlTabularRow
            ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
            "PivotTable1").PivotFields("NO_OF_PACKS_ORDERED"), "Sum of NO_OF_PACKS_ORDERED" _
            , xlSum
            ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
            "PivotTable1").PivotFields("PURCH_ORDER_LINE"), "Count of PURCH_ORDER_LINE", _
            xlCount
            
            
            Cells.Select
            Selection.Copy
            
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Columns("F:F").Select
            Selection.Cut
            Columns("E:E").Select
            Selection.Insert Shift:=xlToRight
            Columns("C:C").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Application.CutCopyMode = False
            Selection.EntireRow.Delete
            Columns("E:E").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Application.CutCopyMode = False
            Selection.EntireRow.Delete
            Columns("F:F").Select
            Selection.Cut
            Columns("E:E").Select
            Selection.Insert Shift:=xlToRight
            Columns("F:F").Select
            Selection.Insert Shift:=xlToRight
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Selection.End(xlToLeft).Select
            Range("A2:G4000").Select
            Selection.Cut
            Range("B3").Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            
            Range("A1:A2").Select
            Application.CutCopyMode = False
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Range("A1:A2").Select
            Selection.AutoFill Destination:=Range("A1:E2"), Type:=xlFillDefault
            Range("A1:E2").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range("A1:A2").Select
            ActiveCell.FormulaR1C1 = "Sr. No."
            Range("B1:B2").Select
            ActiveCell.FormulaR1C1 = "Site"
            Range("C1:C2").Select
            ActiveCell.FormulaR1C1 = "PO No"
            Range("D1:D2").Select
            ActiveCell.FormulaR1C1 = "PO Creation Date"
            Range("E1:E2").Select
            ActiveCell.FormulaR1C1 = "PO Delivery Date"
            Range("F1").Select
            ActiveCell.FormulaR1C1 = "PO Qty "
            Range("F2").Select
            ActiveCell.FormulaR1C1 = "(in Cases/Bags)"
            Range("F1:F2").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range("G1:G2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Selection.AutoFill Destination:=Range("G1:I2"), Type:=xlFillDefault
            Range("G1:I2").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range("G1:G2").Select
            ActiveCell.FormulaR1C1 = "Qty to be delivered (in Cases/Bags)"
            Range("G1:G2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
            Range("H1:H2").Select
            ActiveCell.FormulaR1C1 = "No. of SKU's"
            Range("I1:I2").Select
            ActiveCell.FormulaR1C1 = "No of Vehicles"
            Range("J1:J2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("J1:J2").Select
            ActiveCell.FormulaR1C1 = "Appointment Require Date"
            Range("J1:J2").Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Range(Selection, Selection.End(xlToLeft)).Select
            Range(Selection, Selection.End(xlToLeft)).Select
            With Selection
                .VerticalAlignment = xlBottom
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
            With Selection
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
            Range("K1:K2").Select
            Range("K2").Activate
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Range("L1:L2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Range("M1:M2").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Range("K1:K2").Select
            ActiveCell.FormulaR1C1 = "Appointment Confirm Date"
            Range("L1:L2").Select
            ActiveCell.FormulaR1C1 = "Appointment Time"
            Range("M1:M2").Select
            ActiveCell.FormulaR1C1 = "Remarks (If any)"
            Range("M1:M2").Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("K1:L2").Select
            Range("L1").Activate
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("M1:M2").Select
            Range(Selection, Selection.End(xlToLeft)).Select
            With Selection
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
            Cells.Select
            Range("F2").Select
            Columns("J:J").ColumnWidth = 13.71
            Columns("K:K").ColumnWidth = 12.86
            Columns("L:L").ColumnWidth = 12.57
            Range("A1:M2").Select
            Range("M1").Activate
            With Selection
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
            End With
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range("A3:M30").Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Range("C1:C2").Select
            Columns("C:C").ColumnWidth = 15#
            Columns("G:G").ColumnWidth = 12#
            Columns("F:F").ColumnWidth = 10.86
            Columns("D:D").ColumnWidth = 12.43
            Columns("E:E").ColumnWidth = 11.86
            Range("A3").Select
            ActiveCell.FormulaR1C1 = "=COUNTIF(R3C2:RC[1],RC[1])"
            Range("A3:A4000").Select
            Selection.FillDown
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Range("B3").Select
            Application.CutCopyMode = False
            Range("B3:B4000").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete
            Rows("2:2").EntireRow.AutoFit
            Rows("3:4000").Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Range("A3:M4000").Select
            Range("M4000").Activate
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.End(xlUp).Select
            Columns("C:C").Select
            Selection.Insert Shift:=xlToRight
            
            Dim lastRow As Long
            lastRow = Range("B" & Rows.Count).End(xlUp).row
            For i = 3 To lastRow
                
                If Range("B" & i).Value = "R801" Then
                    Range("C" & i).Value = "KHURDA-DC"
                    End If
                    If Range("B" & i).Value = "SLMR" Then
                    Range("C" & i).Value = "KHURDA-DC"
                End If
            Next i
            
            
            Columns("D:D").Select
            
            Selection.Insert Shift:=xlToRight
            Selection.Insert Shift:=xlToRight
            Range("D3").Select
            ActiveCell.FormulaR1C1 = "-"
            Range("D3:D4000").Select
            Selection.FillDown
            Range("E3").Select
            Application.CutCopyMode = False
            ActiveCell.FormulaR1C1 = "=RC[-3]&RC[-1]&RC[-2]"
            Range("E3:E4000").Select
            Selection.FillDown
            Selection.Copy
            Application.CutCopyMode = False
            Selection.Copy
            Range("B3").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Columns("C:E").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            Selection.End(xlToLeft).Select
            Columns("A:C").Select
            Columns("A:C").EntireColumn.AutoFit
            Dim s As Long
            lastRow = Range("C" & Rows.Count).End(xlUp).row
            For s = 2 To lastRow
                
                If Range("B" & s).Value = "R801-KHURDA-DC" Then
                    Range("N" & s).Value = SettingValue("asn_email.R801-KHURDA-DC")
                    End If
                    If Range("B" & s).Value = "SLMR-KHURDA-DC" Then
                    Range("N" & s).Value = " "
                    
                    
                End If
            Next s
            
            
            Range("G3").Select
            ActiveCell.FormulaR1C1 = "=ROUND(RC[-1]*90%,0)"
            Range("G3:G4000").Select
            Selection.FillDown
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Application.CutCopyMode = False
            Range("I3").Select
            ActiveCell.FormulaR1C1 = "=ROUND(RC[-2]/RC[-3]*RC[-1],0)"
            Range("I3:I4000").Select
            Selection.FillDown
            Selection.Copy
            Range("H3").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Range("I3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Range("I3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("C4:C4000").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete
            Range("A3").Select
            
            Dim lr As Long, r As Long
            
            lr = Range("B" & Rows.Count).End(xlUp).row
            For r = lr To 4 Step -1
                If Range("B" & r).Value <> Range("B" & r - 1).Value Then Rows(r).Insert
            Next r
            
            
            Columns("N:N").EntireColumn.AutoFit
            Range("N2").Select
            ActiveCell.FormulaR1C1 = "Store Mail Id's to Send"
            Range("N1:N2").Select
            Range("N2").Activate
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 6299648
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
            Columns("N:N").Select
            Range("N2").Activate
            Selection.Insert Shift:=xlToRight
            Range("N1:N2").Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Dim Filename As String
            path = Sheets("Sheet1").Range("Y1")
            Filename = "Appointment"
            dt = Format(CStr(Now), "ddmmyyyyhhmmss")
            ActiveWorkbook.SaveAs Filename:=path & Filename & dt & ".xlsx"
            
            Sheets("Sheet1").Select
            Columns("W:AB").Select
            Selection.Delete Shift:=xlToLeft
            Run "Macro3"
            Sheets("Sheet1").Select
            Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
        
            ActiveWorkbook.Save
            Run "Macro3"
            Run "Mail_Selection_Range_Outlook_Body"
            Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformationm, "Error"
        End If
        Exit Sub
    End If
End Sub
Sub Macro3()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Columns("J:J").Select
    Application.CutCopyMode = False
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C10:R1048576C10", Version:=8).CreatePivotTable TableDestination:= _
        "Sheet1!R3100C725", TableName:="PivotTable1", DefaultVersion:=8
    Sheets("Sheet1").Select
    Cells(3100, 725).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    Application.GoTo Reference:="R3100C725"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Company")
        .Orientation = xlRowField
        .Position = 1
    End With
    Columns("AAW:AAW").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AAX3100").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""
    Range("AAX3101").Select
    
    
                Dim lastRow As Long
            Dim vendorEmail As String
            lastRow = Range("AAW" & Rows.Count).End(xlUp).row
            For s = 3 To lastRow
    
                    vendorEmail = SettingValue("vendor_email." & Trim$(CStr(Range("AAW" & s).Value)))
                    If Len(vendorEmail) > 0 Then Range("AAX" & s).Value = vendorEmail
            Next s
            
            Application.GoTo Reference:="R2C1"
    Range("A1:A2").Select
    Range("A2").Activate
    Selection.End(xlUp).Select
    
    Application.ScreenUpdating = True
    
    
End Sub


Sub Mail_Selection_Range_Outlook_Body()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    Set rng = Sheets("Sheet2").Range("A1:M20").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)


    On Error Resume Next
    With OutMail
        .To = ""
        .CC = SettingValue("appointment_cc") & Sheets("Sheet1").Range("AAX3101") & Range("AAX3102") & Range("AAX3103") & Range("AAX3104") & Range("AAX3105") & Range("AAX3106") & Range("AAX3107") & Range("AAX3108") & Range("AAX3109")
        .BCC = ""
        
        .Attachments.Add ActiveWorkbook.FullName
        .Subject = "Delivery Appointment | Hindustan Agencies @" & Sheets("Sheet2").Range("B3")
        .HTMLBody = "<BODY style = font-size:12pt;font-family:Calibri (Body)>" & _
"Hi, <br><br> Pls provide appointment for below po(s).<br>" & RangetoHTML(rng) & .HTMLBody
        .Display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
   
End Sub

 

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    Application.DisplayAlerts = True
End Function



