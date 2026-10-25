Attribute VB_Name = "Print_Setup_RIL"
Sub Print_Setup()
Attribute Print_Setup.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("B3").Select

    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("C:C").ColumnWidth = 12
    Columns("D:D").ColumnWidth = 11.43
    Columns("E:E").ColumnWidth = 10.57
    Columns("F:F").ColumnWidth = 4.57
    Columns("H:I").SelectS
    Selection.EntireColumn.Hidden = True
    Columns("L:L").ColumnWidth = 7.14
    Columns("K:K").ColumnWidth = 5.71
    Columns("M:M").ColumnWidth = 9
    Columns("N:N").ColumnWidth = 7.29
    Columns("O:O").ColumnWidth = 10.71
    Columns("O:O").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_ * #,##0.0_ ;_ * -#,##0.0_ ;_ * ""-""??_ ;_ @_ "
    Selection.NumberFormat = "_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ "
    Range("C7").Select
     Range("J8").Select
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Columns("J:J").ColumnWidth = 40.86
    Columns("M:M").ColumnWidth = 5.57
    Columns("B:B").ColumnWidth = 0
    Columns("A:A").ColumnWidth = 5.14
    Columns("F:F").ColumnWidth = 3.86
    Columns("N:N").ColumnWidth = 6.29
    Columns("J:J").ColumnWidth = 38.29
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints()
        .RightMargin = Application.InchesToPoints()
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints()
        .HeaderMargin = Application.InchesToPoints()
        .FooterMargin = Application.InchesToPoints()
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True
    Columns("O:O").ColumnWidth = 6.29
    Columns("O:O").ColumnWidth = 8
    Columns("C:C").ColumnWidth = 10.43
End Sub
