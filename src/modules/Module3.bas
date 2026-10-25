Attribute VB_Name = "Module3"
Sub SplitData()

Dim objWorksheet As Excel.Worksheet
Dim nLastRow, nRow, nNextRow As Integer
Dim strColumnValue As String
Dim objDictionary As Object
Dim varColumnValues As Variant
Dim varColumnValue As Variant
Dim objExcelWorkbook As Excel.Workbook
Dim objSheet As Excel.Worksheet

Dim aCol As String
aCol = "D"

On Error GoTo err1

 With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.EnableEvents = False
 End With

Set objWorksheet = ActiveSheet
nLastRow = objWorksheet.Range("A" & objWorksheet.Rows.Count).End(xlUp).row

Set objDictionary = CreateObject("Scripting.Dictionary")

 For nRow = 1 To nLastRow

    strColumnValue = Format(objWorksheet.Range(aCol & nRow).Value, "Report_mmm_yyyy")

    If objDictionary.Exists(strColumnValue) = False Then
       objDictionary.Add strColumnValue, 1
    End If
Next

varColumnValues = objDictionary.Keys

For i = LBound(varColumnValues) To UBound(varColumnValues)
    varColumnValue = varColumnValues(i)


    Set objExcelWorkbook = Excel.Application.Workbooks.Add

    Set objSheet = objExcelWorkbook.Sheets(1)
    objSheet.name = objWorksheet.name

    objWorksheet.Rows(1).EntireRow.Copy
    objSheet.Activate
    objSheet.Range("A1:AT12685").Select
    objSheet.Paste

     For nRow = 1 To nLastRow
        If CStr(objWorksheet.Range(aCol & nRow).Value) = CStr(varColumnValue) Then

           objWorksheet.Rows(nRow).EntireRow.Copy

           nNextRow = objSheet.Range("A" & objWorksheet.Rows.Count).End(xlUp).row + 1
           objSheet.Range("A" & nNextRow).Select
           objSheet.Paste
           objSheet.Columns("A:I").AutoFit
        End If
    Next
    objExcelWorkbook.SaveAs (CStr(varColumnValue))
Next

err1:
     With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
     End With

End Sub
