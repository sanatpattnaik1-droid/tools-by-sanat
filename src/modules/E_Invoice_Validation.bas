Attribute VB_Name = "E_Invoice_Validation"
Sub E_Invoice_Validation()
Dim answer As Integer
answer = MsgBox("The Current Sheet Data Will Process Do You Want To Continue ?", vbQuestion + vbYesNo + vbDefaultButton1, "Actoin")
   If answer = vbYes Then
If Range("A5").Value = "SUPPLY_TYPE_CODE" Then
 Dim jsonFile As Variant
 jsonFile = Environ("USERPROFILE") & "\Downloads\JSON_preperation_FormA.xlsm"
 If Dir(jsonFile) = "" Then
     jsonFile = Application.GetOpenFilename("Excel Macro Files (*.xlsm),*.xlsm", , "Select JSON_preperation_FormA.xlsm")
     If jsonFile = False Then Exit Sub
 End If
 Application.ScreenUpdating = False
    Workbooks.Open jsonFile
    Windows("JSON_preperation_FormA.xlsm").Activate
    Dim wb As Workbook

For Each wb In Application.Workbooks
   If wb.name Like "EINVOICE*" Then wb.Activate
   Next wb
    Rows("6:6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("JSON_preperation_FormA.xlsm").Activate
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("JSON_preperation_FormA.xlsm").Activate
    Application.ScreenUpdating = True
    MsgBox "Click Validate Button To Process.", vbInformation, "Status"
Else: MsgBox "Column Sequence Doesn't Match Process Not Completed!.", vbInformation, "Error"
End If
Exit Sub
End If
End Sub


