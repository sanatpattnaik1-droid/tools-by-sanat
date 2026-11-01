Attribute VB_Name = "modSaveOutput"
Option Explicit

' Saves a freshly built output workbook in the folder of the file it was made from, as
' "<prefix> Open PO- dd-mm-yyyy hhmmss.xlsx" (a " (2)" is added if that name is already taken).
' Returns the full path. If it cannot be saved it returns "" and puts the reason in problem; the
' workbook then stays open and unsaved so nothing is lost.
Public Function SaveBesideSource(wb As Workbook, ByVal srcFolder As String, ByVal prefix As String, ByRef problem As String) As String
    Dim sep As String, target As String, stem As String, n As Long, oldAlerts As Boolean

    problem = ""
    If Len(srcFolder) = 0 Then
        problem = "the original file has not been saved yet, so there is no folder to save into"
        Exit Function
    End If

    If LCase$(Left$(srcFolder, 4)) = "http" Then sep = "/" Else sep = "\"   ' OneDrive / SharePoint paths are URLs
    stem = srcFolder & sep & prefix & " Open PO- " & Format$(Now, "dd-mm-yyyy hhnnss")
    target = stem & ".xlsx"
    If sep = "\" Then
        Do While Len(Dir$(target)) > 0
            n = n + 1
            target = stem & " (" & n & ").xlsx"
        Loop
    End If

    oldAlerts = Application.DisplayAlerts
    On Error GoTo Fail
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=target, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = oldAlerts
    SaveBesideSource = wb.FullName
    Exit Function

Fail:
    problem = Err.Description
    Application.DisplayAlerts = oldAlerts
End Function
