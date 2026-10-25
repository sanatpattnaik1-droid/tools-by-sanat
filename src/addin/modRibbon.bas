Attribute VB_Name = "modRibbon"
Option Explicit

' Bumped by build.ps1 from version.json. Shown on the ribbon.
Public Const ADDIN_VERSION As String = "0.1.0"

' Every ribbon button passes its macro name in the control's Tag.
Public Sub RunMacro(control As IRibbonControl)
    On Error GoTo Fail
    Application.Run "'" & ThisWorkbook.Name & "'!" & control.Tag
    Exit Sub
Fail:
    MsgBox "Could not run '" & control.Tag & "'." & vbCrLf & Err.Description, vbExclamation, "Tools By Sanat"
End Sub

' Opens the folder with settings.csv and party_master.csv (creates empty templates if missing).
Public Sub OpenSettings(control As IRibbonControl)
    modSettings.OpenSettingsFolder
End Sub

' Read by the loader to decide whether an update is needed.
Public Function AddinVersion() As String
    AddinVersion = ADDIN_VERSION
End Function

' The loader (separate add-in) downloads and stages updates; they apply at the next Excel start.
Public Sub CheckForUpdates(control As IRibbonControl)
    On Error GoTo NoLoader
    Application.Run "'ToolsBySanatLoader.xlam'!modLoader.CheckNow"
    Exit Sub
NoLoader:
    MsgBox "The update helper is not installed. Run Install.bat from the Tools By Sanat setup folder.", vbExclamation, "Tools By Sanat"
End Sub

Public Sub GetVersionLabel(control As IRibbonControl, ByRef label)
    label = "Version " & ADDIN_VERSION
End Sub
