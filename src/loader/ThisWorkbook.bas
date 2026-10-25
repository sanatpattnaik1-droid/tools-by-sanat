' Code for the loader's ThisWorkbook document module.
' build-loader.ps1 pastes this into ThisWorkbook; it is not imported as a normal module.

Private Sub Workbook_Open()
    modLoader.Startup
End Sub
