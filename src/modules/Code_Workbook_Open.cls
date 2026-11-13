Option Explicit

Private Sub Workbook_Open()
    Bootstrap
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Shutdown
End Sub