Public Sub Rb_Migration(ByVal control As IRibbonControl): FMigration.Show vbModal: End Sub
Public Sub Rb_Mode(ByVal control As IRibbonControl)
    Dim mode As String
    mode = InputBox("Mode de données (XML / SQL) :", "Mode", g_Config.DataMode)
    If mode = "XML" Or mode = "SQL" Then
        g_Config.DataMode = mode
        SauvegarderConfig g_Config
        ShowInfo "Mode", "Mode " & mode & " activé. Redémarrez Excel."
    End If
End Sub
Public Sub Rb_Api(ByVal control As IRibbonControl): FApi.Show vbModeless: End Sub
Public Sub Rb_PowerBi(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim p1 As String, p2 As String
    p1 = GenererPbids()
    p2 = GenererPowerQuery()
    If PromptYesNo("Power BI", "Fichiers générés :" & vbCrLf & p1 & vbCrLf & p2 & vbCrLf & "Ouvrir le dossier ?") Then
        Shell "explorer.exe """ & AppDataPath() & "reports""", vbNormalFocus
    End If
    Exit Sub
ErrHandler:
    ShowError "Power BI", Err.Description
End Sub