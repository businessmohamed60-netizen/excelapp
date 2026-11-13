Option Explicit

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set g_Ribbon = ribbon
    Log LogInfo, "Ribbon", "Ribbon chargé"
End Sub

Public Sub Ribbon_OuvrirDashboard(ByVal control As Object): EnsureBootstrap: FMain.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirClients(ByVal control As Object): EnsureBootstrap: FClient.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirAudits(ByVal control As Object): EnsureBootstrap: FAudit.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirNC(ByVal control As Object): EnsureBootstrap: FNonConformite.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirApqp(ByVal control As Object): EnsureBootstrap: FApqa.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirPpap(ByVal control As Object): EnsureBootstrap: FPpap.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirAmdec(ByVal control As Object): EnsureBootstrap: FAmdec.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirSpc(ByVal control As Object): EnsureBootstrap: FSpc.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirMsa(ByVal control As Object): EnsureBootstrap: FMsa.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirNotifications(ByVal control As Object): EnsureBootstrap: FNotifications.Show vbModeless: End Sub
Public Sub Ribbon_OuvrirRecherche(ByVal control As Object): EnsureBootstrap: FRecherche.Show vbModeless: End Sub

Public Sub Ribbon_OuvrirSettings(ByVal control As Object)
    EnsureBootstrap
    FSettings.Show vbModeless
End Sub

Public Sub Ribbon_OuvrirAbout(ByVal control As Object)
    EnsureBootstrap
    ShowInfo T("menu.about"), "IATF Consultant Assistant v" & g_Config.Version & vbCrLf & _
        "Édition Premium — " & MessageEtatLicence()
End Sub

Private Sub EnsureBootstrap()
    If g_Config Is Nothing Then Bootstrap
End Sub