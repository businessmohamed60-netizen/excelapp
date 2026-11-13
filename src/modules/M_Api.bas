Attribute VB_Name = "M_Api"
Option Explicit

' ============================================================
' M_Api — Façade publique du domaine (point d'entrée stable)
' Utilisé par le Ribbon et les scripts externes
' ============================================================

Public Sub Api_OuvrirDashboard()
    EnsureBootstrap
    FMain.Show vbModeless
End Sub

Public Sub Api_OuvrirClients()
    EnsureBootstrap
    FClient.Show vbModeless
End Sub

Public Sub Api_OuvrirAudits()
    EnsureBootstrap
    FAudit.Show vbModeless
End Sub

Public Sub Api_OuvrirNc()
    EnsureBootstrap
    FNonConformite.Show vbModeless
End Sub

Public Sub Api_OuvrirAmdec()
    EnsureBootstrap
    FAmdec.Show vbModeless
End Sub

Public Sub Api_OuvrirPpap()
    EnsureBootstrap
    FPpap.Show vbModeless
End Sub

Public Sub Api_OuvrirApqp()
    EnsureBootstrap
    FApqa.Show vbModeless
End Sub

Public Sub Api_OuvrirSpc()
    EnsureBootstrap
    FSpc.Show vbModeless
End Sub

Public Sub Api_OuvrirMsa()
    EnsureBootstrap
    FMsa.Show vbModeless
End Sub

Public Sub Api_OuvrirNotifications()
    EnsureBootstrap
    FNotifications.Show vbModeless
End Sub

Public Sub Api_OuvrirRecherche()
    EnsureBootstrap
    FRecherche.Show vbModeless
End Sub

Public Sub Api_OuvrirParametres()
    EnsureBootstrap
    FSettings.Show vbModeless
End Sub

Public Sub Api_OuvrirLicence()
    EnsureBootstrap
    FLicence.Show vbModal
End Sub

Public Function Api_Version() As String
    If g_Config Is Nothing Then
        Api_Version = "non initialisé"
    Else
        Api_Version = g_Config.Version
    End If
End Function

Public Function Api_EtatLicence() As String
    Api_EtatLicence = MessageEtatLicence()
End Function

' ---- Helpers internes ----
Private Sub EnsureBootstrap()
    If g_Config Is Nothing Then Bootstrap
End Sub