Attribute VB_Name = "M_Bootstrap"
Option Explicit

' ============================================================
' M_Bootstrap — Initialisation et arrêt de l'add-in
' Version : 1.1.0
' ============================================================

Public g_Config As CConfig
Public g_CollMgr As CCollectionManager
Public g_Ribbon As Object

Public Sub Bootstrap()
    On Error GoTo ErrHandler
    Log LogInfo, "Bootstrap", "Démarrage add-in"

    ' 1. Configuration
    Set g_Config = New CConfig
    If Not g_Config.Load() Then
        g_Config.InitDefault
        g_Config.Save
        Log LogInfo, "Bootstrap", "Config créée (1re exécution)"
    End If
    SetLanguage g_Config.Language

    ' 2. Gestionnaire de collections
    Set g_CollMgr = New CCollectionManager

    ' 3. Vérification d'intégrité des données
    If Not VerifierIntegrite() Then
        If Not PromptYesNo("Intégrité", _
            "Les fichiers de données semblent modifiés." & vbCrLf & _
            "Continuer malgré tout ?") Then
            Shutdown
            Exit Sub
        End If
    End If

    ' 4. Migration de version (schéma)
    AppliquerMigrations

    ' 5. Sauvegarde automatique si nécessaire
    SauvegardeAutomatiqueSiNecessaire

    ' 6. Préchargement du cache (performances)
    PrechargerCache

    ' 7. Planification des notifications (30 min)
    PlanifierNotifications

    ' 8. Contrôle de licence
    If EtatLicence() = LicenceExpiree Then
        ShowError "Licence", _
            "Licence expirée. Enregistrez une clé valide dans Paramètres > Licence."
        Log LogWarn, "Bootstrap", "Licence expirée"
    End If

    Log LogInfo, "Bootstrap", "Add-in prêt v" & g_Config.Version & _
        " — " & MessageEtatLicence()
    Exit Sub

ErrHandler:
    LogEx "Bootstrap", "Bootstrap()", Err
    ShowError "Démarrage", "Erreur : " & Err.Description
End Sub

Public Sub Shutdown()
    On Error Resume Next
    AnnulerPlanification
    If Not g_Config Is Nothing Then g_Config.Save
    Set g_CollMgr = Nothing
    Set g_Config = Nothing
    Application.StatusBar = False
    Log LogInfo, "Bootstrap", "Arrêt add-in"
End Sub