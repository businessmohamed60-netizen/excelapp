Public Sub Bootstrap()
    ChargerConfig
    Set g_CollMgr = New CCollectionManager
    Set g_SessionHashes = CreateObject("Scripting.Dictionary")
    If Not VerifierStockage() Then
        ShowError "Stockage", "Le chemin de stockage est inaccessible : " & DataPath() & _
            vbCrLf & "Vérifiez les paramètres (onglet Administration)."
    End If
    ConstruireRuban
    If g_Config.NotificationsAuto Then VerifierEcheancesAuDemarrage
    Log LogInfo, "Bootstrap", "Démarrage OK — mode " & g_Config.StorageMode
End Sub