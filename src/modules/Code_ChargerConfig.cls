Public Type TConfig
    Version As String
    InstallDate As Date
    Langue As String
    ' ... champs existants (Phases 1-7) ...
    ' --- Phase 8 : stockage SQL ---
    DataMode As String              ' "XML" (défaut) ou "SQL"
    DbEngine As String              ' "SQLite" (seul supporté)
    DbPath As String                ' chemin complet du fichier .db
    DbAutoBackup As Boolean         ' sauvegarde auto avant migration
End Type

Public g_Config As TConfig

Public Sub ChargerConfig()
    ' ... chargement existant ...
    ' Valeurs par défaut Phase 8
    If Len(g_Config.DataMode) = 0 Then g_Config.DataMode = "XML"
    If Len(g_Config.DbEngine) = 0 Then g_Config.DbEngine = "SQLite"
    If Len(g_Config.DbPath) = 0 Then
        g_Config.DbPath = Environ$("APPDATA") & "\IATFConsultantAssistant\data\iatf.db"
    End If
End Sub

Public Sub SauvegarderConfig(ByRef cfg As TConfig)
    ' ... écriture XML existante, inclure les nouveaux champs ...
End Sub