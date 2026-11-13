Public Sub T_Migration()
    ResetCounters

    Dim oldEngine As String, oldPath As String, oldMode As String
    oldEngine = g_Config.DbEngine
    oldPath = g_Config.DbPath
    oldMode = g_Config.DataMode

    g_Config.DbEngine = "SQLite"
    g_Config.DbPath = AppDataPath() & "data\test_migr_" & Format$(Now, "hhnnss") & ".db"
    g_Config.DataMode = "XML"   ' forcer XML pour la migration

    DbOpen
    CreerSchema

    ' Migrer clients
    Dim n As Long
    n = MigrerCollection("clients")
    AssertTrue n >= 0, "migration clients >= 0"

    ' Vérifier côté SQL
    AssertEquals n, DbCount("clients"), "nb SQL = nb migré"

    DbClose
    Kill g_Config.DbPath

    g_Config.DbEngine = oldEngine
    g_Config.DbPath = oldPath
    g_Config.DataMode = oldMode

    PrintSummary "M_Migration"
End Sub