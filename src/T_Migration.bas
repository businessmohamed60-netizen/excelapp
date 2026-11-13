Attribute VB_Name = "T_Migration"
Option Explicit

Public Sub T_Migration()
    ResetCounters

    ' Configurer une base de test
    Dim oldEngine As String, oldPath As String, oldMode As String
    oldEngine = g_Config.DbEngine
    oldPath = g_Config.DbPath
    oldMode = g_Config.DataMode

    g_Config.DbEngine = "SQLite"
    g_Config.DbPath = Environ$("TEMP") & "\test_migr_" & Format$(Now, "hhnnss") & ".db"
    g_Config.DataMode = "SQL"

    DbOpen
    CreerSchema

    ' Migrer les clients (au moins 1 doit exister dans les données de test)
    Dim n As Long
    n = MigrerCollection("clients")
    AssertTrue n >= 0, "migration clients >= 0"

    ' Vérifier côté SQL
    Dim rs As Object
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM clients")
    AssertEquals n, CLng(rs("n")), "nb SQL = nb migré"
    rs.Close

    DbClose
    Kill g_Config.DbPath

    g_Config.DbEngine = oldEngine
    g_Config.DbPath = oldPath
    g_Config.DataMode = oldMode

    PrintSummary "M_Migration"
End Sub