Attribute VB_Name = "T_KpiSql"
Option Explicit

Public Sub T_KpiSql()
    ResetCounters

    ' Configurer une base de test avec quelques données
    Dim oldEngine As String, oldPath As String, oldMode As String
    oldEngine = g_Config.DbEngine: oldPath = g_Config.DbPath: oldMode = g_Config.DataMode
    g_Config.DbEngine = "SQLite"
    g_Config.DbPath = Environ$("TEMP") & "\test_kpi_" & Format$(Now, "hhnnss") & ".db"
    g_Config.DataMode = "SQL"

    DbOpen
    CreerSchema

    ' Insérer des NC
    DbExecute "INSERT INTO nc (Id, Numero, ClientId, Description, Statut, DateDetection) VALUES ('1', 'NC-2026-001', 'c1', 'Test', 'Ouverte', '2026-01-15')"
    DbExecute "INSERT INTO nc (Id, Numero, ClientId, Description, Statut, DateDetection) VALUES ('2', 'NC-2026-002', 'c1', 'Test', 'Clôturée', '2026-02-15')"

    ' KPI
    Dim kpi As Collection
    Set kpi = CalculerTousKpiSql(2026)
    AssertTrue kpi.Count > 0, "KPI calculés"

    Dim i As Long, k As TKpi, found As Boolean
    For i = 1 To kpi.Count
        k = kpi(i)
        If k.Nom = "NC ouvertes" Then
            AssertEquals 1#, k.Valeur, "1 NC ouverte"
            found = True
        End If
    Next i
    AssertTrue found, "KPI NC ouvertes présent"

    DbClose
    Kill g_Config.DbPath
    g_Config.DbEngine = oldEngine: g_Config.DbPath = oldPath: g_Config.DataMode = oldMode

    PrintSummary "M_KpiSql"
End Sub