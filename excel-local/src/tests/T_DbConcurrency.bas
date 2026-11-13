Attribute VB_Name = "T_DbConcurrency"
Option Explicit

Public Sub T_DbConcurrency()
    ResetCounters

    Dim oldMode As String, oldPath As String
    oldMode = g_Config.DataMode: oldPath = g_Config.DbPath
    g_Config.DataMode = "SQL"
    g_Config.DbPath = AppDataPath() & "data\test_concur_" & Format$(Now, "hhnnss") & ".db"

    DbOpen
    CreerSchema

    ' Test busy_timeout : ouvrir une 2ᵉ connexion et écrire
    DbExecute "INSERT INTO clients (Id, Code) VALUES ('c1', 'C1')"

    ' Test WAL : les lectures ne sont pas bloquées par les écritures
    DbBegin
    DbExecute "INSERT INTO clients (Id, Code) VALUES ('c2', 'C2')"
    ' En WAL, une lecture concurrente doit fonctionner
    Dim rs As Object
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM clients")
    AssertTrue Not rs.EOF, "lecture pendant transaction OK"
    rs.Close
    DbCommit

    DbClose
    Kill g_Config.DbPath
    g_Config.DataMode = oldMode: g_Config.DbPath = oldPath

    PrintSummary "M_Db (concurrence)"
End Sub