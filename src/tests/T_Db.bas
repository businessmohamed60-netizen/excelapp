Attribute VB_Name = "T_Db"
Option Explicit

Public Sub T_Db()
    ResetCounters

    ' Configurer une base de test temporaire
    Dim oldEngine As String, oldPath As String, oldMode As String
    oldEngine = g_Config.DbEngine
    oldPath = g_Config.DbPath
    oldMode = g_Config.DataMode

    g_Config.DbEngine = "SQLite"
    g_Config.DbPath = Environ$("TEMP") & "\test_iatf_" & Format$(Now, "hhnnss") & ".db"
    g_Config.DataMode = "SQL"

    ' Ouvrir
    AssertTrue DbOpen(), "connexion SQLite ouverte"
    AssertTrue DbTest(), "test connexion OK"

    ' Créer une table minimale et insérer
    DbExecute "CREATE TABLE IF NOT EXISTS test (Id TEXT PRIMARY KEY, Nom TEXT)"
    DbExecute "INSERT INTO test (Id, Nom) VALUES (?, ?)", Array("1", "Alice")
    DbExecute "INSERT INTO test (Id, Nom) VALUES (?, ?)", Array("2", "Bob")

    ' Query
    Dim rs As Object
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM test")
    AssertEquals 2, CLng(rs("n")), "2 lignes insérées"
    rs.Close

    ' Update
    DbExecute "UPDATE test SET Nom = ? WHERE Id = ?", Array("Alice2", "1")
    Set rs = DbQuery("SELECT Nom FROM test WHERE Id = ?", Array("1"))
    AssertEquals "Alice2", CStr(rs("Nom")), "update OK"
    rs.Close

    ' Transaction rollback
    DbBegin
    DbExecute "INSERT INTO test (Id, Nom) VALUES (?, ?)", Array("3", "Charlie")
    DbRollback
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM test")
    AssertEquals 2, CLng(rs("n")), "rollback annule l'insertion"
    rs.Close

    ' Commit
    DbBegin
    DbExecute "INSERT INTO test (Id, Nom) VALUES (?, ?)", Array("4", "Dave")
    DbCommit
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM test")
    AssertEquals 3, CLng(rs("n")), "commit conserve l'insertion"
    rs.Close

    DbClose

    ' Nettoyage
    Kill g_Config.DbPath

    ' Restauration
    g_Config.DbEngine = oldEngine
    g_Config.DbPath = oldPath
    g_Config.DataMode = oldMode

    PrintSummary "M_Db"
End Sub