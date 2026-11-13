Attribute VB_Name = "T_Storage"
Option Explicit

Public Sub T_Storage()
    ResetCounters

    ' Hash stable pour un même contenu
    Dim tempFile As String
    tempFile = Environ$("TEMP") & "\test_hash_" & Format$(Now, "hhnnss") & ".txt"
    Dim f As Integer: f = FreeFile
    Open tempFile For Output As #f: Print #f, "contenu test": Close #f
    Dim h1 As String, h2 As String
    h1 = HashFichier(tempFile)
    h2 = HashFichier(tempFile)
    AssertEquals h1, h2, "hash stable pour même contenu"
    AssertTrue Len(h1) = 8, "hash de 8 caractères"

    ' Hash différent pour un contenu modifié
    Dim f2 As Integer: f2 = FreeFile
    Open tempFile For Output As #f2: Print #f2, "contenu modifié": Close #f2
    Dim h3 As String: h3 = HashFichier(tempFile)
    AssertTrue h1 <> h3, "hash différent après modification"

    ' VerifierConflit
    AssertTrue VerifierConflit(tempFile, h1), "conflit détecté si hash attendu diffère"
    AssertFalse VerifierConflit(tempFile, h3), "pas de conflit si hash identique"
    AssertFalse VerifierConflit(tempFile, ""), "pas de conflit si pas de hash attendu"

    ' DataPath non vide
    AssertTrue Len(DataPath()) > 0, "DataPath non vide"
    AssertTrue Right$(DataPath(), 1) = "\", "DataPath se termine par \"

    ' TemplatePath
    AssertTrue InStr(TemplatePath(), "templates") > 0, "TemplatePath contient templates"

    Kill tempFile
    PrintSummary "M_Storage"
End Sub