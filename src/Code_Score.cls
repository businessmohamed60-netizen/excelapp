Option Explicit

Private Function Score(ByVal texte As String, ByVal q As String) As Long
    Dim t As String, lq As String
    t = LCase$(texte): lq = LCase$(q)
    If Len(t) = 0 Or Len(lq) = 0 Then Exit Function
    If t = lq Then Score = 100: Exit Function
    If Left$(t, Len(lq)) = lq Then Score = 80: Exit Function
    If InStr(t, lq) > 0 Then Score = 50: Exit Function
End Function

Public Function RechercherGlobal(ByVal requete As String, Optional ByVal maxParType As Long = 20) As Collection
    Dim resultats As New Collection
    Dim q As String: q = Trim$(requete)
    If Len(q) < 2 Then Set RechercherGlobal = resultats: Exit Function
    ChercherClients q, resultats
    ChercherAudits q, resultats
    ChercherConstats q, resultats
    ChercherNc q, resultats
    ChercherCapa q, resultats
    ChercherAmdec q, resultats
    ChercherPpap q, resultats
    Set RechercherGlobal = TrierParScore(resultats)
End Function

Private Sub ChercherClients(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long, c As CClient
    Set d = g_CollMgr.LoadCollection("clients")
    For Each k In d.Keys
        Set c = d(k)
        s = Score(c.RaisonSociale & " " & c.Code & " " & c.ContactNom, q)
        If s > 0 Then res.Add CreerResultat("Client", c.DisplayName, "Contact : " & c.ContactNom, c.Id, s, "OuvrirClient")
    Next k
End Sub

Private Sub ChercherAudits(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("audits")
    For Each k In d.Keys
        s = Score(d(k).Processus & " " & d(k).Auditeur & " " & d(k).TypeAudit, q)
        If s > 0 Then res.Add CreerResultat("Audit", d(k).Processus & " (" & d(k).Statut & ")", "Auditeur : " & d(k).Auditeur & " — " & FormatIsoDate(d(k).DateDebut), d(k).Id, s, "OuvrirAudit")
    Next k
End Sub

Private Sub ChercherConstats(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("constats")
    For Each k In d.Keys
        s = Score(d(k).Description & " " & d(k).Chapitre, q)
        If s > 0 Then res.Add CreerResultat("Constat", d(k).Classification & " — " & d(k).Chapitre, Left$(d(k).Description, 80), d(k).Id, s, "OuvrirConstat")
    Next k
End Sub

Private Sub ChercherNc(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("nc")
    For Each k In d.Keys
        s = Score(d(k).Numero & " " & d(k).Description & " " & d(k).Processus, q)
        If s > 0 Then res.Add CreerResultat("NC", d(k).Numero & " — " & d(k).Classification, Left$(d(k).Description, 80), d(k).Id, s, "OuvrirNc")
    Next k
End Sub

Private Sub ChercherCapa(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("capa")
    For Each k In d.Keys
        s = Score(d(k).Numero & " " & d(k).Description & " " & d(k).Responsable, q)
        If s > 0 Then res.Add CreerResultat("CAPA", d(k).Numero & " (" & d(k).Statut & ")", Left$(d(k).Description, 80), d(k).Id, s, "OuvrirCapa")
    Next k
End Sub

Private Sub ChercherAmdec(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("amdec")
    For Each k In d.Keys
        s = Score(d(k).Reference & " " & d(k).Piece & " " & d(k).Processus, q)
        If s > 0 Then res.Add CreerResultat("AMDEC", d(k).Reference & " — " & d(k).Piece, d(k).TypeAmdec & " — RPN max " & d(k).RPNMax(), d(k).Id, s, "OuvrirAmdec")
    Next k
End Sub

Private Sub ChercherPpap(ByVal q As String, ByRef res As Collection)
    Dim d As Object, k As Variant, s As Long
    Set d = g_CollMgr.LoadCollection("ppap")
    For Each k In d.Keys
        s = Score(d(k).Reference & " " & d(k).Piece & " " & d(k).ReferencePiece, q)
        If s > 0 Then res.Add CreerResultat("PPAP", d(k).Reference & " — " & d(k).Piece, d(k).Statut & " — Niveau " & d(k).NiveauPpap, d(k).Id, s, "OuvrirPpap")
    Next k
End Sub

Private Function CreerResultat(ByVal type As String, ByVal titre As String, _
                               ByVal detail As String, ByVal id As String, _
                               ByVal score As Long, ByVal action As String) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("Type") = type: r("Titre") = titre: r("Detail") = detail
    r("Id") = id: r("Score") = score: r("Action") = action
    Set CreerResultat = r
End Function

Private Function TrierParScore(ByVal col As Collection) As Collection
    Dim arr() As Object, i As Long, j As Long, n As Long
    n = col.Count
    If n = 0 Then Set TrierParScore = col: Exit Function
    ReDim arr(0 To n - 1)
    For i = 1 To n: Set arr(i - 1) = col(i): Next i
    Dim tmp As Object
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If arr(j)("Score") > arr(i)("Score") Then Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
        Next j
    Next i
    Dim res As New Collection
    For i = 0 To n - 1: res.Add arr(i): Next i
    Set TrierParScore = res
End Function