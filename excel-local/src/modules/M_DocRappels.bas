Option Explicit

' ============================================================
' M_DocRappels — Rappels de revue, archivage auto, rapports
' ============================================================

' ---- Documents dont la revue approche ou est en retard ----
Public Function DocumentsARevoir(ByVal joursHorizon As Long) As Collection
    Dim docs As Object, k As Variant, res As New Collection
    Set docs = g_CollMgr.LoadCollection("documents")
    For Each k In docs.Keys
        Dim d As CDocument: Set d = docs(k)
        If d.PeriodiciteRevueMois > 0 And d.Statut = "Publié" Then
            Dim j As Long: j = d.JoursAvantRevue()
            If j <= joursHorizon Then res.Add d
        End If
    Next k
    Set DocumentsARevoir = res
End Function

' ---- Archivage automatique des documents obsolètes depuis > N jours ----
Public Sub ArchiverDocumentsObsoletes(ByVal joursConserves As Long)
    Dim docs As Object, k As Variant, n As Long
    Set docs = g_CollMgr.LoadCollection("documents")
    For Each k In docs.Keys
        Dim d As CDocument: Set d = docs(k)
        If d.Statut = "Obsolète" And d.DateArchivage = 0 Then
            If DateDiff("d", d.DateCreation, Date) > joursConserves Then
                d.Statut = "Archivé"
                d.DateArchivage = Date
                g_CollMgr.Save d, "documents"
                n = n + 1
            End If
        End If
    Next k
    If n > 0 Then
        g_CollMgr.InvalidateCache "documents"
        Log LogInfo, "M_DocRappels", n & " document(s) archivé(s)"
    End If
End Sub

' ---- Intégration au système de notifications ----
Public Sub ContribuerNotifications(ByRef notifs As Collection)
    Dim docs As Collection: Set docs = DocumentsARevoir(30)
    Dim i As Long
    For i = 1 To docs.Count
        Dim d As CDocument: Set d = docs(i)
        Dim sev As NotifSeverite
        Dim j As Long: j = d.JoursAvantRevue()
        If j < 0 Then sev = NotifCritique Else sev = NotifAvertissement
        Dim titre As String, detail As String
        If j < 0 Then
            titre = "Document en retard de revue : " & d.Code
            detail = d.Titre & " — " & (-j) & " jour(s) de retard"
        Else
            titre = "Document à revoir : " & d.Code
            detail = d.Titre & " — échéance dans " & j & " jour(s)"
        End If
        Dim n As Object: Set n = CreateObject("Scripting.Dictionary")
        n("Severite") = sev
        n("Categorie") = "Document"
        n("Titre") = titre
        n("Detail") = detail
        n("Action") = "OuvrirDocument"
        n("ObjetId") = d.Id
        n("Date") = Now
        notifs.Add n
    Next i
End Sub

' ---- Statistiques pour le dashboard ----
Public Function StatistiquesDocuments() As Object
    Dim stats As Object: Set stats = CreateObject("Scripting.Dictionary")
    stats("Total") = 0: stats("Publiés") = 0: stats("En revue") = 0
    stats("Brouillons") = 0: stats("Obsolètes") = 0: stats("Archivés") = 0
    stats("En retard de revue") = 0: stats("À revoir 30j") = 0
    Dim docs As Object, k As Variant
    Set docs = g_CollMgr.LoadCollection("documents")
    For Each k In docs.Keys
        Dim d As CDocument: Set d = docs(k)
        stats("Total") = stats("Total") + 1
        Select Case d.Statut
            Case "Publié": stats("Publiés") = stats("Publiés") + 1
            Case "En revue": stats("En revue") = stats("En revue") + 1
            Case "Brouillon": stats("Brouillons") = stats("Brouillons") + 1
            Case "Obsolète": stats("Obsolètes") = stats("Obsolètes") + 1
            Case "Archivé": stats("Archivés") = stats("Archivés") + 1
        End Select
        If d.Statut = "Publié" And d.PeriodiciteRevueMois > 0 Then
            Dim j As Long: j = d.JoursAvantRevue()
            If j < 0 Then stats("En retard de revue") = stats("En retard de revue") + 1
            ElseIf j <= 30 Then stats("À revoir 30j") = stats("À revoir 30j") + 1
        End If
    Next k
    Set StatistiquesDocuments = stats
End Function