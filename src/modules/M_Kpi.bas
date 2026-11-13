Option Explicit

' ============================================================
' M_Kpi — Calcul des 12 KPI avec tendances et scoring
' ============================================================

' Structure d'un KPI (retournée comme Dictionary)
'  - Cle : identifiant
'  - Valeur : numérique
'  - Label : texte
'  - Icone : emoji
'  - Niveau : "ok" / "avertissement" / "critique" / "neutre"
'  - Tendance : "hausse" / "baisse" / "stable"
'  - Variation : différence avec la période précédente (numérique)

Public Function CalculerTousKpis() As Collection
    Dim kpis As New Collection
    On Error GoTo ErrHandler
    kpis.Add CalculerKpiClients()
    kpis.Add CalculerKpiAuditsEnCours()
    kpis.Add CalculerKpiNcOuvertes()
    kpis.Add CalculerKpiNcMajeures()
    kpis.Add CalculerKpiCapaRetard()
    kpis.Add CalculerKpiTauxCloture()
    kpis.Add CalculerKpiApqpEnCours()
    kpis.Add CalculerKpiPpapAttente()
    kpis.Add CalculerKpiAmdecH()
    kpis.Add CalculerKpiAlertesCritiques()
    kpis.Add CalculerKpiScoreGlobal()
    kpis.Add CalculerKpiDerniereActivite()
    Set CalculerTousKpis = kpis
    Exit Function
ErrHandler:
    LogEx "M_Kpi", "CalculerTousKpis", Err
    Set CalculerTousKpis = kpis
End Function

' ---- Helpers ----
Private Function CreerKpi(ByVal cle As String, ByVal valeur As Double, ByVal label As String, _
                          ByVal niveau As String, Optional ByVal tendance As String = "stable", _
                          Optional ByVal variation As Double = 0) As Object
    Dim k As Object
    Set k = CreateObject("Scripting.Dictionary")
    k("Cle") = cle
    k("Valeur") = valeur
    k("Label") = label
    k("Icone") = IconeKpi(cle)
    k("Niveau") = niveau
    k("Tendance") = tendance
    k("Variation") = variation
    Set CreerKpi = k
End Function

Private Function FormaterValeur(ByVal v As Double) As String
    If v = Int(v) Then FormaterValeur = Format$(v, "#,##0") Else FormaterValeur = Format$(v, "0.0")
End Function

' ---- KPI 1 : Clients ----
Private Function CalculerKpiClients() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("clients")
    Dim actifs As Long, k As Variant
    For Each k In d.Keys: If d(k).Actif Then actifs = actifs + 1: Next k
    Set CalculerKpiClients = CreerKpi("clients", actifs, "Clients actifs", "neutre")
End Function

' ---- KPI 2 : Audits en cours ----
Private Function CalculerKpiAuditsEnCours() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("audits")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).Statut = "En cours" Then n = n + 1
    Next k
    Dim niveau As String: niveau = IIf(n > 5, "avertissement", "neutre")
    Set CalculerKpiAuditsEnCours = CreerKpi("audits_en_cours", n, "Audits en cours", niveau)
End Function

' ---- KPI 3 : NC ouvertes ----
Private Function CalculerKpiNcOuvertes() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("nc")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).Statut <> "Clôturée" Then n = n + 1
    Next k
    Set CalculerKpiNcOuvertes = CreerKpi("nc", n, "NC ouvertes", IIf(n > 10, "avertissement", "ok"))
End Function

' ---- KPI 4 : NC majeures ----
Private Function CalculerKpiNcMajeures() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("nc")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).Statut <> "Clôturée" And (d(k).Classification = "Majeur" Or d(k).Classification = "Critique") Then n = n + 1
    Next k
    Set CalculerKpiNcMajeures = CreerKpi("nc_majeures", n, "dont majeures/critiques", IIf(n > 0, "critique", "ok"))
End Function

' ---- KPI 5 : CAPA en retard ----
Private Function CalculerKpiCapaRetard() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("capa")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).EstEnRetard() Then n = n + 1
    Next k
    Set CalculerKpiCapaRetard = CreerKpi("capa_retard", n, "CAPA en retard", IIf(n > 0, "critique", "ok"))
End Function

' ---- KPI 6 : Taux de clôture audits ----
Private Function CalculerKpiTauxCloture() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("audits")
    Dim total As Long, clotures As Long, k As Variant
    For Each k In d.Keys
        total = total + 1
        If d(k).Statut = "Clôturé" Then clotures = clotures + 1
    Next k
    Dim taux As Double
    If total > 0 Then taux = clotures / total * 100
    Dim niveau As String
    If taux >= 80 Then niveau = "ok" ElseIf taux >= 50 Then niveau = "avertissement" Else niveau = "critique"
    Set CalculerKpiTauxCloture = CreerKpi("taux_cloture", taux, "Taux de clôture audits (%)", niveau)
End Function

' ---- KPI 7 : APQP en cours ----
Private Function CalculerKpiApqpEnCours() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("apqp")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).Statut = "En cours" Then n = n + 1
    Next k
    Set CalculerKpiApqpEnCours = CreerKpi("apqp", n, "Projets APQP en cours", "neutre")
End Function

' ---- KPI 8 : PPAP à soumettre ----
Private Function CalculerKpiPpapAttente() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("ppap")
    Dim n As Long, k As Variant
    For Each k In d.Keys
        If d(k).Statut = "En préparation" Then n = n + 1
    Next k
    Set CalculerKpiPpapAttente = CreerKpi("ppap", n, "PPAP à soumettre", IIf(n > 3, "avertissement", "neutre"))
End Function

' ---- KPI 9 : Actions AMDEC H ----
Private Function CalculerKpiAmdecH() As Object
    Dim d As Object: Set d = g_CollMgr.LoadCollection("amdec")
    Dim n As Long, k As Variant
    For Each k In d.Keys: n = n + d(k).NbActionsH(): Next k
    Set CalculerKpiAmdecH = CreerKpi("amdec_h", n, "Actions AMDEC haute priorité", IIf(n > 0, "critique", "ok"))
End Function

' ---- KPI 10 : Alertes critiques ----
Private Function CalculerKpiAlertesCritiques() As Object
    Dim notifs As Collection: Set notifs = CollecterNotifications()
    Dim n As Long: n = CompterParSeverite(notifs, NotifCritique)
    Set CalculerKpiAlertesCritiques = CreerKpi("alertes", n, "Alertes critiques", IIf(n > 0, "critique", "ok"))
End Function

' ---- KPI 11 : Score global de santé qualité (0-100) ----
Private Function CalculerKpiScoreGlobal() As Object
    Dim score As Double: score = 100
    Dim kpis As Collection, k As Object, i As Long
    ' Pénalités pondérées
    Dim d As Object
    Set d = g_CollMgr.LoadCollection("capa")
    Dim nCapa As Long, kk As Variant
    For Each kk In d.Keys: If d(kk).EstEnRetard() Then nCapa = nCapa + 1: Next kk
    score = score - nCapa * 3
    Set d = g_CollMgr.LoadCollection("nc")
    Dim nMaj As Long
    For Each kk In d.Keys
        If d(kk).Statut <> "Clôturée" And (d(kk).Classification = "Majeur" Or d(kk).Classification = "Critique") Then nMaj = nMaj + 1
    Next kk
    score = score - nMaj * 5
    Set d = g_CollMgr.LoadCollection("audits")
    Dim total As Long, clotures As Long
    For Each kk In d.Keys
        total = total + 1
        If d(kk).Statut = "Clôturé" Then clotures = clotures + 1
    Next kk
    If total > 0 Then score = score + (clotures / total - 0.5) * 20
    If score < 0 Then score = 0
    If score > 100 Then score = 100
    Dim niveau As String
    If score >= 80 Then niveau = "ok" ElseIf score >= 50 Then niveau = "avertissement" Else niveau = "critique"
    Set CalculerKpiScoreGlobal = CreerKpi("score_global", score, "Score global santé qualité", niveau)
End Function

' ---- KPI 12 : Dernière activité ----
Private Function CalculerKpiDerniereActivite() As Object
    Dim d As Object, k As Variant, derniere As Date
    Set d = g_CollMgr.LoadCollection("nc")
    For Each k In d.Keys
        If d(k).DateCreation > derniere Then derniere = d(k).DateCreation
    Next k
    Set d = g_CollMgr.LoadCollection("audits")
    For Each k In d.Keys
        If d(k).DateCreation > derniere Then derniere = d(k).DateCreation
    Next k
    Dim jours As Long
    If derniere > 0 Then jours = DateDiff("d", derniere, Date) Else jours = 999
    Set CalculerKpiDerniereActivite = CreerKpi("score_global", jours, "Jours depuis la dernière activité", IIf(jours > 7, "avertissement", "ok"))
End Function