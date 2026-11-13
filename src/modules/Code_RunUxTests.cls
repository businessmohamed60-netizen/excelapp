Option Explicit

Public Sub RunUxTests()
    ResetCounters
    Dim res As Collection: Set res = RechercherGlobal("a")
    AssertEquals 0&, res.Count, "requête trop courte = 0"
    Set res = RechercherGlobal("test")
    AssertNotNothing res, "recherche valide"
    AssertEquals "🟢", IndicateurStatut("Clôturée"), "statut vert"
    AssertEquals "🔴", IndicateurStatut("Ouverte"), "statut rouge"
    AssertEquals "🟠", IndicateurEcheance(Date + 3, "En cours"), "échéance proche orange"
    PrintSummary "UX"
End Sub