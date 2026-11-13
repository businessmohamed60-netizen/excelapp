Attribute VB_Name = "T_Notify"
Option Explicit

Public Sub T_Notify()
    ResetCounters

    ' Construction d'une échéance
    Dim e As TEcheance
    e = ConstruireEcheance("CAPA", "CAPA-2026-001", "Test", Date + 5, 5, "J. Dupont")
    AssertEquals "Attention", e.Niveau, "5 jours → Attention"
    e = ConstruireEcheance("CAPA", "CAPA-2026-001", "Test", Date - 2, -2, "J. Dupont")
    AssertEquals "Critique", e.Niveau, "-2 jours → Critique"
    e = ConstruireEcheance("CAPA", "CAPA-2026-001", "Test", Date + 3, 3, "J. Dupont")
    AssertEquals "Urgent", e.Niveau, "3 jours → Urgent"
    e = ConstruireEcheance("CAPA", "CAPA-2026-001", "Test", Date + 60, 60, "J. Dupont")
    AssertEquals "Info", e.Niveau, "60 jours → Info"

    ' Corps email
    Dim c As New Collection
    c.Add ConstruireEcheance("CAPA", "CAPA-2026-001", "Test", Date + 5, 5, "J. Dupont")
    Dim html As String: html = GenererCorpsEmail(c)
    AssertTrue InStr(html, "CAPA-2026-001") > 0, "HTML contient la référence"
    AssertTrue InStr(html, "<table") > 0, "HTML contient un tableau"

    PrintSummary "M_Notify"
End Sub