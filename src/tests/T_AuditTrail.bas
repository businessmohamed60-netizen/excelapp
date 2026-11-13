Attribute VB_Name = "T_AuditTrail"
Option Explicit

Public Sub T_AuditTrail()
    ResetCounters

    Dim e As CAuditTrail
    Set e = New CAuditTrail
    e.InitNew "Création", "nc", "id-123", "NC-2026-001"
    AssertTrue Len(e.Id) > 0, "Id généré"
    AssertEquals "Création", e.Action, "action"
    AssertEquals "nc", e.Collection, "collection"
    AssertEquals Environ$("USERNAME"), e.Utilisateur, "utilisateur = USERNAME"
    AssertEquals Environ$("COMPUTERNAME"), e.Poste, "poste = COMPUTERNAME"
    AssertTrue IsDate(e.DateHeure), "date définie"

    ' DisplayName
    AssertTrue InStr(e.DisplayName(), "NC-2026-001") > 0, "DisplayName contient la référence"

    PrintSummary "CAuditTrail"
End Sub