Attribute VB_Name = "T_AuditProgramme"
Option Explicit

Public Sub T_AuditProgramme()
    ResetCounters
    Dim p As CAuditProgramme
    Set p = New CAuditProgramme
    p.InitNew
    AssertTrue Len(p.Id) > 0, "Id généré"
    AssertEquals "Planifié", p.Statut, "statut par défaut"
    AssertEquals "Processus", p.TypeAudit, "type par défaut"
    AssertEquals "Moyen", p.Risque, "risque par défaut"
    AssertEquals 12, p.FrequenceMois, "fréquence par défaut"
    AssertEquals Year(Date), p.Annee, "année par défaut"

    ' Validation
    AssertFalse p.Validate(), "validation KO sans processus"
    p.Processus = "Production"
    p.DatePrevue = DateSerial(2026, 3, 15)
    AssertTrue p.Validate(), "validation OK"

    ' EstEnRetard / EstEchu
    p.DatePrevue = Date - 5: p.Statut = "Planifié"
    AssertTrue p.EstEnRetard(), "en retard si date passée"
    p.DatePrevue = Date + 10
    AssertTrue p.EstEchu(), "échu si <= 30 jours"
    p.DatePrevue = Date + 60
    AssertFalse p.EstEchu(), "pas échu si > 30 jours"

    ' Réalisé sans date KO
    p.Statut = "Réalisé"
    AssertFalse p.Validate(), "réalisé sans date KO"
    p.DateRealisee = Date
    AssertTrue p.Validate(), "réalisé avec date OK"

    PrintSummary "CAuditProgramme"
End Sub