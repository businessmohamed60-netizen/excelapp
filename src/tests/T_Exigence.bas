Attribute VB_Name = "T_Exigence"
Option Explicit

Public Sub T_Exigence()
    ResetCounters
    Dim e As CExigence
    Set e = New CExigence
    e.InitNew
    AssertTrue Len(e.Id) > 0, "Id généré"
    AssertTrue e.Applicable, "applicable par défaut"
    AssertEquals "Non évaluée", e.Conformite, "conformité par défaut"
    AssertEquals "IATF 16949", e.Source, "source par défaut"

    ' Validation
    AssertFalse e.Validate(), "validation KO sans référence"
    e.Reference = "8.5.1.1": e.Intitule = "Fiche d'instruction"
    AssertTrue e.Validate(), "validation OK"

    ' EstNonConforme
    e.Conformite = "Non conforme"
    AssertTrue e.EstNonConforme(), "non conforme détecté"
    e.Conformite = "Partielle"
    AssertTrue e.EstNonConforme(), "partielle détectée"
    e.Conformite = "Conforme"
    AssertFalse e.EstNonConforme(), "conforme non détecté"
    e.Conformite = "Non conforme": e.Applicable = False
    AssertFalse e.EstNonConforme(), "non applicable ignoré"

    PrintSummary "CExigence"
End Sub