Attribute VB_Name = "T_Document"
Option Explicit

Public Sub T_Document()
    ResetCounters
    Dim d As CDocument
    Set d = New CDocument
    d.InitNew
    AssertTrue Len(d.Id) > 0, "Id généré"
    AssertEquals "Brouillon", d.Statut, "statut par défaut"
    AssertEquals "1.0", d.Version, "version par défaut"
    AssertEquals "Procédure", d.TypeDoc, "type par défaut"
    AssertTrue d.DateRevuePrevue > Date, "revue prévue future"

    ' Validation
    AssertFalse d.Validate(), "validation KO sans code"
    d.Code = "PR-QUA-001": d.Titre = "Maîtrise des documents"
    AssertTrue d.Validate(), "validation OK brouillon"

    ' Approuvé sans date KO
    d.Statut = "Approuvé"
    AssertFalse d.Validate(), "approuvé sans date KO"
    d.DateApprobation = Date
    AssertTrue d.Validate(), "approuvé avec date OK"

    ' RevueEchue
    d.DateRevuePrevue = Date - 1
    AssertTrue d.RevueEchue(), "revue échue"
    d.DateRevuePrevue = Date + 30
    AssertFalse d.RevueEchue(), "revue non échue"
    d.Statut = "Brouillon"
    AssertFalse d.RevueEchue(), "brouillon jamais échu"

    ' DisplayName
    AssertTrue InStr(d.DisplayName(), "PR-QUA-001") > 0, "DisplayName contient le code"

    PrintSummary "CDocument"
End Sub