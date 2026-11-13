Attribute VB_Name = "T_Caracteristique"
Option Explicit

Public Sub T_Caracteristique()
    ResetCounters

    Dim c As CCaracteristique
    Set c = New CCaracteristique
    c.InitNew
    AssertTrue Len(c.Id) > 0, "Id généré"
    AssertEquals "Active", c.Statut, "statut par défaut"
    AssertEquals "SC", c.TypeCar, "type par défaut"

    AssertFalse c.Validate(), "validation KO sans libellé"
    c.Libelle = "Diamètre critique Ø25 ±0.05"
    AssertTrue c.Validate(), "validation OK"
    c.TypeCar = "XX"
    AssertFalse c.Validate(), "type invalide KO"
    c.TypeCar = "CC"
    AssertTrue c.Validate(), "CC valide"

    ' DisplayName
    c.Code = "CC-001"
    AssertTrue InStr(c.DisplayName(), "CC-001") > 0, "DisplayName contient le code"

    PrintSummary "CCaracteristique"
End Sub