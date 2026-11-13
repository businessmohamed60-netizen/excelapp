Attribute VB_Name = "T_8D"
Option Explicit

Public Sub T_8D()
    ResetCounters

    Dim d As C8D
    Set d = New C8D
    d.InitNew
    AssertTrue Len(d.Id) > 0, "8D Id généré"
    AssertEquals "Ouvert", d.Statut, "statut par défaut"
    AssertEquals "D1", d.EtapeCourante(), "étape initiale D1"

    ' Progression
    d.D2_Date = DateSerial(2026, 3, 1)
    AssertEquals "D2", d.EtapeCourante(), "étape D2"
    d.D4_Date = DateSerial(2026, 3, 10)
    AssertEquals "D4", d.EtapeCourante(), "étape D4"
    d.D8_Date = DateSerial(2026, 4, 1)
    AssertEquals "D8", d.EtapeCourante(), "étape D8"

    ' Validation
    Set d = New C8D: d.InitNew
    AssertFalse d.Validate(), "validation KO sans NC"
    d.NcId = NewGuid()
    AssertFalse d.Validate(), "validation KO sans D2"
    d.D2_Description = "Rupture de stock fournisseur"
    AssertFalse d.Validate(), "validation KO sans cause racine D4"
    d.D4_CauseRacine = "Absence de stock de sécurité"
    AssertTrue d.Validate(), "validation OK"

    ' Chronologie
    d.D2_Date = DateSerial(2026, 3, 1)
    d.D3_Date = DateSerial(2026, 3, 2)
    d.D4_Date = DateSerial(2026, 3, 5)
    d.D5_Date = DateSerial(2026, 3, 10)
    d.D6_Date = DateSerial(2026, 3, 15)
    d.D7_Date = DateSerial(2026, 3, 20)
    d.D8_Date = DateSerial(2026, 4, 1)
    AssertTrue d.ValidateChronologie(), "chronologie OK"

    d.D5_Date = DateSerial(2026, 3, 1)
    AssertFalse d.ValidateChronologie(), "chronologie KO si D5 < D4"

    ' Format numéro
    Set d = New C8D: d.InitNew: d.AssignNumero
    AssertTrue Left$(d.Numero, 3) = "8D-", "préfixe 8D-"
    AssertEquals 10, Len(d.Numero), "longueur 8D-AAAA-NNN"

    PrintSummary "C8D"
End Sub