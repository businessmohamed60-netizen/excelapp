Option Explicit

Public Sub RunQualiteTests()
    ResetCounters
    Dim nc As CNonConformite: Set nc = New CNonConformite: nc.InitNew
    AssertEquals "Ouverte", nc.Statut, "statut NC par défaut"
    AssertFalse nc.Validate(), "NC KO sans client"
    nc.ClientId = NewGuid(): nc.Description = "Test"
    AssertTrue nc.Validate(), "NC valide"
    Dim capa As CCapa: Set capa = New CCapa: capa.InitNew
    capa.NcId = NewGuid(): capa.Description = "Test CAPA"
    AssertTrue capa.Validate(), "CAPA valide"
    capa.DateEcheance = Date - 5
    AssertTrue capa.EstEnRetard(), "CAPA en retard"
    Dim d As C8D: Set d = New C8D: d.InitNew
    d.NcId = NewGuid(): d.D2_Description = "Desc"
    AssertFalse d.Validate(), "8D KO sans D4"
    d.D4_CauseRacine = "Cause"
    AssertTrue d.Validate(), "8D valide"
    PrintSummary "Qualité"
End Sub