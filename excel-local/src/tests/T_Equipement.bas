Attribute vb_Name = "T_Equipement"
Option Explicit

Public Sub T_Equipement()
    ResetCounters

    Dim e As CEquipement
    Set e = New CEquipement
    e.InitNew
    AssertTrue Len(e.Id) > 0, "Id généré"
    AssertEquals "En service", e.Statut, "statut par défaut"
    AssertEquals "Machine", e.TypeEquipement, "type par défaut"

    ' Validation
    AssertFalse e.Validate(), "validation KO sans désignation"
    e.Designation = "Presse hydraulique 200T"
    AssertTrue e.Validate(), "validation OK"

    ' Code
    e.AssignCode
    AssertTrue Left$(e.Code, 3) = "EQ-", "préfixe EQ-"

    ' TRS
    e.TempsOuverturePlanifie = 100
    e.TempsArretPlanifie = 10
    e.TempsArretNonPlanifie = 10
    e.PiecesProduites = 1000
    e.PiecesBonnes = 950
    AssertEquals 80#, e.Disponibilite(), "disponibilité (100-10-10)/100 = 80%"
    AssertEquals 95#, e.Qualite(), "qualité 950/1000 = 95%"
    AssertTrue e.TRS() > 0, "TRS calculé"

    ' Maintenance échue
    e.ProchaineMaintenance = Date - 3
    AssertTrue e.EstMaintenanceEchue(), "maintenance échue"

    PrintSummary "CEquipement"
End Sub