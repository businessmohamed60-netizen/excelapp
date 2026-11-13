Attribute VB_Name = "T_MoyenMesure"
Option Explicit

Public Sub T_MoyenMesure()
    ResetCounters

    Dim m As CMoyenMesure
    Set m = New CMoyenMesure
    m.InitNew
    AssertTrue Len(m.Id) > 0, "Id généré"
    AssertEquals "En service", m.Statut, "statut par défaut"
    AssertEquals 12, m.PeriodiciteMois, "périodicité 12"

    AssertFalse m.Validate(), "validation KO sans désignation"
    m.Designation = "Pied à coulisse 150 mm"
    m.TypeMoyen = "Pied à coulisse"
    AssertTrue m.Validate(), "validation OK"

    ' Code
    m.AssignCode
    AssertTrue Left$(m.Code, 3) = "MM-", "préfixe MM-"

    ' Recalcul échéance
    m.DernierEtalonnage = DateSerial(2025, 6, 15)
    m.RecalculerProchaineEcheance
    AssertEquals DateSerial(2026, 6, 15), m.ProchainEtalonnage, "prochaine échéance +12 mois"

    ' EstEtalonnageEchu
    m.ProchainEtalonnage = Date - 10
    AssertTrue m.EstEtalonnageEchu(), "étalonnage échu"
    m.ProchainEtalonnage = Date + 10
    AssertFalse m.EstEtalonnageEchu(), "étalonnage à jour"
    AssertTrue m.EtalonnageProche(), "étalonnage proche (<=30j)"

    PrintSummary "CMoyenMesure"
End Sub