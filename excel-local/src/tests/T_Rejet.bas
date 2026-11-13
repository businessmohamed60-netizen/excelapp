Attribute vb_Name = "T_Rejet"
Option Explicit

Public Sub T_Rejet()
    ResetCounters

    Dim r As CRejetInterne
    Set r = New CRejetInterne
    r.InitNew
    AssertTrue Len(r.Id) > 0, "Id généré"
    AssertEquals "Ouvert", r.Statut, "statut par défaut"
    AssertEquals "Dimensionnel", r.TypeRejet, "type par défaut"

    ' Validation
    r.QuantiteRejetee = 10: r.QuantiteTotale = 0
    AssertFalse r.Validate(), "validation KO si qté totale = 0"
    r.QuantiteTotale = 5
    AssertFalse r.Validate(), "validation KO si rejet > total"
    r.QuantiteTotale = 100
    AssertTrue r.Validate(), "validation OK"

    ' Numéro
    r.DateRejet = DateSerial(2026, 3, 15)
    r.AssignNumero
    AssertTrue Left$(r.Numero, 3) = "RI-", "préfixe RI-"

    ' Coût
    r.CoutUnitaire = 12.5
    r.RecalculerCout
    AssertEquals 125#, r.CoutTotal, "coût total 10*12.5"

    ' Taux de rebut
    AssertEquals 10#, r.TauxRebut(), "taux rebut 10/100 = 10%"

    PrintSummary "CRejetInterne"
End Sub