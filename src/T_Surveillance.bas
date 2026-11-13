Attribute vb_Name = "T_Surveillance"
Option Explicit

Public Sub T_Surveillance()
    ResetCounters

    Dim p As CPlanSurveillance
    Set p = New CPlanSurveillance
    p.InitNew
    AssertTrue Len(p.Id) > 0, "Id généré"
    AssertEquals "Brouillon", p.Statut, "statut par défaut"
    AssertEquals "1.0", p.Version, "version par défaut"

    AssertFalse p.Validate(), "validation KO sans produit"
    p.Produit = "Arbre transmission"
    AssertTrue p.Validate(), "validation OK"

    p.DateCreation = Date: p.AssignReference
    AssertTrue Left$(p.Reference, 3) = "PS-", "préfixe PS-"

    Dim l As CLigneSurveillance
    Set l = New CLigneSurveillance
    l.InitNew
    AssertEquals "Normale", l.TypeCaracteristique, "type par défaut"
    AssertEquals 5, l.TailleEchantillon, "échantillon par défaut"
    AssertEquals "OK", l.StatutCar, "statut par défaut"

    AssertFalse l.Validate(), "validation KO sans plan"
    l.PlanId = NewGuid(): l.Caracteristique = "Diamètre Ø25"
    AssertTrue l.Validate(), "validation OK"

    PrintSummary "CPlanSurveillance"
End Sub