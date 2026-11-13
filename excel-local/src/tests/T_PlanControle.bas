Attribute VB_Name = "T_PlanControle"
Option Explicit

Public Sub T_PlanControle()
    ResetCounters

    Dim p As CPlanControle
    Set p = New CPlanControle
    p.InitNew
    AssertTrue Len(p.Id) > 0, "Id généré"
    AssertEquals "Brouillon", p.Statut, "statut par défaut"
    AssertEquals "1.0", p.Version, "version par défaut"

    AssertFalse p.Validate(), "validation KO sans produit/processus"
    p.Produit = "Arbre transmission": p.Processus = "Usinage": p.Responsable = "M. Martin"
    AssertTrue p.Validate(), "validation OK"

    p.AssignReference
    AssertTrue Left$(p.Reference, 3) = "PC-", "préfixe PC-"

    Dim l As CLignePlan
    Set l = New CLignePlan
    l.InitNew
    AssertEquals "Normale", l.TypeCaracteristique, "type par défaut"
    AssertEquals 5, l.TailleEchantillon, "échantillon par défaut"

    AssertFalse l.Validate(), "validation KO sans plan"
    l.PlanId = NewGuid()
    AssertFalse l.Validate(), "validation KO sans caractéristique"
    l.Caracteristique = "Diamètre Ø25"
    AssertFalse l.Validate(), "validation KO sans moyen"
    l.MoyenControle = "Pied à coulisse"
    AssertTrue l.Validate(), "validation OK"

    PrintSummary "CPlanControle / CLignePlan"
End Sub