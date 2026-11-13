Attribute VB_Name = "T_Capa"
Option Explicit

Public Sub T_Capa()
    ResetCounters

    Dim c As CCapa
    Set c = New CCapa
    c.InitNew
    AssertTrue Len(c.Id) > 0, "CAPA Id généré"
    AssertEquals "Ouverte", c.Statut, "statut par défaut"
    AssertEquals "Corrective", c.TypeAction, "type par défaut"
    AssertEquals "À vérifier", c.Efficacite, "efficacité par défaut"
    AssertTrue c.DateEcheance > c.DateOuverture, "échéance > ouverture"

    ' EstEnRetard
    c.DateEcheance = Date - 5: c.Statut = "En cours"
    AssertTrue c.EstEnRetard(), "en retard si échéance passée"
    c.DateEcheance = Date + 5
    AssertFalse c.EstEnRetard(), "pas en retard si échéance future"
    c.DateEcheance = Date - 5: c.Statut = "Clôturée"
    AssertFalse c.EstEnRetard(), "pas en retard si clôturée"

    ' Validation
    AssertFalse c.Validate(), "validation KO sans NC"
    c.NcId = NewGuid()
    AssertFalse c.Validate(), "validation KO sans description"
    c.Description = "Mettre en place un Poka-Yoke"
    AssertTrue c.Validate(), "validation OK"
    c.DateEcheance = c.DateOuverture - 1
    AssertFalse c.Validate(), "validation KO si échéance < ouverture"

    ' Format numéro
    c.DateOuverture = Date: c.DateEcheance = Date + 30
    c.AssignNumero
    AssertTrue Left$(c.Numero, 5) = "CAPA-", "préfixe CAPA-"
    AssertEquals 11, Len(c.Numero), "longueur CAPA-AAAA-NNN"

    PrintSummary "CCapa"
End Sub