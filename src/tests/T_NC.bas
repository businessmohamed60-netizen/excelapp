Attribute VB_Name = "T_NC"
Option Explicit

Public Sub T_NonConformite()
    ResetCounters

    Dim nc As CNonConformite
    Set nc = New CNonConformite
    nc.InitNew
    AssertTrue Len(nc.Id) > 0, "NC Id généré"
    AssertEquals "Ouverte", nc.Statut, "statut par défaut"
    AssertEquals "Mineur", nc.Classification, "classification par défaut"
    AssertEquals 2, nc.Gravite, "gravité par défaut"
    AssertTrue IsDate(nc.DateDetection), "date de détection définie"

    ' Validation KO — sans client
    AssertFalse nc.Validate(), "validation KO sans client"

    ' Validation OK
    nc.ClientId = NewGuid()
    nc.Description = "Défaut dimensionnel sur axe X"
    AssertTrue nc.Validate(), "validation OK avec client + description"

    ' Clôture sans date KO
    nc.Statut = "Clôturée"
    AssertFalse nc.Validate(), "validation KO si clôturée sans date"
    nc.DateCloture = Date
    AssertTrue nc.Validate(), "validation OK si clôturée avec date"

    ' Format numéro
    nc.AssignNumero
    AssertTrue Left$(nc.Numero, 3) = "NC-", "préfixe NC-"
    AssertEquals 9, Len(nc.Numero), "longueur NC-AAAA-NNN"

    PrintSummary "CNonConformite"
End Sub