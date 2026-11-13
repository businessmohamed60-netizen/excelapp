Attribute VB_Name = "T_Carte"
Option Explicit

Public Sub T_Carte()
    ResetCounters

    Dim c As CCarteControle
    Set c = New CCarteControle
    c.InitNew
    AssertTrue Len(c.Id) > 0, "Id généré"
    AssertEquals "Brouillon", c.Statut, "statut par défaut"
    AssertEquals "Xbar-R", c.TypeCarte, "type par défaut"
    AssertEquals 5, c.SousGroupeTaille, "taille 5 par défaut"

    AssertFalse c.Validate(), "validation KO sans caractéristique"
    c.Caracteristique = "Diamètre Ø25"
    AssertTrue c.Validate(), "validation OK"

    ' Calcul carte avec 5 sous-groupes de 5 valeurs
    Dim sg(1 To 5, 1 To 5) As Variant, i As Long, j As Long
    For i = 1 To 5
        For j = 1 To 5
            sg(i, j) = 25 + (i * 0.01) + (j * 0.001)
        Next j
    Next i
    CalculerCarteXbarR c, sg
    AssertEquals 5, c.NbSousGroupes, "5 sous-groupes"
    AssertEquals 5, c.SousGroupeTaille, "taille 5"
    AssertTrue c.LigneCentre > 25, "ligne centre > 25"
    AssertTrue c.LimiteSup > c.LigneCentre, "UCL > centre"
    AssertTrue c.LimiteInf < c.LigneCentre, "LCL < centre"
    AssertTrue c.LimiteSupR > 0, "UCLR > 0"

    ' Référence
    c.DateEtude = Date: c.AssignReference
    AssertTrue Left$(c.Reference, 4) = "SPC-", "préfixe SPC-"

    PrintSummary "CCarteControle"
End Sub