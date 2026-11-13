Attribute VB_Name = "T_Msa"
Option Explicit

Public Sub T_Msa()
    ResetCounters

    Dim e As CEtudeMsa
    Set e = New CEtudeMsa
    e.InitNew
    AssertTrue Len(e.Id) > 0, "Id généré"
    AssertEquals "Brouillon", e.Statut, "statut par défaut"
    AssertEquals 3, e.NbOperateurs, "3 opérateurs par défaut"
    AssertEquals 10, e.NbPieces, "10 pièces par défaut"
    AssertEquals 3, e.NbEssais, "3 essais par défaut"

    AssertFalse e.Validate(), "validation KO sans moyen"
    e.MoyenId = NewGuid()
    AssertFalse e.Validate(), "validation KO sans caractéristique"
    e.Caracteristique = "Diamètre Ø25"
    AssertTrue e.Validate(), "validation OK"

    ' Interprétation MSA
    e.PourcentGrRTolerance = 8
    AssertEquals "Acceptable", e.Interpretation(), "%GRR 8% acceptable"
    e.PourcentGrRTolerance = 20
    AssertEquals "Conditionnel", e.Interpretation(), "%GRR 20% conditionnel"
    e.PourcentGrRTolerance = 35
    AssertEquals "Rejeté", e.Interpretation(), "%GRR 35% rejeté"

    ' Référence
    e.DateEtude = Date: e.AssignReference
    AssertTrue Left$(e.Reference, 4) = "MSA-", "préfixe MSA-"

    PrintSummary "CEtudeMsa"
End Sub