Attribute VB_Name = "T_Capabilite"
Option Explicit

Public Sub T_Capabilite()
    ResetCounters

    Dim e As CEtudeCapabilite
    Set e = New CEtudeCapabilite
    e.InitNew
    AssertTrue Len(e.Id) > 0, "Id généré"
    AssertEquals "Brouillon", e.Statut, "statut par défaut"

    AssertFalse e.Validate(), "validation KO sans caractéristique"
    e.Caracteristique = "Diamètre Ø25"
    e.USL = 25.1: e.LSL = 24.9
    AssertTrue e.Validate(), "validation OK"

    ' Calcul sur valeurs maîtrisées : moyenne 25, écart-type ~0.0167
    Dim v As Variant
    v = Array(24.98, 25.0, 25.02, 24.99, 25.01, 25.0, 25.0, 25.01, 24.99, 25.0)
    CalculerCapabilite e, v
    AssertEquals 10, e.NbValeurs, "10 valeurs"
    AssertTrue e.Moyenne > 24.99 And e.Moyenne < 25.01, "moyenne ≈ 25"
    AssertTrue e.Cp > 0, "Cp calculé"
    AssertTrue e.Cpk > 0, "Cpk calculé"
    AssertTrue e.Cpk <= e.Cp, "Cpk <= Cp (décentrage)"

    ' Interprétation
    e.Cpk = 1.8: AssertEquals "Capable", e.Interpretation(), "Cpk 1.8 capable"
    e.Cpk = 1.4: AssertEquals "Acceptable", e.Interpretation(), "Cpk 1.4 acceptable"
    e.Cpk = 1.1: AssertEquals "Insuffisant", e.Interpretation(), "Cpk 1.1 insuffisant"

    ' Référence
    e.DateEtude = Date: e.AssignReference
    AssertTrue Left$(e.Reference, 4) = "CAP-", "préfixe CAP-"

    PrintSummary "CEtudeCapabilite"
End Sub