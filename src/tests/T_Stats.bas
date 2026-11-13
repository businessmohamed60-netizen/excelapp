Attribute VB_Name = "T_Stats"
Option Explicit

Public Sub T_Stats()
    ResetCounters

    ' Moyenne
    Dim a As Variant
    a = Array(1, 2, 3, 4, 5)
    AssertEquals 3#, Moyenne(a), "moyenne 1..5 = 3"

    ' Écart-type global
    a = Array(2, 4, 4, 4, 5, 5, 7, 9)
    AssertEquals 2.13809, Round(EcartTypeGlobal(a), 5), "écart-type global = 2.13809"

    ' Étendue
    a = Array(3, 7, 1, 9, 5)
    AssertEquals 8#, Etendue(a), "étendue max-min = 8"

    ' Constantes
    AssertEquals 1.128, ConstanteD2(2), "d2(2)"
    AssertEquals 2.326, ConstanteD2(5), "d2(5)"
    AssertEquals 0.577, ConstanteA2(5), "A2(5)"
    AssertEquals 2.114, ConstanteD4(5), "D4(5)"
    AssertEquals 0#, ConstanteD3(5), "D3(5) = 0"

    ' Règles de Nelson — règle 1 (hors limites)
    Dim v As Variant, alertes As Collection
    v = Array(5, 5, 5, 15, 5, 5, 5)
    Set alertes = AppliquerReglesNelson(v, 5, 12, -2)
    AssertTrue alertes.Count >= 1, "règle 1 détecte un point hors limites"

    ' Règle 2 — 9 points du même côté
    v = Array(6, 6, 6, 6, 6, 6, 6, 6, 6)
    Set alertes = AppliquerReglesNelson(v, 5, 12, -2)
    AssertTrue alertes.Count >= 1, "règle 2 détecte 9 points du même côté"

    ' Règle 3 — 6 points en tendance
    v = Array(1, 2, 3, 4, 5, 6, 7)
    Set alertes = AppliquerReglesNelson(v, 5, 12, -2)
    AssertTrue alertes.Count >= 1, "règle 3 détecte une tendance"

    PrintSummary "M_Stats"
End Sub