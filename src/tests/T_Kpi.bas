Attribute VB_Name = "T_Kpi"
Option Explicit

Public Sub T_Kpi()
    ResetCounters
    Dim k As TKpi

    ' Sens Haut
    k = EvaluerKpi("Test", 95, 90, "Haut", "%")
    AssertEquals "OK", k.Statut, "95 >= 90 OK"
    k = EvaluerKpi("Test", 75, 90, "Haut", "%")
    AssertEquals "Alerte", k.Statut, "75 entre 72 et 90 alerte"
    k = EvaluerKpi("Test", 50, 90, "Haut", "%")
    AssertEquals "Critique", k.Statut, "50 < 72 critique"

    ' Sens Bas
    k = EvaluerKpi("Test", 1, 2, "Bas", "")
    AssertEquals "OK", k.Statut, "1 <= 2 OK"
    k = EvaluerKpi("Test", 3, 2, "Bas", "")
    AssertEquals "Alerte", k.Statut, "3 entre 2 et 2.4 alerte"
    k = EvaluerKpi("Test", 5, 2, "Bas", "")
    AssertEquals "Critique", k.Statut, "5 > 2.4 critique"

    ' Tendances
    AssertEquals "→", CalculerTendance(10, 10, "Haut"), "stable"
    AssertEquals "↑", CalculerTendance(12, 10, "Haut"), "hausse favorable (Haut)"
    AssertEquals "↓", CalculerTendance(8, 10, "Haut"), "baisse défavorable (Haut)"
    AssertEquals "↑", CalculerTendance(8, 10, "Bas"), "baisse favorable (Bas)"
    AssertEquals "↓", CalculerTendance(12, 10, "Bas"), "hausse défavorable (Bas)"

    PrintSummary "Kpi"
End Sub