Option Explicit

' ============================================================
' M_KpiFormatter — Application des styles premium aux KPI
' ============================================================

' Applique un KPI à un groupe de 4 contrôles (carte)
'   - lblChiffre : la valeur numérique
'   - lblLabel : le libellé
'   - lblIcone : l'icône
'   - frameCarte : la bordure de la carte (Frame ou Label avec BackColor)
Public Sub AppliquerKpi(ByVal kpi As Object, ByVal lblChiffre As Object, _
                        ByVal lblLabel As Object, ByVal lblIcone As Object, _
                        ByVal frameCarte As Object)
    On Error Resume Next

    ' Valeur
    lblChiffre.Caption = FormaterValeurKpi(kpi("Valeur"))
    lblChiffre.Font.Name = PoliceChiffres()
    lblChiffre.Font.Size = TailleKpiChiffre()
    lblChiffre.Font.Bold = True
    lblChiffre.ForeColor = CouleurPourNiveau(kpi("Niveau"))

    ' Libellé
    lblLabel.Caption = kpi("Label")
    lblLabel.Font.Name = PoliceKpiLabel()
    lblLabel.Font.Size = TailleKpiLabel()
    lblLabel.ForeColor = TexteSecondaire()

    ' Icône
    lblIcone.Caption = kpi("Icone")
    lblIcone.Font.Size = 20

    ' Carte (bordure + fond)
    frameCarte.BackColor = CouleurCarteFond()
    frameCarte.BorderStyle = 1
    frameCarte.BorderColor = CouleurPourNiveau(kpi("Niveau"))
    frameCarte.SpecialEffect = 0

    Err.Clear
    On Error GoTo 0
End Sub

Private Function FormaterValeurKpi(ByVal v As Variant) As String
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        If d = Int(d) Then FormaterValeurKpi = Format$(d, "#,##0") Else FormaterValeurKpi = Format$(d, "0.0")
    Else
        FormaterValeurKpi = CStr(v)
    End If
End Function

Private Function PoliceKpiLabel() As String: PoliceKpiLabel = "Segoe UI": End Function