Option Explicit

Public Function CouleurCritique() As Long: CouleurCritique = RGB(255, 220, 220): End Function
Public Function CouleurAvertissement() As Long: CouleurAvertissement = RGB(255, 240, 200): End Function
Public Function CouleurOK() As Long: CouleurOK = RGB(220, 245, 220): End Function
Public Function CouleurNeutre() As Long: CouleurNeutre = RGB(240, 240, 240): End Function

Public Function CouleurPourStatut(ByVal statut As String) As Long
    Select Case LCase$(statut)
        Case "clôturée", "clôturé", "accepté", "validé", "efficace", "ok"
            CouleurPourStatut = CouleurOK()
        Case "en cours", "en préparation", "soumis", "réalisée", "vérifiée"
            CouleurPourStatut = CouleurAvertissement()
        Case "ouverte", "ouvert", "refusé", "annulé", "rejeté"
            CouleurPourStatut = CouleurCritique()
        Case Else
            CouleurPourStatut = CouleurNeutre()
    End Select
End Function

Public Function CouleurPourEcheance(ByVal dateEcheance As Date, ByVal statut As String) As Long
    If LCase$(statut) = "clôturée" Or LCase$(statut) = "clôturé" Then
        CouleurPourEcheance = CouleurNeutre(): Exit Function
    End If
    Dim jours As Long: jours = DateDiff("d", Date, dateEcheance)
    If jours < 0 Then
        CouleurPourEcheance = CouleurCritique()
    ElseIf jours <= 7 Then
        CouleurPourEcheance = CouleurAvertissement()
    Else
        CouleurPourEcheance = CouleurOK()
    End If
End Function

Public Function IndicateurStatut(ByVal statut As String) As String
    Select Case LCase$(statut)
        Case "clôturée", "clôturé", "accepté", "validé", "efficace", "ok"
            IndicateurStatut = "🟢"
        Case "en cours", "en préparation", "soumis", "réalisée", "vérifiée"
            IndicateurStatut = "🟠"
        Case "ouverte", "ouvert", "refusé", "annulé", "rejeté"
            IndicateurStatut = "🔴"
        Case Else
            IndicateurStatut = "⚪"
    End Select
End Function

Public Function IndicateurEcheance(ByVal dateEcheance As Date, ByVal statut As String) As String
    If LCase$(statut) = "clôturée" Or LCase$(statut) = "clôturé" Then
        IndicateurEcheance = "⚪": Exit Function
    End If
    Dim jours As Long: jours = DateDiff("d", Date, dateEcheance)
    If jours < 0 Then
        IndicateurEcheance = "🔴"
    ElseIf jours <= 7 Then
        IndicateurEcheance = "🟠"
    Else
        IndicateurEcheance = "🟢"
    End If
End Function