Option Explicit

' ============================================================
' M_Theme — Palette, typographie, styles du dashboard premium
' ============================================================

' ---- Couleurs principales (RGB) ----
Public Function CouleurPrimaire() As Long: CouleurPrimaire = RGB(31, 78, 121): End Function        ' Bleu IATF
Public Function CouleurSecondaire() As Long: CouleurSecondaire = RGB(46, 116, 181): End Function  ' Bleu clair
Public Function CouleurAccent() As Long: CouleurAccent = RGB(237, 125, 49): End Function          ' Orange
Public Function CouleurFond() As Long: CouleurFond = RGB(248, 249, 250): End Function             ' Gris très clair
Public Function CouleurCarteFond() As Long: CouleurCarteFond = RGB(255, 255, 255): End Function   ' Blanc
Public Function CouleurBordure() As Long: CouleurBordure = RGB(222, 226, 230): End Function        ' Gris bordure

' ---- Couleurs sémantiques ----
Public Function CouleurSucces() As Long: CouleurSucces = RGB(25, 135, 84): End Function           ' Vert
Public Function CouleurAvertissement() As Long: CouleurAvertissement = RGB(255, 193, 7): End Function ' Jaune
Public Function CouleurDanger() As Long: CouleurDanger = RGB(220, 53, 69): End Function           ' Rouge
Public Function CouleurInfo() As Long: CouleurInfo = RGB(13, 202, 240): End Function              ' Cyan
Public Function CouleurNeutre() As Long: CouleurNeutre = RGB(108, 117, 125): End Function         ' Gris

' ---- Couleurs de texte ----
Public Function TextePrincipal() As Long: TextePrincipal = RGB(33, 37, 41): End Function
Public Function TexteSecondaire() As Long: TexteSecondaire = RGB(108, 117, 125): End Function
Public Function TexteInverse() As Long: TexteInverse = RGB(255, 255, 255): End Function

' ---- Typographie (noms de polices) ----
Public Function PoliceTitre() As String: PoliceTitre = "Segoe UI Semibold": End Function
Public Function PoliceCorps() As String: PoliceCorps = "Segoe UI": End Function
Public Function PoliceChiffres() As String: PoliceChiffres = "Segoe UI Light": End Function

' ---- Tailles ----
Public Function TailleTitre() As Single: TailleTitre = 18: End Function
Public Function TailleKpiChiffre() As Single: TailleKpiChiffre = 28: End Function
Public Function TailleKpiLabel() As Single: TailleKpiLabel = 9: End Function
Public Function TailleCorps() As Single: TailleCorps = 10: End Function
Public Function TaillePetit() As Single: TaillePetit = 8: End Function

' ---- Icônes sémantiques ----
Public Function IconeKpi(ByVal type As String) As String
    Select Case type
        Case "clients": IconeKpi = "👥"
        Case "audits": IconeKpi = "🔍"
        Case "audits_en_cours": IconeKpi = "⏳"
        Case "nc": IconeKpi = "⚠️"
        Case "nc_majeures": IconeKpi = "🚨"
        Case "capa_retard": IconeKpi = "⏰"
        Case "apqp": IconeKpi = "📋"
        Case "ppap": IconeKpi = "📦"
        Case "amdec_h": IconeKpi = "🔥"
        Case "taux_cloture": IconeKpi = "✅"
        Case "alertes": IconeKpi = "🔔"
        Case "score_global": IconeKpi = "⭐"
        Case Else: IconeKpi = "•"
    End Select
End Function

' ---- Couleur selon le niveau d'alerte d'une valeur ----
Public Function CouleurPourNiveau(ByVal niveau As String) As Long
    Select Case LCase$(niveau)
        Case "critique", "danger", "h", "rejeté": CouleurPourNiveau = CouleurDanger()
        Case "avertissement", "warn", "m": CouleurPourNiveau = CouleurAvertissement()
        Case "ok", "succes", "l", "acceptable": CouleurPourNiveau = CouleurSucces()
        Case Else: CouleurPourNiveau = CouleurNeutre()
    End Select
End Function