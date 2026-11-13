Private Sub SaveFromForm()
    m_Current.D1_Champion = txtD1_Champion.Text
    m_Current.D1_Leader = txtD1_Leader.Text
    m_Current.D1_Membres = txtD1_Membres.Text

    m_Current.D2_Description = txtD2_Description.Text
    If Len(txtD2_Date.Text) > 0 Then m_Current.D2_Date = ParseIsoDate(txtD2_Date.Text)
    m_Current.D2_5W2H = txtD2_5W2H.Text

    m_Current.D3_Actions = txtD3_Actions.Text
    If Len(txtD3_Date.Text) > 0 Then m_Current.D3_Date = ParseIsoDate(txtD3_Date.Text)
    m_Current.D3_Verification = txtD3_Verification.Text

    m_Current.D4_Ishikawa = txtD4_Ishikawa.Text
    m_Current.D4_5Pourquoi = txtD4_5Pourquoi.Text
    m_Current.D4_CauseRacine = txtD4_CauseRacine.Text
    m_Current.D4_CauseEchappee = txtD4_CauseEchappee.Text
    If Len(txtD4_Date.Text) > 0 Then m_Current.D4_Date = ParseIsoDate(txtD4_Date.Text)

    m_Current.D5_Actions = txtD5_Actions.Text
    m_Current.D5_Responsable = txtD5_Responsable.Text
    If Len(txtD5_Date.Text) > 0 Then m_Current.D5_Date = ParseIsoDate(txtD5_Date.Text)

    m_Current.D6_Actions = txtD6_Actions.Text
    If Len(txtD6_Date.Text) > 0 Then m_Current.D6_Date = ParseIsoDate(txtD6_Date.Text)
    m_Current.D6_Validation = txtD6_Validation.Text

    m_Current.D7_Actions = txtD7_Actions.Text
    If Len(txtD7_Date.Text) > 0 Then m_Current.D7_Date = ParseIsoDate(txtD7_Date.Text)

    If Len(txtD8_Date.Text) > 0 Then m_Current.D8_Date = ParseIsoDate(txtD8_Date.Text)
    m_Current.D8_Bilan = txtD8_Bilan.Text

    ' Statut dérivé de l'étape courante
    If IsDate(m_Current.D8_Date) And m_Current.D8_Date > 0 Then
        m_Current.Statut = "Clôturé"
    ElseIf Len(m_Current.D2_Description) > 0 Then
        m_Current.Statut = "En cours"
    Else
        m_Current.Statut = "Ouvert"
    End If

    lblEtape.Caption = "Étape courante : " & m_Current.EtapeCourante()
End Sub