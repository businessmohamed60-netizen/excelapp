Private Sub cmdCreerCaracteristique_Click()
    If m_CurrentLigne Is Nothing Then Exit Sub
    If m_CurrentLigne.TypeCaracteristique <> "CC" And m_CurrentLigne.TypeCaracteristique <> "SC" Then
        ShowError "Caractéristique", "Le type CC ou SC est requis pour créer une caractéristique spéciale."
        Exit Sub
    End If
    Dim f As FCaracteristique
    Set f = New FCaracteristique
    f.PrefillTypeCar = m_CurrentLigne.TypeCaracteristique
    f.PrefillLibelle = m_CurrentLigne.Caracteristique
    f.PrefillSpecification = m_CurrentLigne.Specification
    f.PrefillPlanId = m_CurrentLigne.PlanId
    f.Show vbModal
    RefreshLignes
End Sub