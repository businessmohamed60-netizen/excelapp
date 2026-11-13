Attribute VB_Name = "FAmdec"
Option Explicit

Private m_Amdecs As Object
Private m_Lignes As Object
Private m_FilteredLignes As Object
Private m_CurrentAmdec As CAMDEC
Private m_CurrentLigne As CLigneAMDEC

Private Sub UserForm_Initialize()
    Caption = "AMDEC — Analyse des modes de défaillance"
    cmdNewAmdec.Caption = T("btn.new")
    cmdSaveAmdec.Caption = T("btn.save")
    cmdDeleteAmdec.Caption = T("btn.delete")
    cmdNewLigne.Caption = "Nouvelle ligne"
    cmdSaveLigne.Caption = T("btn.save")
    cmdDeleteLigne.Caption = T("btn.delete")
    cmdRapport.Caption = "Générer rapport"
    cmdClose.Caption = T("btn.close")

    InitCombos
    InitNotes
    RefreshAmdecs
End Sub

Private Sub InitCombos()
    Dim i As Long
    With cboTypeAmdec: .Clear
        .AddItem "Processus": .AddItem "Produit": .AddItem "Moyens": .AddItem "Fournisseur"
    End With
    With cboStatutAmdec: .Clear
        .AddItem "Brouillon": .AddItem "En revue": .AddItem "Approuvée": .AddItem "Obsolète"
    End With
    With cboStatutAction: .Clear
        .AddItem "Ouverte": .AddItem "En cours": .AddItem "Clôturée"
    End With
    With cboProcessus: .Clear
        .AddItem "Production": .AddItem "Qualité": .AddItem "Logistique"
        .AddItem "Achats": .AddItem "Maintenance": .AddItem "RH"
        .AddItem "Direction": .AddItem "Commercial": .AddItem "Autre"
    End With
    For i = 1 To 10
        cboGravite.AddItem CStr(i): cboOccurrence.AddItem CStr(i): cboDetection.AddItem CStr(i)
        cboGraviteApres.AddItem CStr(i): cboOccurrenceApres.AddItem CStr(i): cboDetectionApres.AddItem CStr(i)
    Next i
End Sub

Private Sub InitNotes()
    ' À appeler une seule fois après création des contrôles (ou dans Initialize)
    ' Rien ici : les notes sont saisies par l'utilisateur
End Sub

Private Sub RefreshAmdecs()
    Set m_Amdecs = g_CollMgr.LoadCollection("amdec")
    lstAmdecs.Clear
    With lstAmdecs: .ColumnCount = 3: .ColumnWidths = "140;70;70": End With
    Dim k As Variant
    For Each k In m_Amdecs.Keys
        lstAmdecs.AddItem m_Amdecs(k).Reference
        lstAmdecs.List(lstAmdecs.ListCount - 1, 1) = m_Amdecs(k).TypeAmdec
        lstAmdecs.List(lstAmdecs.ListCount - 1, 2) = m_Amdecs(k).Statut
    Next k
    lblStatus.Caption = m_Amdecs.Count & " AMDEC"
End Sub

Private Sub lstAmdecs_Click()
    Dim idx As Long: idx = lstAmdecs.ListIndex
    If idx < 0 Then Exit Sub
    Dim k As Variant, i As Long
    i = 0
    For Each k In m_Amdecs.Keys
        If i = idx Then
            Set m_CurrentAmdec = m_Amdecs(k)
            LoadAmdecForm
            RefreshLignes
            Exit Sub
        End If
        i = i + 1
    Next k
End Sub

Private Sub LoadAmdecForm()
    If m_CurrentAmdec Is Nothing Then Exit Sub
    cboTypeAmdec.Text = m_CurrentAmdec.TypeAmdec
    cboProcessus.Text = m_CurrentAmdec.Processus
    txtProduit.Text = m_CurrentAmdec.Produit
    txtResponsable.Text = m_CurrentAmdec.Responsable
    txtEquipe.Text = m_CurrentAmdec.Equipe
    txtVersion.Text = m_CurrentAmdec.Version
    cboStatutAmdec.Text = m_CurrentAmdec.Statut
    txtApprobateur.Text = m_CurrentAmdec.Approbateur
    If IsDate(m_CurrentAmdec.DateApprobation) And m_CurrentAmdec.DateApprobation > 0 Then
        txtDateApprobation.Text = FormatIsoDate(m_CurrentAmdec.DateApprobation)
    Else
        txtDateApprobation.Text = ""
    End If
End Sub

Private Sub RefreshLignes()
    If m_CurrentAmdec Is Nothing Then
        lstLignes.Clear: Exit Sub
    End If
    Set m_Lignes = g_CollMgr.LoadCollection("ligneamdec")
    Set m_FilteredLignes = CreateObject("Scripting.Dictionary")
    lstLignes.Clear
    With lstLignes
        .ColumnCount = 9
        .ColumnWidths = "30;90;140;25;25;25;40;60;70"
    End With
    Dim k As Variant, ligne As CLigneAMDEC
    For Each k In m_Lignes.Keys
        If m_Lignes(k).AmdecId = m_CurrentAmdec.Id Then
            Set ligne = m_Lignes(k)
            m_FilteredLignes.Add k, k
            lstLignes.AddItem CStr(ligne.NumeroLigne)
            lstLignes.List(lstLignes.ListCount - 1, 1) = ligne.EtapeProcessus
            lstLignes.List(lstLignes.ListCount - 1, 2) = ligne.ModeDefaillance
            lstLignes.List(lstLignes.ListCount - 1, 3) = CStr(ligne.Gravite)
            lstLignes.List(lstLignes.ListCount - 1, 4) = CStr(ligne.Occurrence)
            lstLignes.List(lstLignes.ListCount - 1, 5) = CStr(ligne.Detection)
            lstLignes.List(lstLignes.ListCount - 1, 6) = CStr(ligne.Rpn)
            lstLignes.List(lstLignes.ListCount - 1, 7) = ligne.Criticite
            lstLignes.List(lstLignes.ListCount - 1, 8) = ligne.StatutAction
        End If
    Next k
    lblStatus.Caption = m_FilteredLignes.Count & " ligne(s) — RPN max : " & _
        StatistiquesAmdec(m_CurrentAmdec.Id).RpnMax
End Sub

Private Sub lstLignes_Click()
    Dim idx As Long: idx = lstLignes.ListIndex
    If idx < 0 Then Exit Sub
    Dim k As Variant, i As Long
    i = 0
    For Each k In m_FilteredLignes.Keys
        If i = idx Then
            Set m_CurrentLigne = m_Lignes(k)
            LoadLigneForm
            Exit Sub
        End If
        i = i + 1
    Next k
End Sub

Private Sub LoadLigneForm()
    If m_CurrentLigne Is Nothing Then Exit Sub
    txtEtape.Text = m_CurrentLigne.EtapeProcessus
    txtFonction.Text = m_CurrentLigne.Fonction
    txtMode.Text = m_CurrentLigne.ModeDefaillance
    txtEffet.Text = m_CurrentLigne.Effet
    cboGravite.Text = CStr(m_CurrentLigne.Gravite)
    txtCause.Text = m_CurrentLigne.Cause
    cboOccurrence.Text = CStr(m_CurrentLigne.Occurrence)
    txtControlesPrev.Text = m_CurrentLigne.ControlesPreventifs
    cboDetection.Text = CStr(m_CurrentLigne.Detection)
    txtControlesDet.Text = m_CurrentLigne.ControlesDetectifs
    lblRpn.Caption = "RPN = " & m_CurrentLigne.Rpn
    lblCriticite.Caption = m_CurrentLigne.Criticite
    txtAction.Text = m_CurrentLigne.ActionRecommandee
    txtRespAction.Text = m_CurrentLigne.ResponsableAction
    If IsDate(m_CurrentLigne.DateEcheanceAction) And m_CurrentLigne.DateEcheanceAction > 0 Then
        txtDateEcheance.Text = FormatIsoDate(m_CurrentLigne.DateEcheanceAction)
    Else
        txtDateEcheance.Text = ""
    End If
    cboStatutAction.Text = m_CurrentLigne.StatutAction
    If m_CurrentLigne.GraviteApres > 0 Then
        cboGraviteApres.Text = CStr(m_CurrentLigne.GraviteApres)
        cboOccurrenceApres.Text = CStr(m_CurrentLigne.OccurrenceApres)
        cboDetectionApres.Text = CStr(m_CurrentLigne.DetectionApres)
        lblRpnApres.Caption = "RPN après = " & m_CurrentLigne.RpnApres
    Else
        cboGraviteApres.ListIndex = -1: cboOccurrenceApres.ListIndex = -1
        cboDetectionApres.ListIndex = -1: lblRpnApres.Caption = ""
    End If
End Sub

Private Sub ClearLigneForm()
    txtEtape.Text = "": txtFonction.Text = "": txtMode.Text = "": txtEffet.Text = ""
    cboGravite.Text = "5": txtCause.Text = "": cboOccurrence.Text = "5"
    txtControlesPrev.Text = "": cboDetection.Text = "5": txtControlesDet.Text = ""
    lblRpn.Caption = "RPN = 125": lblCriticite.Caption = "Élevée"
    txtAction.Text = "": txtRespAction.Text = "": txtDateEcheance.Text = ""
    cboStatutAction.Text = "Ouverte"
    cboGraviteApres.ListIndex = -1: cboOccurrenceApres.ListIndex = -1
    cboDetectionApres.ListIndex = -1: lblRpnApres.Caption = ""
End Sub

Private Sub cboGravite_Click(): RecalculRpn: End Sub
Private Sub cboOccurrence_Click(): RecalculRpn: End Sub
Private Sub cboDetection_Click(): RecalculRpn: End Sub
Private Sub cboGraviteApres_Click(): RecalculRpn: End Sub
Private Sub cboOccurrenceApres_Click(): RecalculRpn: End Sub
Private Sub cboDetectionApres_Click(): RecalculRpn: End Sub

Private Sub RecalculRpn()
    Dim g As Long, o As Long, d As Long
    g = CLng(Val(cboGravite.Text)): o = CLng(Val(cboOccurrence.Text)): d = CLng(Val(cboDetection.Text))
    If g >= 1 And o >= 1 And d >= 1 Then
        Dim l As New CLigneAMDEC
        l.Gravite = g: l.Occurrence = o: l.Detection = d
        l.RecalculerRpn
        lblRpn.Caption = "RPN = " & l.Rpn
        lblCriticite.Caption = l.Criticite
        If Len(cboGraviteApres.Text) > 0 And Len(cboOccurrenceApres.Text) > 0 And Len(cboDetectionApres.Text) > 0 Then
            Dim l2 As New CLigneAMDEC
            l2.GraviteApres = CLng(Val(cboGraviteApres.Text))
            l2.OccurrenceApres = CLng(Val(cboOccurrenceApres.Text))
            l2.DetectionApres = CLng(Val(cboDetectionApres.Text))
            l2.RecalculerRpn
            lblRpnApres.Caption = "RPN après = " & l2.RpnApres
        End If
    End If
End Sub

Private Sub cmdNewAmdec_Click()
    Set m_CurrentAmdec = New CAMDEC
    m_CurrentAmdec.InitNew
    ClearAmdecForm
    lblStatus.Caption = "Nouvelle AMDEC — référence attribuée à l'enregistrement"
End Sub

Private Sub ClearAmdecForm()
    cboTypeAmdec.Text = "Processus": cboProcessus.ListIndex = -1
    txtProduit.Text = "": txtResponsable.Text = "": txtEquipe.Text = ""
    txtVersion.Text = "1.0": cboStatutAmdec.Text = "Brouillon"
    txtApprobateur.Text = "": txtDateApprobation.Text = ""
End Sub

Private Sub cmdSaveAmdec_Click()
    On Error GoTo ErrHandler
    If m_CurrentAmdec Is Nothing Then Exit Sub
    m_CurrentAmdec.TypeAmdec = cboTypeAmdec.Text
    m_CurrentAmdec.Processus = cboProcessus.Text
    m_CurrentAmdec.Produit = txtProduit.Text
    m_CurrentAmdec.Responsable = txtResponsable.Text
    m_CurrentAmdec.Equipe = txtEquipe.Text
    m_CurrentAmdec.Version = txtVersion.Text
    m_CurrentAmdec.Statut = cboStatutAmdec.Text
    m_CurrentAmdec.Approbateur = txtApprobateur.Text
    If Len(txtDateApprobation.Text) > 0 Then
        m_CurrentAmdec.DateApprobation = ParseIsoDate(txtDateApprobation.Text)
    End If
    If Len(m_CurrentAmdec.Reference) = 0 Then m_CurrentAmdec.AssignReference()
    If Not m_CurrentAmdec.Validate() Then
        ShowError "Validation", "Processus ou produit, et responsable, sont requis."
        Exit Sub
    End If
    g_CollMgr.Save m_CurrentAmdec, "amdec"
    g_CollMgr.InvalidateCache "amdec"
    RefreshAmdecs
    lblStatus.Caption = "AMDEC " & m_CurrentAmdec.Reference & " enregistrée."
    Exit Sub
ErrHandler:
    LogEx "FAmdec", "cmdSaveAmdec_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdDeleteAmdec_Click()
    If m_CurrentAmdec Is Nothing Then Exit Sub
    If Not PromptYesNo("Suppression", "Supprimer l'AMDEC et toutes ses lignes ?") Then Exit Sub
    ' Supprimer les lignes rattachées
    Dim lignes As Object, k As Variant
    Set lignes = g_CollMgr.LoadCollection("ligneamdec")
    For Each k In lignes.Keys
        If lignes(k).AmdecId = m_CurrentAmdec.Id Then
            g_CollMgr.DeleteById k, "ligneamdec"
        End If
    Next k
    g_CollMgr.DeleteById m_CurrentAmdec.Id, "amdec"
    g_CollMgr.InvalidateCache "amdec": g_CollMgr.InvalidateCache "ligneamdec"
    Set m_CurrentAmdec = Nothing
    ClearAmdecForm
    RefreshAmdecs
    lstLignes.Clear
End Sub

Private Sub cmdNewLigne_Click()
    If m_CurrentAmdec Is Nothing Then
        ShowError "Ligne AMDEC", "Sélectionnez ou créez une AMDEC d'abord."
        Exit Sub
    End If
    If Len(m_CurrentAmdec.Reference) = 0 Then
        ShowError "Ligne AMDEC", "Enregistrez l'AMDEC avant d'ajouter des lignes."
        Exit Sub
    End If
    Set m_CurrentLigne = New CLigneAMDEC
    m_CurrentLigne.InitNew
    m_CurrentLigne.AmdecId = m_CurrentAmdec.Id
    m_CurrentLigne.NumeroLigne = m_FilteredLignes.Count + 1
    ClearLigneForm
    lblStatus.Caption = "Nouvelle ligne AMDEC n° " & m_CurrentLigne.NumeroLigne
End Sub

Private Sub cmdSaveLigne_Click()
    On Error GoTo ErrHandler
    If m_CurrentLigne Is Nothing Then Exit Sub
    If Not m_CurrentAmdec Is Nothing Then m_CurrentLigne.AmdecId = m_CurrentAmdec.Id

    m_CurrentLigne.EtapeProcessus = txtEtape.Text
    m_CurrentLigne.Fonction = txtFonction.Text
    m_CurrentLigne.ModeDefaillance = txtMode.Text
    m_CurrentLigne.Effet = txtEffet.Text
    m_CurrentLigne.Gravite = CLng(Val(cboGravite.Text))
    m_CurrentLigne.Cause = txtCause.Text
    m_CurrentLigne.Occurrence = CLng(Val(cboOccurrence.Text))
    m_CurrentLigne.ControlesPreventifs = txtControlesPrev.Text
    m_CurrentLigne.Detection = CLng(Val(cboDetection.Text))
    m_CurrentLigne.ControlesDetectifs = txtControlesDet.Text
    m_CurrentLigne.ActionRecommandee = txtAction.Text
    m_CurrentLigne.ResponsableAction = txtRespAction.Text
    If Len(txtDateEcheance.Text) > 0 Then
        m_CurrentLigne.DateEcheanceAction = ParseIsoDate(txtDateEcheance.Text)
    End If
    m_CurrentLigne.StatutAction = cboStatutAction.Text
    If Len(cboGraviteApres.Text) > 0 Then
        m_CurrentLigne.GraviteApres = CLng(Val(cboGraviteApres.Text))
        m_CurrentLigne.OccurrenceApres = CLng(Val(cboOccurrenceApres.Text))
        m_CurrentLigne.DetectionApres = CLng(Val(cboDetectionApres.Text))
    End If
    m_CurrentLigne.RecalculerRpn

    If Not m_CurrentLigne.Validate() Then
        ShowError "Validation", "Mode de défaillance et notes (1..10) requis."
        Exit Sub
    End If
    g_CollMgr.Save m_CurrentLigne, "ligneamdec"
    g_CollMgr.InvalidateCache "ligneamdec"
    RefreshLignes
    lblStatus.Caption = "Ligne " & m_CurrentLigne.NumeroLigne & " enregistrée (RPN = " & m_CurrentLigne.Rpn & ")."
    Exit Sub
ErrHandler:
    LogEx "FAmdec", "cmdSaveLigne_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdDeleteLigne_Click()
    If m_CurrentLigne Is Nothing Then Exit Sub
    If Not PromptYesNo("Suppression", "Supprimer cette ligne AMDEC ?") Then Exit Sub
    g_CollMgr.DeleteById m_CurrentLigne.Id, "ligneamdec"
    g_CollMgr.InvalidateCache "ligneamdec"
    Set m_CurrentLigne = Nothing
    ClearLigneForm
    RefreshLignes
End Sub

Private Sub cmdRapport_Click()
    On Error GoTo ErrHandler
    If m_CurrentAmdec Is Nothing Then Exit Sub
    Dim path As String
    path = GenererRapportAmdec(m_CurrentAmdec)
    If PromptYesNo("Rapport AMDEC", "Généré :" & vbCrLf & path & vbCrLf & "Ouvrir ?") Then
        Workbooks.Open path
    End If
    Exit Sub
ErrHandler:
    LogEx "FAmdec", "cmdRapport_Click", Err
    ShowError "Rapport", Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub