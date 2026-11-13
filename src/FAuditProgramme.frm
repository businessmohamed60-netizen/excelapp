Attribute VB_Name = "FAuditProgramme"
Option Explicit

Private m_Programmes As Object
Private m_Clients As Object
Private m_ClientIds As Object
Private m_Current As CAuditProgramme

Private Sub UserForm_Initialize()
    Caption = "Programme annuel d'audits"
    cmdNew.Caption = T("btn.new")
    cmdSave.Caption = T("btn.save")
    cmdDelete.Caption = T("btn.delete")
    cmdClose.Caption = T("btn.close")
    cmdGenererAnnee.Caption = "Générer l'année"
    InitCombos
    LoadClients
    InitAnnees
    RefreshList
End Sub

Private Sub InitCombos()
    With cboTypeAudit
        .Clear
        .AddItem "Processus": .AddItem "Système": .AddItem "Produit": .AddItem "Fournisseur"
    End With
    With cboRisque
        .Clear
        .AddItem "Faible": .AddItem "Moyen": .AddItem "Élevé": .AddItem "Critique"
    End With
    With cboFrequence
        .Clear
        .AddItem "3": .AddItem "6": .AddItem "12": .AddItem "24"
    End With
    With cboStatut
        .Clear
        .AddItem "Planifié": .AddItem "Réalisé": .AddItem "Reporté": .AddItem "Annulé"
    End With
    With cboProcessus
        .Clear
        .AddItem "Production": .AddItem "Qualité": .AddItem "Logistique"
        .AddItem "Achats": .AddItem "Maintenance": .AddItem "RH"
        .AddItem "Direction": .AddItem "Commercial": .AddItem "Autre"
    End With
End Sub

Private Sub LoadClients()
    Set m_Clients = g_CollMgr.LoadCollection("clients")
    Set m_ClientIds = CreateObject("Scripting.Dictionary")
    cboClient.Clear
    Dim k As Variant
    For Each k In m_Clients.Keys
        cboClient.AddItem m_Clients(k).DisplayName
        m_ClientIds.Add m_Clients(k).DisplayName, k
    Next k
End Sub

Private Sub InitAnnees()
    Dim y As Long, current As Long
    current = Year(Date)
    cboAnnee.Clear
    For y = current + 1 To current - 2 Step -1
        cboAnnee.AddItem CStr(y)
    Next y
    cboAnnee.ListIndex = 1
End Sub

Private Sub RefreshList()
    Set m_Programmes = g_CollMgr.LoadCollection("auditprogramme")
    lstProgramme.Clear
    With lstProgramme
        .ColumnCount = 6
        .ColumnWidths = "50;100;80;70;80;80"
    End With
    Dim k As Variant, anneeFiltre As Long
    anneeFiltre = CLng(cboAnnee.Text)
    For Each k In m_Programmes.Keys
        If m_Programmes(k).Annee = anneeFiltre Then
            lstProgramme.AddItem CStr(m_Programmes(k).Annee)
            lstProgramme.List(lstProgramme.ListCount - 1, 1) = m_Programmes(k).Processus
            lstProgramme.List(lstProgramme.ListCount - 1, 2) = m_Programmes(k).TypeAudit
            lstProgramme.List(lstProgramme.ListCount - 1, 3) = m_Programmes(k).Risque
            lstProgramme.List(lstProgramme.ListCount - 1, 4) = FormatIsoDate(m_Programmes(k).DatePrevue)
            lstProgramme.List(lstProgramme.ListCount - 1, 5) = m_Programmes(k).Statut
        End If
    Next k
    lblStatus.Caption = lstProgramme.ListCount & " audit(s) programmé(s) pour " & anneeFiltre
End Sub

Private Sub cboAnnee_Change()
    If cboAnnee.ListIndex >= 0 Then RefreshList
End Sub

Private Sub lstProgramme_Click()
    Dim idx As Long: idx = lstProgramme.ListIndex
    If idx < 0 Then Exit Sub
    Dim k As Variant, i As Long, anneeFiltre As Long
    anneeFiltre = CLng(cboAnnee.Text)
    i = 0
    For Each k In m_Programmes.Keys
        If m_Programmes(k).Annee = anneeFiltre Then
            If i = idx Then
                Set m_Current = m_Programmes(k)
                LoadForm
                Exit Sub
            End If
            i = i + 1
        End If
    Next k
End Sub

Private Sub LoadForm()
    If m_Current Is Nothing Then Exit Sub
    cboAnnee.Text = CStr(m_Current.Annee)
    If m_Clients.Exists(m_Current.ClientId) Then
        cboClient.Text = m_Clients(m_Current.ClientId).DisplayName
    Else
        cboClient.ListIndex = -1
    End If
    cboProcessus.Text = m_Current.Processus
    cboTypeAudit.Text = m_Current.TypeAudit
    cboRisque.Text = m_Current.Risque
    cboFrequence.Text = CStr(m_Current.FrequenceMois)
    txtDatePrevue.Text = FormatIsoDate(m_Current.DatePrevue)
    txtAuditeur.Text = m_Current.Auditeur
    cboStatut.Text = m_Current.Statut
    txtCommentaire.Text = m_Current.Commentaire
End Sub

Private Sub ClearForm()
    cboClient.ListIndex = -1
    cboProcessus.ListIndex = -1
    cboTypeAudit.Text = "Processus"
    cboRisque.Text = "Moyen"
    cboFrequence.Text = "12"
    txtDatePrevue.Text = FormatIsoDate(DateSerial(CLng(cboAnnee.Text), 1, 15))
    txtAuditeur.Text = ""
    cboStatut.Text = "Planifié"
    txtCommentaire.Text = ""
End Sub

Private Sub cmdNew_Click()
    Set m_Current = New CAuditProgramme
    m_Current.InitNew
    m_Current.Annee = CLng(cboAnnee.Text)
    ClearForm
    lblStatus.Caption = "Nouvelle ligne de programme"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    If m_Current Is Nothing Then Exit Sub
    m_Current.Annee = CLng(cboAnnee.Text)
    If m_ClientIds.Exists(cboClient.Text) Then
        m_Current.ClientId = m_ClientIds(cboClient.Text)
    End If
    m_Current.Processus = cboProcessus.Text
    m_Current.TypeAudit = cboTypeAudit.Text
    m_Current.Risque = cboRisque.Text
    m_Current.FrequenceMois = CLng(cboFrequence.Text)
    m_Current.DatePrevue = ParseIsoDate(txtDatePrevue.Text)
    m_Current.Auditeur = txtAuditeur.Text
    m_Current.Statut = cboStatut.Text
    m_Current.Commentaire = txtCommentaire.Text
    If Not m_Current.Validate() Then
        ShowError "Validation", "Processus, année et date prévue sont obligatoires."
        Exit Sub
    End If
    g_CollMgr.Save m_Current, "auditprogramme"
    g_CollMgr.InvalidateCache "auditprogramme"
    RefreshList
    lblStatus.Caption = "Ligne de programme enregistrée."
    Exit Sub
ErrHandler:
    LogEx "FAuditProgramme", "cmdSave_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdDelete_Click()
    If m_Current Is Nothing Then Exit Sub
    If Not PromptYesNo("Suppression", T("msg.confirm_delete")) Then Exit Sub
    g_CollMgr.DeleteById m_Current.Id, "auditprogramme"
    g_CollMgr.InvalidateCache "auditprogramme"
    Set m_Current = Nothing
    ClearForm
    RefreshList
End Sub

' Génère 12 lignes de programme (une par mois) pour l'année sélectionnée
Private Sub cmdGenererAnnee_Click()
    Dim annee As Long: annee = CLng(cboAnnee.Text)
    If Not PromptYesNo("Génération", "Générer un programme mensuel pour " & annee & " ?") Then Exit Sub
    Dim mois As Long, p As CAuditProgramme
    For mois = 1 To 12
        Set p = New CAuditProgramme
        p.InitNew
        p.Annee = annee
        p.Processus = "À définir"
        p.DatePrevue = DateSerial(annee, mois, 15)
        p.Statut = "Planifié"
        p.Commentaire = "Ligne générée automatiquement — à compléter"
        g_CollMgr.Save p, "auditprogramme"
    Next mois
    g_CollMgr.InvalidateCache "auditprogramme"
    RefreshList
    lblStatus.Caption = "12 lignes générées pour " & annee
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub