Attribute VB_Name = "FDocument"
Option Explicit

Private m_Docs As Object
Private m_Filtered As Object
Private m_Current As CDocument

Private Sub UserForm_Initialize()
    Caption = "Gestion documentaire"
    cmdNew.Caption = T("btn.new")
    cmdSave.Caption = T("btn.save")
    cmdDelete.Caption = T("btn.delete")
    cmdClose.Caption = T("btn.close")
    cmdNouvelleVersion.Caption = "Nouvelle version"
    InitCombos
    RefreshList
End Sub

Private Sub InitCombos()
    With cboTypeDoc: .Clear
        .AddItem "Procédure": .AddItem "Instruction": .AddItem "Formulaire"
        .AddItem "Enregistrement": .AddItem "Politique": .AddItem "Manuel"
    End With
    With cboTypeDocEdit: .Clear
        .AddItem "Procédure": .AddItem "Instruction": .AddItem "Formulaire"
        .AddItem "Enregistrement": .AddItem "Politique": .AddItem "Manuel"
    End With
    With cboStatut: .Clear
        .AddItem "(Tous)": .AddItem "Brouillon": .AddItem "En revue"
        .AddItem "Approuvé": .AddItem "Obsolète"
    End With
    With cboStatutEdit: .Clear
        .AddItem "Brouillon": .AddItem "En revue": .AddItem "Approuvé": .AddItem "Obsolète"
    End With
    With cboProcessus: .Clear
        .AddItem "Production": .AddItem "Qualité": .AddItem "Logistique"
        .AddItem "Achats": .AddItem "Maintenance": .AddItem "RH"
        .AddItem "Direction": .AddItem "Commercial": .AddItem "Autre"
    End With
End Sub

Private Sub RefreshList()
    Set m_Docs = g_CollMgr.LoadCollection("document")
    Set m_Filtered = CreateObject("Scripting.Dictionary")
    lstDocuments.Clear
    With lstDocuments
        .ColumnCount = 6
        .ColumnWidths = "80;200;50;70;80;50"
    End With
    Dim k As Variant, filtreStatut As String, filtreType As String, rech As String
    filtreStatut = cboStatut.Text
    filtreType = cboTypeDoc.Text
    rech = LCase$(Trim$(txtRecherche.Text))
    For Each k In m_Docs.Keys
        If (filtreStatut = "(Tous)" Or m_Docs(k).Statut = filtreStatut) And _
           (Len(filtreType) = 0 Or m_Docs(k).TypeDoc = filtreType) And _
           (Len(rech) = 0 Or InStr(LCase$(m_Docs(k).Titre), rech) > 0 Or _
            InStr(LCase$(m_Docs(k).Code), rech) > 0) Then
            m_Filtered.Add k, k
            lstDocuments.AddItem m_Docs(k).Code
            lstDocuments.List(lstDocuments.ListCount - 1, 1) = m_Docs(k).Titre
            lstDocuments.List(lstDocuments.ListCount - 1, 2) = m_Docs(k).Version
            lstDocuments.List(lstDocuments.ListCount - 1, 3) = m_Docs(k).Statut
            lstDocuments.List(lstDocuments.ListCount - 1, 4) = FormatIsoDate(m_Docs(k).DateRevuePrevue)
            lstDocuments.List(lstDocuments.ListCount - 1, 5) = IIf(m_Docs(k).RevueEchue(), "OUI", "")
        End If
    Next k
    lblStatus.Caption = m_Filtered.Count & " document(s)"
End Sub

Private Sub cboStatut_Change(): RefreshList: End Sub
Private Sub cboTypeDoc_Change(): RefreshList: End Sub
Private Sub txtRecherche_Change(): RefreshList: End Sub

Private Sub lstDocuments_Click()
    Dim idx As Long: idx = lstDocuments.ListIndex
    If idx < 0 Then Exit Sub
    Dim k As Variant, i As Long
    i = 0
    For Each k In m_Filtered.Keys
        If i = idx Then
            Set m_Current = m_Docs(k)
            LoadForm
            Exit Sub
        End If
        i = i + 1
    Next k
End Sub

Private Sub LoadForm()
    If m_Current Is Nothing Then Exit Sub
    txtCode.Text = m_Current.Code
    txtTitre.Text = m_Current.Titre
    cboTypeDocEdit.Text = m_Current.TypeDoc
    txtVersion.Text = m_Current.Version
    cboStatutEdit.Text = m_Current.Statut
    txtRedacteur.Text = m_Current.Redacteur
    txtApprobateur.Text = m_Current.Approbateur
    If IsDate(m_Current.DateApprobation) And m_Current.DateApprobation > 0 Then
        txtDateApprobation.Text = FormatIsoDate(m_Current.DateApprobation)
    Else
        txtDateApprobation.Text = ""
    End If
    txtDateRevuePrevue.Text = FormatIsoDate(m_Current.DateRevuePrevue)
    cboProcessus.Text = m_Current.Processus
    txtChemin.Text = m_Current.CheminFichier
    txtCommentaire.Text = m_Current.Commentaire
End Sub

Private Sub ClearForm()
    txtCode.Text = ""
    txtTitre.Text = ""
    cboTypeDocEdit.Text = "Procédure"
    txtVersion.Text = "1.0"
    cboStatutEdit.Text = "Brouillon"
    txtRedacteur.Text = ""
    txtApprobateur.Text = ""
    txtDateApprobation.Text = ""
    txtDateRevuePrevue.Text = FormatIsoDate(Date + 365)
    cboProcessus.ListIndex = -1
    txtChemin.Text = ""
    txtCommentaire.Text = ""
End Sub

Private Sub cmdNew_Click()
    Set m_Current = New CDocument
    m_Current.InitNew
    ClearForm
    lblStatus.Caption = "Nouveau document"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    If m_Current Is Nothing Then Exit Sub
    m_Current.Code = txtCode.Text
    m_Current.Titre = txtTitre.Text
    m_Current.TypeDoc = cboTypeDocEdit.Text
    m_Current.Version = txtVersion.Text
    m_Current.Statut = cboStatutEdit.Text
    m_Current.Redacteur = txtRedacteur.Text
    m_Current.Approbateur = txtApprobateur.Text
    If Len(txtDateApprobation.Text) > 0 Then
        m_Current.DateApprobation = ParseIsoDate(txtDateApprobation.Text)
    End If
    m_Current.DateRevuePrevue = ParseIsoDate(txtDateRevuePrevue.Text)
    m_Current.Processus = cboProcessus.Text
    m_Current.CheminFichier = txtChemin.Text
    m_Current.Commentaire = txtCommentaire.Text
    If Not m_Current.Validate() Then
        ShowError "Validation", "Code, titre et date d'approbation (si approuvé) requis."
        Exit Sub
    End If
    g_CollMgr.Save m_Current, "document"
    g_CollMgr.InvalidateCache "document"
    RefreshList
    lblStatus.Caption = "Document " & m_Current.Code & " enregistré."
    Exit Sub
ErrHandler:
    LogEx "FDocument", "cmdSave_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdDelete_Click()
    If m_Current Is Nothing Then Exit Sub
    If Not PromptYesNo("Suppression", T("msg.confirm_delete")) Then Exit Sub
    g_CollMgr.DeleteById m_Current.Id, "document"
    g_CollMgr.InvalidateCache "document"
    Set m_Current = Nothing
    ClearForm
    RefreshList
End Sub

' Incrémente la version mineure (1.0 → 1.1) et remet le statut à "En revue"
Private Sub cmdNouvelleVersion_Click()
    If m_Current Is Nothing Then Exit Sub
    Dim v As String, parts() As String, maj As Long, min As Long
    v = m_Current.Version
    parts = Split(v, ".")
    If UBound(parts) >= 1 Then
        maj = CLng(Val(parts(0))): min = CLng(Val(parts(1)))
    Else
        maj = CLng(Val(v)): min = 0
    End If
    min = min + 1
    If min >= 10 Then maj = maj + 1: min = 0
    m_Current.Version = maj & "." & min
    m_Current.Statut = "En revue"
    g_CollMgr.Save m_Current, "document"
    g_CollMgr.InvalidateCache "document"
    LoadForm
    RefreshList
    lblStatus.Caption = "Nouvelle version " & m_Current.Version & " — statut En revue"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub