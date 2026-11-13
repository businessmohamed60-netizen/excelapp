Attribute VB_Name = "FNotifications"
Option Explicit

Private m_Echeances As Collection

Private Sub UserForm_Initialize()
    Caption = "Centre de notifications"
    cmdEnvoyerEmail.Caption = "Envoyer par email"
    cmdRafraichir.Caption = "Actualiser"
    cmdClose.Caption = T("btn.close")
    With cboFiltreNiveau
        .Clear
        .AddItem "(Tous)": .AddItem "Critique": .AddItem "Urgent"
        .AddItem "Attention": .AddItem "Info"
    End With
    With cboFiltreCategorie
        .Clear
        .AddItem "(Toutes)": .AddItem "CAPA": .AddItem "Étalonnage"
        .AddItem "Revue documentaire": .AddItem "Action AMDEC": .AddItem "Audit programmé"
    End With
    cboFiltreNiveau.ListIndex = 0
    cboFiltreCategorie.ListIndex = 0
    txtHorizon.Text = "30"
    RefreshList
End Sub

Private Sub RefreshList()
    Set m_Echeances = CollecterEcheances(CLng(Val(txtHorizon.Text)))
    lstEcheances.Clear
    With lstEcheances
        .ColumnCount = 7
        .ColumnWidths = "60;80;60;180;70;40;90"
    End With
    Dim i As Long, e As TEcheance, n As Long
    For i = 1 To m_Echeances.Count
        e = m_Echeances(i)
        If (cboFiltreNiveau.Text = "(Tous)" Or e.Niveau = cboFiltreNiveau.Text) And _
           (cboFiltreCategorie.Text = "(Toutes)" Or e.Categorie = cboFiltreCategorie.Text) Then
            lstEcheances.AddItem e.Niveau
            lstEcheances.List(lstEcheances.ListCount - 1, 1) = e.Categorie
            lstEcheances.List(lstEcheances.ListCount - 1, 2) = e.Reference
            lstEcheances.List(lstEcheances.ListCount - 1, 3) = e.Description
            lstEcheances.List(lstEcheances.ListCount - 1, 4) = Format$(e.DateEcheance, "yyyy-mm-dd")
            lstEcheances.List(lstEcheances.ListCount - 1, 5) = CStr(e.JoursRestants)
            lstEcheances.List(lstEcheances.ListCount - 1, 6) = e.Responsable
            n = n + 1
        End If
    Next i
    lblStatus.Caption = n & " échéance(s) affichée(s) sur " & m_Echeances.Count
End Sub

Private Sub cboFiltreNiveau_Change(): RefreshList: End Sub
Private Sub cboFiltreCategorie_Change(): RefreshList: End Sub
Private Sub txtHorizon_Change(): RefreshList: End Sub
Private Sub cmdRafraichir_Click(): g_CollMgr.InvalidateAll: RefreshList: End Sub

Private Sub cmdEnvoyerEmail_Click()
    If m_Echeances.Count = 0 Then
        ShowInfo "Notifications", "Aucune échéance à notifier."
        Exit Sub
    End If
    If EnvoyerNotification(g_Config.NotificationsDestinataire, CLng(Val(txtHorizon.Text))) Then
        lblStatus.Caption = "Email préparé dans Outlook."
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub