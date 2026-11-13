Attribute VB_Name = "FParametres"
Option Explicit

Private Sub UserForm_Initialize()
    Caption = "Paramètres"
    cmdSave.Caption = T("btn.save")
    cmdClose.Caption = T("btn.close")
    cmdTesterStockage.Caption = "Tester"
    ChargerParametres
End Sub

Private Sub ChargerParametres()
    With cboStorageMode
        .Clear
        .AddItem "Local": .AddItem "OneDrive": .AddItem "SharePoint": .AddItem "Réseau"
    End With
    cboStorageMode.Text = g_Config.StorageMode
    txtStoragePath.Text = g_Config.StoragePath
    txtStoragePath.Enabled = (g_Config.StorageMode = "SharePoint" Or g_Config.StorageMode = "Réseau")
    chkNotificationsAuto.Value = g_Config.NotificationsAuto
    txtDestinataire.Text = g_Config.NotificationsDestinataire
    txtHorizonJours.Text = CStr(g_Config.NotificationsHorizonJours)
    chkJournallisation.Value = g_Config.JournallisationActive
    txtJoursConservation.Text = CStr(g_Config.JournallisationJours)
    chkSignatureRequise.Value = g_Config.SignatureRequise
End Sub

Private Sub cboStorageMode_Change()
    txtStoragePath.Enabled = (cboStorageMode.Text = "SharePoint" Or cboStorageMode.Text = "Réseau")
End Sub

Private Sub cmdTesterStockage_Click()
    Dim mode As String, path As String
    mode = cboStorageMode.Text
    path = txtStoragePath.Text
    If mode = "SharePoint" Or mode = "Réseau" Then
        If Dir(path, vbDirectory) = "" Then
            ShowError "Stockage", "Chemin inaccessible : " & path
            Exit Sub
        End If
    End If
    ' Test d'écriture dans le chemin cible
    Dim testFile As String
    testFile = path & "_test_" & Format$(Now, "yyyymmddhhnnss") & ".tmp"
    On Error GoTo ErrHandler
    Dim f As Integer: f = FreeFile
    Open testFile For Output As #f
    Print #f, "test"
    Close #f
    Kill testFile
    lblStatus.Caption = "Stockage accessible : " & path
    Exit Sub
ErrHandler:
    ShowError "Stockage", "Écriture impossible dans " & path
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    g_Config.StorageMode = cboStorageMode.Text
    g_Config.StoragePath = txtStoragePath.Text
    g_Config.NotificationsAuto = chkNotificationsAuto.Value
    g_Config.NotificationsDestinataire = txtDestinataire.Text
    g_Config.NotificationsHorizonJours = CLng(Val(txtHorizonJours.Text))
    g_Config.JournallisationActive = chkJournallisation.Value
    g_Config.JournallisationJours = CLng(Val(txtJoursConservation.Text))
    g_Config.SignatureRequise = chkSignatureRequise.Value
    SauvegarderConfig g_Config
    lblStatus.Caption = "Paramètres enregistrés."
    Exit Sub
ErrHandler:
    LogEx "FParametres", "cmdSave_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub