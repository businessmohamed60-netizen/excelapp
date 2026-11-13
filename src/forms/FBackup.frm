Attribute VB_Name = "FBackup"
Option Explicit

Private Sub UserForm_Initialize()
    Caption = "Sauvegarde et restauration"
    cmdSauvegarder.Caption = "Sauvegarder (ZIP)"
    cmdRestaurer.Caption = "Restaurer..."
    cmdPurgerJournal.Caption = "Purger le journal"
    cmdClose.Caption = T("btn.close")
    lblDataPath.Caption = "Données : " & DataPath()
    txtJoursPurge.Text = CStr(g_Config.JournallisationJours)
End Sub

Private Sub cmdSauvegarder_Click()
    On Error GoTo ErrHandler
    Dim zipPath As String, fso As Object, folder As Object, file As Object
    Dim tempDir As String

    zipPath = Environ$("USERPROFILE") & "\Desktop\IATF_Backup_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".zip"
    tempDir = Environ$("TEMP") & "\IATFBackup_" & Format$(Now, "hhnnss") & "\"
    MkDir tempDir

    ' Copier tous les XML dans le dossier temporaire
    Dim xmlFile As String
    xmlFile = Dir(DataPath() & "*.xml")
    Do While Len(xmlFile) > 0
        FileCopy DataPath() & xmlFile, tempDir & xmlFile
        xmlFile = Dir
    Loop

    ' Créer le ZIP
    CreateObject("Shell.Application").NameSpace(zipPath).CopyHere _
        CreateObject("Shell.Application").NameSpace(tempDir).Items

    ' Nettoyer
    Kill tempDir & "*.*"
    RmDir tempDir

    lblStatus.Caption = "Sauvegarde créée : " & zipPath
    Exit Sub
ErrHandler:
    LogEx "FBackup", "cmdSauvegarder_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdRestaurer_Click()
    On Error GoTo ErrHandler
    If Not PromptYesNo("Restauration", _
        "La restauration écrase les données actuelles. Continuer ?") Then Exit Sub

    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Sélectionner le fichier de sauvegarde"
    fd.Filters.Clear
    fd.Filters.Add "Archives ZIP", "*.zip"
    If fd.Show <> -1 Then Exit Sub
    Dim zipPath As String: zipPath = fd.SelectedItems(1)

    ' Extraire dans DataPath (écrasement)
    CreateObject("Shell.Application").NameSpace(DataPath()).CopyHere _
        CreateObject("Shell.Application").NameSpace(zipPath).Items

    g_CollMgr.InvalidateAll
    lblStatus.Caption = "Restauration effectuée depuis " & zipPath
    ShowInfo "Restauration", "Les données ont été restaurées. Redémarrez Excel pour recharger."
    Exit Sub
ErrHandler:
    LogEx "FBackup", "cmdRestaurer_Click", Err
    ShowError "Restauration", Err.Description
End Sub

Private Sub cmdPurgerJournal_Click()
    On Error GoTo ErrHandler
    Dim jours As Long: jours = CLng(Val(txtJoursPurge.Text))
    If jours < 30 Then
        ShowError "Purge", "Conserver au moins 30 jours (exigence IATF §7.5.3.2)."
        Exit Sub
    End If
    If Not PromptYesNo("Purge", "Purger le journal au-delà de " & jours & " jours ?") Then Exit Sub
    PurgerJournal jours
    lblStatus.Caption = "Journal purgé (rétention " & jours & " jours)."
    Exit Sub
ErrHandler:
    LogEx "FBackup", "cmdPurgerJournal_Click", Err
    ShowError "Purge", Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub