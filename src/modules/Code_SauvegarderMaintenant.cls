Option Explicit

Public Function SauvegarderMaintenant() As Boolean
    On Error GoTo ErrHandler
    Dim fso As Object, src As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim horodatage As String, dossierCible As String
    horodatage = Format$(Now, "yyyymmdd_hhnnss")
    dossierCible = BackupPath() & "backup_" & horodatage & "\"
    MkDir dossierCible
    If fso.FolderExists(DataPath()) Then
        Set src = fso.GetFolder(DataPath())
        Dim f As Object
        For Each f In src.Files
            If LCase$(fso.GetExtensionName(f.Name)) = "xml" Then fso.CopyFile f.Path, dossierCible & f.Name
        Next f
    End If
    Log LogInfo, "M_Backup", "Sauvegarde : " & dossierCible
    SauvegarderMaintenant = True
    NettoyerAnciennesSauvegardes 30
    Exit Function
ErrHandler:
    LogEx "M_Backup", "SauvegarderMaintenant", Err
    SauvegarderMaintenant = False
End Function

Public Sub NettoyerAnciennesSauvegardes(ByVal joursConserves As Long)
    On Error Resume Next
    Dim fso As Object, folder As Object, sub_ As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(BackupPath()) Then Exit Sub
    Set folder = fso.GetFolder(BackupPath())
    For Each sub_ In folder.SubFolders
        If DateDiff("d", sub_.DateCreated, Now) > joursConserves Then
            fso.DeleteFolder sub_.Path, True
            Log LogInfo, "M_Backup", "Ancienne sauvegarde supprimée : " & sub_.Name
        End If
    Next sub_
End Sub

Public Sub SauvegardeAutomatiqueSiNecessaire()
    On Error Resume Next
    Dim fso As Object, folder As Object, sub_ As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(BackupPath()) Then SauvegarderMaintenant: Exit Sub
    Set folder = fso.GetFolder(BackupPath())
    Dim derniereDate As Date: derniereDate = 0
    For Each sub_ In folder.SubFolders
        If sub_.DateCreated > derniereDate Then derniereDate = sub_.DateCreated
    Next sub_
    If derniereDate = 0 Or DateDiff("d", derniereDate, Now) >= 1 Then SauvegarderMaintenant
End Sub