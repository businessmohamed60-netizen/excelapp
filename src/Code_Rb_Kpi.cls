Public Sub Rb_Kpi(ByVal control As IRibbonControl)
    FKpi.Show vbModeless
End Sub
Public Sub Rb_Programme(ByVal control As IRibbonControl)
    FAuditProgramme.Show vbModeless
End Sub
Public Sub Rb_Documents(ByVal control As IRibbonControl)
    FDocument.Show vbModeless
End Sub
Public Sub Rb_Matrice(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim p As String
    p = GenererMatriceConformite()
    If PromptYesNo("Matrice", "Générée :" & vbCrLf & p & vbCrLf & "Ouvrir ?") Then Workbooks.Open p
    Exit Sub
ErrHandler:
    ShowError "Matrice", Err.Description
End Sub
Public Sub Rb_Revue(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim annee As String
    annee = InputBox("Année de la revue :", "Revue de direction", Year(Date))
    If Len(annee) = 0 Then Exit Sub
    Dim p As String
    p = GenererRevueDirection(CLng(annee))
    If PromptYesNo("Revue", "Générée :" & vbCrLf & p & vbCrLf & "Ouvrir ?") Then Workbooks.Open p
    Exit Sub
ErrHandler:
    ShowError "Revue", Err.Description
End Sub