Attribute VB_Name = "FMsa"
Option Explicit

Private m_Current As CEtudeMsa
Private m_Mesures As Object

Private Sub UserForm_Initialize()
    Caption = "Étude MSA / R&R"
    cmdSave.Caption = T("btn.save")
    cmdRapport.Caption = "Générer rapport"
    cmdClose.Caption = T("btn.close")
    cmdCalculer.Caption = "Calculer GRR"
    cmdAddMesure.Caption = "Ajouter mesure"
    cmdClearMesures.Caption = "Effacer mesures"
    InitCombos
    Set m_Current = New CEtudeMsa
    m_Current.InitNew
    ClearForm
End Sub

Private Sub InitCombos()
    Dim i As Long
    cboOperateur.Clear: cboPiece.Clear: cboEssai.Clear
    For i = 1 To 10: cboOperateur.AddItem CStr(i): cboPiece.AddItem CStr(i): cboEssai.AddItem CStr(i): Next i
End Sub

Private Sub ClearForm()
    txtReference.Text = "": cboMoyen.Clear: txtCaracteristique.Text = ""
    txtNbOperateurs.Text = "3": txtNbPieces.Text = "10": txtNbEssais.Text = "3"
    txtUSL.Text = "": txtLSL.Text = "": txtTolerance.Text = "": txtResponsable.Text = ""
    lstMesures.Clear
    lblEV.Caption = "": lblAV.Caption = "": lblGrR.Caption = ""
    lblPctGrR.Caption = "": lblPctGrRTol.Caption = "": lblNdc.Caption = ""
    lblInterpretation.Caption = ""
End Sub

Private Sub cmdAddMesure_Click()
    On Error GoTo ErrHandler
    If m_Current Is Nothing Then Exit Sub
    Dim o As Long, p As Long, e As Long, v As Double
    o = CLng(Val(cboOperateur.Text)): p = CLng(Val(cboPiece.Text)): e = CLng(Val(cboEssai.Text))
    v = Val(Replace(txtValeur.Text, ",", "."))
    If o < 1 Or p < 1 Or e < 1 Then
        ShowError "Mesure", "Opérateur, pièce et essai doivent être renseignés."
        Exit Sub
    End If

    Dim ligne As CLigneMsa
    Set ligne = New CLigneMsa
    ligne.InitNew
    ligne.EtudeId = m_Current.Id
    ligne.Operateur = o: ligne.Piece = p: ligne.Essai = e: ligne.Valeur = v
    g_CollMgr.Save ligne, "lignemsa"
    g_CollMgr.InvalidateCache "lignemsa"
    RefreshMesures
    txtValeur.Text = ""
    Exit Sub
ErrHandler:
    LogEx "FMsa", "cmdAddMesure_Click", Err
    ShowError "Mesure", Err.Description
End Sub

Private Sub RefreshMesures()
    lstMesures.Clear
    With lstMesures: .ColumnCount = 4: .ColumnWidths = "70;60;60;80": End With
    Dim lignes As Object, k As Variant
    Set lignes = g_CollMgr.LoadCollection("lignemsa")
    For Each k In lignes.Keys
        If lignes(k).EtudeId = m_Current.Id Then
            lstMesures.AddItem CStr(lignes(k).Operateur)
            lstMesures.List(lstMesures.ListCount - 1, 1) = CStr(lignes(k).Piece)
            lstMesures.List(lstMesures.ListCount - 1, 2) = CStr(lignes(k).Essai)
            lstMesures.List(lstMesures.ListCount - 1, 3) = Format$(lignes(k).Valeur, "0.000")
        End If
    Next k
    lblStatus.Caption = lstMesures.ListCount & " mesure(s)"
End Sub

Private Sub cmdClearMesures_Click()
    Dim lignes As Object, k As Variant
    Set lignes = g_CollMgr.LoadCollection("lignemsa")
    For Each k In lignes.Keys
        If lignes(k).EtudeId = m_Current.Id Then g_CollMgr.DeleteById k, "lignemsa"
    Next k
    g_CollMgr.InvalidateCache "lignemsa"
    RefreshMesures
End Sub

Private Sub cmdCalculer_Click()
    On Error GoTo ErrHandler
    m_Current.NbOperateurs = CLng(Val(txtNbOperateurs.Text))
    m_Current.NbPieces = CLng(Val(txtNbPieces.Text))
    m_Current.NbEssais = CLng(Val(txtNbEssais.Text))
    m_Current.USL = Val(Replace(txtUSL.Text, ",", "."))
    m_Current.LSL = Val(Replace(txtLSL.Text, ",", "."))
    m_Current.Tolerance = m_Current.USL - m_Current.LSL

    CalculerGrR m_Current

    lblEV.Caption = "EV (répétabilité) = " & Format$(m_Current.EcartTypeRepetabilite, "0.0000")
    lblAV.Caption = "AV (reproductibilité) = " & Format$(m_Current.EcartTypeReproductibilite, "0.0000")
    lblGrR.Caption = "GRR = " & Format$(m_Current.GrR, "0.0000")
    lblPctGrR.Caption = "%GRR/TV = " & Format$(m_Current.PourcentGrR, "0.00") & " %"
    lblPctGrRTol.Caption = "%GRR/Tol = " & Format$(m_Current.PourcentGrRTolerance, "0.00") & " %"
    lblNdc.Caption = "ndc = " & m_Current.Ndc
    lblInterpretation.Caption = "Interprétation : " & m_Current.Interpretation()
    Exit Sub
ErrHandler:
    LogEx "FMsa", "cmdCalculer_Click", Err
    ShowError "Calcul", Err.Description
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    m_Current.MoyenId = GetMoyenIdFromCombo()
    m_Current.Caracteristique = txtCaracteristique.Text
    m_Current.Responsable = txtResponsable.Text
    If Len(m_Current.Reference) = 0 Then m_Current.AssignReference()
    If Not m_Current.Validate() Then
        ShowError "Validation", "Moyen et caractéristique requis."
        Exit Sub
    End If
    g_CollMgr.Save m_Current, "etudemsa"
    g_CollMgr.InvalidateCache "etudemsa"
    txtReference.Text = m_Current.Reference
    lblStatus.Caption = "Étude MSA " & m_Current.Reference & " enregistrée."
    Exit Sub
ErrHandler:
    LogEx "FMsa", "cmdSave_Click", Err
    ShowError "Sauvegarde", Err.Description
End Sub

Private Sub cmdRapport_Click()
    On Error GoTo ErrHandler
    If Len(m_Current.Reference) = 0 Then
        ShowError "Rapport", "Enregistrez l'étude avant de générer le rapport."
        Exit Sub
    End If
    Dim path As String
    path = GenererRapportMsa(m_Current)
    If PromptYesNo("Rapport MSA", "Généré :" & vbCrLf & path & vbCrLf & "Ouvrir ?") Then
        Workbooks.Open path
    End If
    Exit Sub
ErrHandler:
    LogEx "FMsa", "cmdRapport_Click", Err
    ShowError "Rapport", Err.Description
End Sub

Private Function GetMoyenIdFromCombo() As String
    If Len(cboMoyen.Text) = 0 Then Exit Function
    Dim moyens As Object, k As Variant
    Set moyens = g_CollMgr.LoadCollection("moyenmesure")
    For Each k In moyens.Keys
        If moyens(k).DisplayName = cboMoyen.Text Then
            GetMoyenIdFromCombo = k
            Exit Function
        End If
    Next k
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub