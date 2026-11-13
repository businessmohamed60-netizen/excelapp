Attribute VB_Name = "FMigration"
Option Explicit

Private Sub UserForm_Initialize()
    Caption = "Migration XML → base de données"
    cmdMigrerTout.Caption = "Migrer tout"
    cmdMigrerSelection.Caption = "Migrer la sélection"
    cmdClose.Caption = T("btn.close")
    lblInfo.Caption = "Source : " & DataPath() & vbCrLf & _
        "Destination : " & g_Config.DbPath & " (" & g_Config.DbEngine & ")"
    With lstCollections
        .ColumnCount = 4
        .ColumnWidths = "100;60;60;80"
    End With
    RefreshList
End Sub

Private Sub RefreshList()
    lstCollections.Clear
    Dim tables As Variant, i As Long, nXml As Long, nSql As Long
    tables = Array("clients", "audits", "constats", "nc", "capa", "huitd", _
        "auditprogramme", "document", "exigence", _
        "amdec", "ligneamdec", "plancontrole", "ligneplan", "caracteristique", _
        "moyenmesure", "etalonnage", "etudemsa", "lignemsa", _
        "etudecapabilite", "cartecontrole", "lignecarte", "audittrail")
    For i = LBound(tables) To UBound(tables)
        Dim coll As Object
        On Error Resume Next
        Set coll = g_CollMgr.LoadCollection(CStr(tables(i)))
        nXml = 0: If Not coll Is Nothing Then nXml = coll.Count
        Err.Clear
        Set coll = Nothing

        ' Compter côté SQL
        nSql = 0
        On Error Resume Next
        Dim rs As Object
        Set rs = DbQuery("SELECT COUNT(*) AS n FROM " & tables(i))
        If Not rs.EOF Then nSql = rs("n")
        If Not rs Is Nothing Then rs.Close
        Err.Clear

        lstCollections.AddItem CStr(tables(i))
        lstCollections.List(lstCollections.ListCount - 1, 1) = CStr(nXml)
        lstCollections.List(lstCollections.ListCount - 1, 2) = CStr(nSql)
        lstCollections.List(lstCollections.ListCount - 1, 3) = IIf(nSql >= nXml, "OK", "À migrer")
    Next i
End Sub

Private Sub cmdMigrerTout_Click()
    If Not PromptYesNo("Migration", "Migrer toutes les collections ?") Then Exit Sub
    MigrerTout
    RefreshList
End Sub

Private Sub cmdMigrerSelection_Click()
    Dim i As Long, n As Long
    For i = 0 To lstCollections.ListCount - 1
        If lstCollections.Selected(i) Then
            n = MigrerCollection(lstCollections.List(i, 0))
            lblStatus.Caption = "Migration de " & lstCollections.List(i, 0) & " : " & n & " enregistrements"
        End If
    Next i
    RefreshList
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub