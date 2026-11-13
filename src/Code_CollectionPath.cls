Option Explicit

Private Const REPO_VERSION As String = "1.0"

Public Function CollectionPath(ByVal collection As String) As String
    CollectionPath = DataPath() & collection & ".xml"
End Function

Public Sub Save(ByVal obj As Object, ByVal collection As String)
    Dim dom As Object, root As Object, item As Object
    Dim fname As String, f As Integer
    fname = CollectionPath(collection)
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Dir(fname) <> "" Then
        If Not dom.Load(fname) Then Err.Raise vbObjectError + 2001, "CRepository", "XML corrompu : " & fname
    Else
        dom.loadXML "<" & collection & " version=""" & REPO_VERSION & """/>"
    End If
    Set root = dom.documentElement
    Set item = FindItemById(root, obj.Id)
    If item Is Nothing Then
        Set item = dom.createElement("item")
        root.appendChild item
    End If
    SerializeItem obj, item, dom
    If Dir(fname) <> "" Then
        On Error Resume Next
        Kill fname & ".bak"
        On Error GoTo 0
        FileCopy fname, fname & ".bak"
    End If
    f = FreeFile
    Open fname For Output Lock Write As #f
    Print #f, dom.XML
    Close #f
    Log LogInfo, "CRepository", "Saved " & collection & " id=" & obj.Id
End Sub

Public Function LoadById(ByVal id As String, ByVal collection As String, ByVal obj As Object) As Boolean
    Dim dom As Object, item As Object
    If Dir(CollectionPath(collection)) = "" Then Exit Function
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(CollectionPath(collection)) Then Exit Function
    Set item = FindItemById(dom.documentElement, id)
    If item Is Nothing Then Exit Function
    DeserializeItem item, obj
    LoadById = True
End Function

Public Function LoadAll(ByVal collection As String, ByVal factory As Object, ByRef outDict As Object) As Long
    Dim dom As Object, item As Object, newObj As Object
    Set outDict = CreateObject("Scripting.Dictionary")
    If Dir(CollectionPath(collection)) = "" Then Exit Function
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(CollectionPath(collection)) Then Exit Function
    For Each item In dom.documentElement.ChildNodes
        Set newObj = factory.CreateEmpty(collection)
        DeserializeItem item, newObj
        If Len(newObj.Id) > 0 Then outDict.Add newObj.Id, newObj
    Next item
    LoadAll = outDict.Count
End Function

Public Function DeleteById(ByVal id As String, ByVal collection As String) As Boolean
    Dim dom As Object, item As Object
    If Dir(CollectionPath(collection)) = "" Then Exit Function
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(CollectionPath(collection)) Then Exit Function
    Set item = FindItemById(dom.documentElement, id)
    If item Is Nothing Then Exit Function
    dom.documentElement.removeChild item
    dom.save CollectionPath(collection)
    DeleteById = True
    Log LogInfo, "CRepository", "Deleted " & collection & " id=" & id
End Function

Private Function FindItemById(ByVal root As Object, ByVal id As String) As Object
    Dim item As Object, n As Object
    For Each item In root.ChildNodes
        Set n = item.selectSingleNode("Id")
        If Not n Is Nothing Then
            If n.Text = id Then Set FindItemById = item: Exit Function
        End If
    Next item
End Function

Private Sub SerializeItem(ByVal obj As Object, ByVal item As Object, ByVal dom As Object)
    Dim props As Variant, i As Long, e As Object, v As Variant, old As Object
    props = DomainProperties(TypeName(obj))
    For i = LBound(props) To UBound(props)
        On Error Resume Next
        v = CallByName(obj, props(i), VbGet)
        If Err.Number = 0 Then
            Set e = dom.createElement(CStr(props(i)))
            If IsDate(v) Then e.Text = FormatIsoDate(CDate(v)) Else e.Text = CStr(v)
            Set old = item.selectSingleNode(CStr(props(i)))
            If Not old Is Nothing Then item.removeChild old
            item.appendChild e
        End If
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

Private Sub DeserializeItem(ByVal item As Object, ByVal obj As Object)
    Dim props As Variant, i As Long, n As Object, v As Variant
    props = DomainProperties(TypeName(obj))
    For i = LBound(props) To UBound(props)
        Set n = item.selectSingleNode(CStr(props(i)))
        If Not n Is Nothing Then
            v = n.Text
            On Error Resume Next
            If IsDateLike(CStr(props(i))) Then
                CallByName obj, props(i), VbLet, ParseIsoDate(CStr(v))
            Else
                CallByName obj, props(i), VbLet, CStr(v)
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function DomainProperties(ByVal clsName As String) As Variant
    Select Case clsName
        Case "CClient"
            DomainProperties = Array("Id", "Code", "RaisonSociale", "Secteur", _
                "ContactNom", "ContactEmail", "ContactTel", "Adresse", "Pays", "Norme", "Actif", "DateCreation")
        Case "CAudit"
            DomainProperties = Array("Id", "ClientId", "Processus", "TypeAudit", "Norme", _
                "DateDebut", "DateFin", "Auditeur", "Statut", "Score", "DateCreation")
        Case "CConstat"
            DomainProperties = Array("Id", "AuditId", "Chapitre", "Description", "Classification", "Preuve", "DateCreation")
        Case "CNonConformite"
            DomainProperties = Array("Id", "Numero", "ClientId", "AuditId", "ConstatId", "Source", "Processus", _
                "Description", "Classification", "Gravite", "DateDetection", "DetectePar", "ContainmentImmediate", _
                "Statut", "DateCloture", "DateCreation")
        Case "CCapa"
            DomainProperties = Array("Id", "Numero", "NcId", "TypeAction", "Description", "Responsable", _
                "DateOuverture", "DateEcheance", "DateRealisation", "Efficacite", "Statut", "DateCreation")
        Case "C8D"
            DomainProperties = Array("Id", "Numero", "NcId", "Statut", _
                "D1_Champion", "D1_Leader", "D1_Membres", _
                "D2_Description", "D2_Date", "D2_5W2H", _
                "D3_Actions", "D3_Date", "D3_Verification", _
                "D4_Ishikawa", "D4_5Pourquoi", "D4_CauseRacine", "D4_CauseEchappee", "D4_Date", _
                "D5_Actions", "D5_Responsable", "D5_Date", _
                "D6_Actions", "D6_Date", "D6_Validation", _
                "D7_Actions", "D7_Date", "D8_Date", "D8_Bilan", "DateCreation")
        Case "CProjetAPQP"
            DomainProperties = Array("Id", "Reference", "ClientId", "NomProjet", "Programme", "Responsable", _
                "DateLancement", "DateCibleSOP", "Phase1_Planification", "Phase2_Conception", _
                "Phase3_Industrialisation", "Phase4_Validation", "Phase5_Production", "Statut", "Avancement", "DateCreation")
        Case "CPpapDossier"
            DomainProperties = Array("Id", "Reference", "ClientId", "ApqpId", "Piece", "ReferencePiece", _
                "NumeroSoumission", "NiveauPpap", "Statut", "DateSoumission", "DateAcceptation", "DateCreation")
        Case "CAmdec"
            DomainProperties = Array("Id", "Reference", "ClientId", "TypeAmdec", "Piece", "Processus", "Equipe", "DateCreation")
        Case "CSpcEtude"
            DomainProperties = Array("Id", "Reference", "ClientId", "Caracteristique", "Unite", _
                "LimiteSpecInf", "LimiteSpecSup", "TailleSousGroupe", "DateCreation")
        Case "CMsaEtude"
            DomainProperties = Array("Id", "Reference", "ClientId", "MoyenMesure", "Caracteristique", _
                "Tolerance", "NbPieces", "NbOperateurs", "NbEssais", "DateCreation")
        Case Else
            DomainProperties = Array("Id")
    End Select
End Function

Private Function IsDateLike(ByVal propName As String) As Boolean
    Select Case propName
        Case "DateCreation", "DateDebut", "DateFin", "DateCloture", "DateEcheance", "DateOuverture", _
             "DateRealisation", "DateDetection", "DateSoumission", "DateAcceptation", "DateLancement", _
             "DateCibleSOP", "Phase1_Planification", "Phase2_Conception", "Phase3_Industrialisation", _
             "Phase4_Validation", "Phase5_Production", "D2_Date", "D3_Date", "D4_Date", "D5_Date", _
             "D6_Date", "D7_Date", "D8_Date"
            IsDateLike = True
    End Select
End Function