Option Explicit

Private m_Id As String
Private m_Reference As String
Private m_ClientId As String
Private m_NomProjet As String
Private m_Programme As String
Private m_Responsable As String
Private m_DateLancement As Date
Private m_DateCibleSOP As Date
Private m_Phase1_Planification As Date
Private m_Phase2_Conception As Date
Private m_Phase3_Industrialisation As Date
Private m_Phase4_Validation As Date
Private m_Phase5_Production As Date
Private m_Statut As String
Private m_Avancement As Long
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get NomProjet() As String: NomProjet = m_NomProjet: End Property
Public Property Let NomProjet(ByVal v As String): m_NomProjet = v: End Property
Public Property Get Programme() As String: Programme = m_Programme: End Property
Public Property Let Programme(ByVal v As String): m_Programme = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get DateLancement() As Date: DateLancement = m_DateLancement: End Property
Public Property Let DateLancement(ByVal v As Date): m_DateLancement = v: End Property
Public Property Get DateCibleSOP() As Date: DateCibleSOP = m_DateCibleSOP: End Property
Public Property Let DateCibleSOP(ByVal v As Date): m_DateCibleSOP = v: End Property
Public Property Get Phase1_Planification() As Date: Phase1_Planification = m_Phase1_Planification: End Property
Public Property Let Phase1_Planification(ByVal v As Date): m_Phase1_Planification = v: End Property
Public Property Get Phase2_Conception() As Date: Phase2_Conception = m_Phase2_Conception: End Property
Public Property Let Phase2_Conception(ByVal v As Date): m_Phase2_Conception = v: End Property
Public Property Get Phase3_Industrialisation() As Date: Phase3_Industrialisation = m_Phase3_Industrialisation: End Property
Public Property Let Phase3_Industrialisation(ByVal v As Date): m_Phase3_Industrialisation = v: End Property
Public Property Get Phase4_Validation() As Date: Phase4_Validation = m_Phase4_Validation: End Property
Public Property Let Phase4_Validation(ByVal v As Date): m_Phase4_Validation = v: End Property
Public Property Get Phase5_Production() As Date: Phase5_Production = m_Phase5_Production: End Property
Public Property Let Phase5_Production(ByVal v As Date): m_Phase5_Production = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Avancement() As Long: Avancement = m_Avancement: End Property
Public Property Let Avancement(ByVal v As Long): m_Avancement = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateLancement = Date
    m_DateCibleSOP = Date + 180
    m_Statut = "En cours"
    m_Avancement = 0
End Sub

Public Sub AssignReference()
    Dim dom As Object, root As Object, yr As String
    Dim fname As String, item As Object, numeroNode As Object
    Dim maxN As Long, parts() As String, n As Long
    fname = DataPath() & "apqp.xml"
    yr = Format$(Year(Date), "0000")
    If Dir(fname) <> "" Then
        Set dom = CreateObject("MSXML2.DOMDocument.6.0")
        dom.async = False
        If dom.Load(fname) Then
            Set root = dom.documentElement
            For Each item In root.ChildNodes
                Set numeroNode = item.selectSingleNode("Reference")
                If Not numeroNode Is Nothing Then
                    parts = Split(numeroNode.Text, "-")
                    If UBound(parts) >= 2 Then
                        If parts(1) = yr Then
                            n = CLng(Val(parts(2)))
                            If n > maxN Then maxN = n
                        End If
                    End If
                End If
            Next item
        End If
    End If
    m_Reference = "APQP-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function CalculerAvancement() As Long
    Dim n As Long
    If IsDate(m_Phase1_Planification) And m_Phase1_Planification > 0 Then n = n + 1
    If IsDate(m_Phase2_Conception) And m_Phase2_Conception > 0 Then n = n + 1
    If IsDate(m_Phase3_Industrialisation) And m_Phase3_Industrialisation > 0 Then n = n + 1
    If IsDate(m_Phase4_Validation) And m_Phase4_Validation > 0 Then n = n + 1
    If IsDate(m_Phase5_Production) And m_Phase5_Production > 0 Then n = n + 1
    m_Avancement = n * 20
    CalculerAvancement = m_Avancement
End Function

Public Function Validate() As Boolean
    If Len(m_ClientId) = 0 Then Exit Function
    If Len(Trim$(m_NomProjet)) = 0 Then Exit Function
    If m_DateCibleSOP < m_DateLancement Then Exit Function
    Validate = True
End Function