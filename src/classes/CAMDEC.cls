Attribute VB_Name = "CAMDEC"
Option Explicit

' ============================================================
' CAMDEC — En-tête d'analyse AMDEC
' IATF §8.5.1.1 (AMDEC processus) / §8.3.5.1 (AMDEC produit)
' ============================================================

Private m_Id As String
Private m_Reference As String           ' AMDEC-PRO-2026-001
Private m_TypeAmdec As String          ' Processus / Produit / Moyens / Fournisseur
Private m_ClientId As String
Private m_Processus As String
Private m_Produit As String
Private m_Responsable As String
Private m_Equipe As String
Private m_DateCreation As Date
Private m_DateRevue As Date
Private m_Version As String            ' 1.0
Private m_Statut As String             ' Brouillon / En revue / Approuvée / Obsolète
Private m_Approbateur As String
Private m_DateApprobation As Date
Private m_Commentaire As String

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get TypeAmdec() As String: TypeAmdec = m_TypeAmdec: End Property
Public Property Let TypeAmdec(ByVal v As String): m_TypeAmdec = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Processus() As String: Processus = m_Processus: End Property
Public Property Let Processus(ByVal v As String): m_Processus = v: End Property
Public Property Get Produit() As String: Produit = m_Produit: End Property
Public Property Let Produit(ByVal v As String): m_Produit = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get Equipe() As String: Equipe = m_Equipe: End Property
Public Property Let Equipe(ByVal v As String): m_Equipe = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property
Public Property Get DateRevue() As Date: DateRevue = m_DateRevue: End Property
Public Property Let DateRevue(ByVal v As Date): m_DateRevue = v: End Property
Public Property Get Version() As String: Version = m_Version: End Property
Public Property Let Version(ByVal v As String): m_Version = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Approbateur() As String: Approbateur = m_Approbateur: End Property
Public Property Let Approbateur(ByVal v As String): m_Approbateur = v: End Property
Public Property Get DateApprobation() As Date: DateApprobation = m_DateApprobation: End Property
Public Property Let DateApprobation(ByVal v As Date): m_DateApprobation = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateRevue = Date + 365
    m_Version = "1.0"
    m_Statut = "Brouillon"
    m_TypeAmdec = "Processus"
End Sub

Public Sub AssignReference()
    Dim dom As Object, root As Object, yr As String
    Dim fname As String, item As Object, node As Object
    Dim maxN As Long, parts() As String, n As Long
    fname = DataPath() & "amdec.xml"
    yr = Format$(Year(m_DateCreation), "0000")

    If Dir(fname) <> "" Then
        Set dom = CreateObject("MSXML2.DOMDocument.6.0")
        dom.async = False
        If dom.Load(fname) Then
            Set root = dom.documentElement
            For Each item In root.ChildNodes
                Set node = item.selectSingleNode("Reference")
                If Not node Is Nothing Then
                    parts = Split(node.Text, "-")
                    If UBound(parts) >= 3 Then
                        If parts(2) = yr Then
                            n = CLng(Val(parts(3)))
                            If n > maxN Then maxN = n
                        End If
                    End If
                End If
            Next item
        End If
    End If

    Dim candidate As String
    Do
        maxN = maxN + 1
        candidate = "AMDEC-" & UCase$(Left$(m_TypeAmdec, 3)) & "-" & yr & "-" & Format$(maxN, "000")
    Loop While ReferenceExiste(candidate, fname)
    m_Reference = candidate
End Sub

Private Function ReferenceExiste(ByVal ref As String, ByVal fname As String) As Boolean
    If Dir(fname) = "" Then Exit Function
    Dim dom As Object, item As Object, node As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(fname) Then Exit Function
    For Each item In dom.documentElement.ChildNodes
        Set node = item.selectSingleNode("Reference")
        If Not node Is Nothing Then
            If node.Text = ref Then ReferenceExiste = True: Exit Function
        End If
    Next item
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_Processus)) = 0 And Len(Trim$(m_Produit)) = 0 Then
        Validate = False: Exit Function
    End If
    If Len(Trim$(m_Responsable)) = 0 Then Validate = False: Exit Function
    If m_Statut = "Approuvée" And Not IsDate(m_DateApprobation) Then
        Validate = False: Exit Function
    End If
    Validate = True
End Function

Public Function DisplayName() As String
    DisplayName = m_Reference & " — " & IIf(Len(m_Processus) > 0, m_Processus, m_Produit)
End Function