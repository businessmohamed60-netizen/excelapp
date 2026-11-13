Attribute vb_Name = "CPlanSurveillance"
Option Explicit

' ============================================================
' CPlanSurveillance — Plan de surveillance (IATF §8.5.1.1)
' ============================================================

Private m_Id As String
Private m_Reference As String           ' PS-2026-001
Private m_ClientId As String
Private m_Produit As String
Private m_Processus As String
Private m_Responsable As String
Private m_DateCreation As Date
Private m_DateRevue As Date
Private m_Version As String
Private m_Statut As String              ' Brouillon / En revue / Approuvé / Obsolète
Private m_Approbateur As String
Private m_DateApprobation As Date
Private m_Commentaire As String

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Produit() As String: Produit = m_Produit: End Property
Public Property Let Produit(ByVal v As String): m_Produit = v: End Property
Public Property Get Processus() As String: Processus = m_Processus: End Property
Public Property Let Processus(ByVal v As String): m_Processus = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
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
    m_Statut = "Brouillon"
    m_Version = "1.0"
End Sub

Public Sub AssignReference()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("plansurveillance")
    yr = Format$(Year(m_DateCreation), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Reference, 3) = "PS-" & yr Then
            n = CLng(Val(Mid$(coll(k).Reference, 9)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Reference = "PS-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(Trim$(m_Produit)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function