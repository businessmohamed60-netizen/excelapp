Attribute VB_Name = "CCaracteristique"
Option Explicit

' ============================================================
' CCaracteristique — Caractéristique spéciale (CC / SC)
' IATF §8.5.1.1 — registre et traçabilité
' ============================================================

Private m_Id As String
Private m_Code As String                ' CC-001 / SC-001
Private m_Libelle As String
Private m_TypeCar As String             ' CC (Critique) / SC (Significative)
Private m_ClientId As String
Private m_Produit As String
Private m_Specification As String
Private m_Tolerance As String
Private m_SymboleClient As String       ' symbole exigé par le client (ex: losange Ford)
Private m_AmdecId As String             ' AMDEC d'origine
Private m_LigneAmdecId As String        ' ligne AMDEC d'origine
Private m_PlanControleId As String      ' plan de contrôle associé
Private m_Statut As String              ' Active / Inactive / Obsolète
Private m_DateCreation As Date
Private m_Commentaire As String

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Code() As String: Code = m_Code: End Property
Public Property Let Code(ByVal v As String): m_Code = v: End Property
Public Property Get Libelle() As String: Libelle = m_Libelle: End Property
Public Property Let Libelle(ByVal v As String): m_Libelle = v: End Property
Public Property Get TypeCar() As String: TypeCar = m_TypeCar: End Property
Public Property Let TypeCar(ByVal v As String): m_TypeCar = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Produit() As String: Produit = m_Produit: End Property
Public Property Let Produit(ByVal v As String): m_Produit = v: End Property
Public Property Get Specification() As String: Specification = m_Specification: End Property
Public Property Let Specification(ByVal v As String): m_Specification = v: End Property
Public Property Get Tolerance() As String: Tolerance = m_Tolerance: End Property
Public Property Let Tolerance(ByVal v As String): m_Tolerance = v: End Property
Public Property Get SymboleClient() As String: SymboleClient = m_SymboleClient: End Property
Public Property Let SymboleClient(ByVal v As String): m_SymboleClient = v: End Property
Public Property Get AmdecId() As String: AmdecId = m_AmdecId: End Property
Public Property Let AmdecId(ByVal v As String): m_AmdecId = v: End Property
Public Property Get LigneAmdecId() As String: LigneAmdecId = m_LigneAmdecId: End Property
Public Property Let LigneAmdecId(ByVal v As String): m_LigneAmdecId = v: End Property
Public Property Get PlanControleId() As String: PlanControleId = m_PlanControleId: End Property
Public Property Let PlanControleId(ByVal v As String): m_PlanControleId = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "Active"
    m_TypeCar = "SC"
End Sub

Public Sub AssignCode()
    Dim coll As Object, k As Variant, n As Long, maxN As Long
    Set coll = g_CollMgr.LoadCollection("caracteristique")
    For Each k In coll.Keys
        If coll(k).TypeCar = m_TypeCar Then
            n = CLng(Val(Mid$(coll(k).Code, 4)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Code = m_TypeCar & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(Trim$(m_Libelle)) = 0 Then Validate = False: Exit Function
    If m_TypeCar <> "CC" And m_TypeCar <> "SC" Then Validate = False: Exit Function
    Validate = True
End Function

Public Function DisplayName() As String
    DisplayName = m_Code & " — " & m_Libelle
End Function