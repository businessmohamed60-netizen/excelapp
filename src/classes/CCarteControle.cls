Attribute VB_Name = "CCarteControle"
Option Explicit

' ============================================================
' CCarteControle — Carte de contrôle SPC (X̄-R)
' IATF §9.1.1.1 / règles de Nelson
' ============================================================

Private m_Id As String
Private m_Reference As String           ' SPC-2026-001
Private m_Caracteristique As String
Private m_MoyenId As String
Private m_TypeCarte As String           ' Xbar-R / Xbar-S / p / c
Private m_SousGroupeTaille As Long      ' n (ex: 5)
Private m_USL As Double
Private m_LSL As Double
Private m_Cible As Double
Private m_LigneCentre As Double         ' X̄̄
Private m_LimiteSup As Double           ' UCL
Private m_LimiteInf As Double           ' LCL
Private m_LigneCentreR As Double        ' R̄
Private m_LimiteSupR As Double          ' UCLR
Private m_LimiteInfR As Double          ' LCLR
Private m_NbSousGroupes As Long
Private m_Statut As String              ' Brouillon / Validée
Private m_DateEtude As Date
Private m_Responsable As String
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get Caracteristique() As String: Caracteristique = m_Caracteristique: End Property
Public Property Let Caracteristique(ByVal v As String): m_Caracteristique = v: End Property
Public Property Get MoyenId() As String: MoyenId = m_MoyenId: End Property
Public Property Let MoyenId(ByVal v As String): m_MoyenId = v: End Property
Public Property Get TypeCarte() As String: TypeCarte = m_TypeCarte: End Property
Public Property Let TypeCarte(ByVal v As String): m_TypeCarte = v: End Property
Public Property Get SousGroupeTaille() As Long: SousGroupeTaille = m_SousGroupeTaille: End Property
Public Property Let SousGroupeTaille(ByVal v As Long): m_SousGroupeTaille = v: End Property
Public Property Get USL() As Double: USL = m_USL: End Property
Public Property Let USL(ByVal v As Double): m_USL = v: End Property
Public Property Get LSL() As Double: LSL = m_LSL: End Property
Public Property Let LSL(ByVal v As Double): m_LSL = v: End Property
Public Property Get Cible() As Double: Cible = m_Cible: End Property
Public Property Let Cible(ByVal v As Double): m_Cible = v: End Property
Public Property Get LigneCentre() As Double: LigneCentre = m_LigneCentre: End Property
Public Property Let LigneCentre(ByVal v As Double): m_LigneCentre = v: End Property
Public Property Get LimiteSup() As Double: LimiteSup = m_LimiteSup: End Property
Public Property Let LimiteSup(ByVal v As Double): m_LimiteSup = v: End Property
Public Property Get LimiteInf() As Double: LimiteInf = m_LimiteInf: End Property
Public Property Let LimiteInf(ByVal v As Double): m_LimiteInf = v: End Property
Public Property Get LigneCentreR() As Double: LigneCentreR = m_LigneCentreR: End Property
Public Property Let LigneCentreR(ByVal v As Double): m_LigneCentreR = v: End Property
Public Property Get LimiteSupR() As Double: LimiteSupR = m_LimiteSupR: End Property
Public Property Let LimiteSupR(ByVal v As Double): m_LimiteSupR = v: End Property
Public Property Get LimiteInfR() As Double: LimiteInfR = m_LimiteInfR: End Property
Public Property Let LimiteInfR(ByVal v As Double): m_LimiteInfR = v: End Property
Public Property Get NbSousGroupes() As Long: NbSousGroupes = m_NbSousGroupes: End Property
Public Property Let NbSousGroupes(ByVal v As Long): m_NbSousGroupes = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get DateEtude() As Date: DateEtude = m_DateEtude: End Property
Public Property Let DateEtude(ByVal v As Date): m_DateEtude = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateEtude = Date
    m_Statut = "Brouillon"
    m_TypeCarte = "Xbar-R"
    m_SousGroupeTaille = 5
End Sub

Public Sub AssignReference()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("cartecontrole")
    yr = Format$(Year(m_DateEtude), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Reference, 9) = "SPC-" & yr Then
            n = CLng(Val(Mid$(coll(k).Reference, 10)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Reference = "SPC-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(Trim$(m_Caracteristique)) = 0 Then Validate = False: Exit Function
    If m_SousGroupeTaille < 2 Then Validate = False: Exit Function
    Validate = True
End Function