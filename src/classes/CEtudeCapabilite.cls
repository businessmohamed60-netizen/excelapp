Attribute VB_Name = "CEtudeCapabilite"
Option Explicit

' ============================================================
' CEtudeCapabilite — Étude de capabilité (Cp, Cpk, Pp, Ppk)
' IATF §9.1.1.1
' ============================================================

Private m_Id As String
Private m_Reference As String           ' CAP-2026-001
Private m_Caracteristique As String
Private m_MoyenId As String
Private m_ClientId As String
Private m_USL As Double
Private m_LSL As Double
Private m_Cible As Double
Private m_NbValeurs As Long
Private m_Moyenne As Double
Private m_EcartTypeCourt As Double      ' écart-type intra-série (R̄/d2)
Private m_EcartTypeLong As Double       ' écart-type global (tous les points)
Private m_Cp As Double
Private m_Cpk As Double
Private m_Pp As Double
Private m_Ppk As Double
Private m_Statut As String              ' Brouillon / Validée / Rejetée
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
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get USL() As Double: USL = m_USL: End Property
Public Property Let USL(ByVal v As Double): m_USL = v: End Property
Public Property Get LSL() As Double: LSL = m_LSL: End Property
Public Property Let LSL(ByVal v As Double): m_LSL = v: End Property
Public Property Get Cible() As Double: Cible = m_Cible: End Property
Public Property Let Cible(ByVal v As Double): m_Cible = v: End Property
Public Property Get NbValeurs() As Long: NbValeurs = m_NbValeurs: End Property
Public Property Let NbValeurs(ByVal v As Long): m_NbValeurs = v: End Property
Public Property Get Moyenne() As Double: Moyenne = m_Moyenne: End Property
Public Property Let Moyenne(ByVal v As Double): m_Moyenne = v: End Property
Public Property Get EcartTypeCourt() As Double: EcartTypeCourt = m_EcartTypeCourt: End Property
Public Property Let EcartTypeCourt(ByVal v As Double): m_EcartTypeCourt = v: End Property
Public Property Get EcartTypeLong() As Double: EcartTypeLong = m_EcartTypeLong: End Property
Public Property Let EcartTypeLong(ByVal v As Double): m_EcartTypeLong = v: End Property
Public Property Get Cp() As Double: Cp = m_Cp: End Property
Public Property Let Cp(ByVal v As Double): m_Cp = v: End Property
Public Property Get Cpk() As Double: Cpk = m_Cpk: End Property
Public Property Let Cpk(ByVal v As Double): m_Cpk = v: End Property
Public Property Get Pp() As Double: Pp = m_Pp: End Property
Public Property Let Pp(ByVal v As Double): m_Pp = v: End Property
Public Property Get Ppk() As Double: Ppk = m_Ppk: End Property
Public Property Let Ppk(ByVal v As Double): m_Ppk = v: End Property
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
End Sub

Public Sub AssignReference()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("etudecapabilite")
    yr = Format$(Year(m_DateEtude), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Reference, 9) = "CAP-" & yr Then
            n = CLng(Val(Mid$(coll(k).Reference, 10)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Reference = "CAP-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(Trim$(m_Caracteristique)) = 0 Then Validate = False: Exit Function
    If m_USL <= m_LSL Then Validate = False: Exit Function
    Validate = True
End Function

' Interprétation IATF : Cpk ≥ 1.67 acceptable, ≥ 1.33 conditionnel, < 1.33 insuffisant
Public Function Interpretation() As String
    If m_Cpk <= 0 Then Interpretation = "Non calculé": Exit Function
    If m_Cpk >= 1.67 Then
        Interpretation = "Capable"
    ElseIf m_Cpk >= 1.33 Then
        Interpretation = "Acceptable"
    Else
        Interpretation = "Insuffisant"
    End If
End Function