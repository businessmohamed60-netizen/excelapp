Attribute VB_Name = "CEtudeMsa"
Option Explicit

' ============================================================
' CEtudeMsa — Étude MSA / R&R (Gage R&R)
' IATF §7.1.5.1.1 / MSA 4ᵉ éd.
' ============================================================

Private m_Id As String
Private m_Reference As String           ' MSA-2026-001
Private m_MoyenId As String
Private m_Caracteristique As String
Private m_ClientId As String
Private m_NbOperateurs As Long
Private m_NbPieces As Long
Private m_NbEssais As Long
Private m_Tolerance As Double           ' tolérance (USL-LSL)
Private m_USL As Double
Private m_LSL As Double
Private m_Moyenne As Double
Private m_EtendueMoyenne As Double
Private m_EcartTypeRepetabilite As Double   ' EV
Private m_EcartTypeReproductibilite As Double ' AV
Private m_GrR As Double                 ' GRR = sqrt(EV² + AV²)
Private m_PourcentGrR As Double         ' %GRR / TV
Private m_PourcentGrRTolerance As Double ' %GRR / tolérance
Private m_Ndc As Long                   ' nombre de catégories distinctes
Private m_Statut As String              ' Brouillon / Validée / Rejetée
Private m_DateEtude As Date
Private m_Responsable As String
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get MoyenId() As String: MoyenId = m_MoyenId: End Property
Public Property Let MoyenId(ByVal v As String): m_MoyenId = v: End Property
Public Property Get Caracteristique() As String: Caracteristique = m_Caracteristique: End Property
Public Property Let Caracteristique(ByVal v As String): m_Caracteristique = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get NbOperateurs() As Long: NbOperateurs = m_NbOperateurs: End Property
Public Property Let NbOperateurs(ByVal v As Long): m_NbOperateurs = v: End Property
Public Property Get NbPieces() As Long: NbPieces = m_NbPieces: End Property
Public Property Let NbPieces(ByVal v As Long): m_NbPieces = v: End Property
Public Property Get NbEssais() As Long: NbEssais = m_NbEssais: End Property
Public Property Let NbEssais(ByVal v As Long): m_NbEssais = v: End Property
Public Property Get Tolerance() As Double: Tolerance = m_Tolerance: End Property
Public Property Let Tolerance(ByVal v As Double): m_Tolerance = v: End Property
Public Property Get USL() As Double: USL = m_USL: End Property
Public Property Let USL(ByVal v As Double): m_USL = v: End Property
Public Property Get LSL() As Double: LSL = m_LSL: End Property
Public Property Let LSL(ByVal v As Double): m_LSL = v: End Property
Public Property Get Moyenne() As Double: Moyenne = m_Moyenne: End Property
Public Property Let Moyenne(ByVal v As Double): m_Moyenne = v: End Property
Public Property Get EtendueMoyenne() As Double: EtendueMoyenne = m_EtendueMoyenne: End Property
Public Property Let EtendueMoyenne(ByVal v As Double): m_EtendueMoyenne = v: End Property
Public Property Get EcartTypeRepetabilite() As Double: EcartTypeRepetabilite = m_EcartTypeRepetabilite: End Property
Public Property Let EcartTypeRepetabilite(ByVal v As Double): m_EcartTypeRepetabilite = v: End Property
Public Property Get EcartTypeReproductibilite() As Double: EcartTypeReproductibilite = m_EcartTypeReproductibilite: End Property
Public Property Let EcartTypeReproductibilite(ByVal v As Double): m_EcartTypeReproductibilite = v: End Property
Public Property Get GrR() As Double: GrR = m_GrR: End Property
Public Property Let GrR(ByVal v As Double): m_GrR = v: End Property
Public Property Get PourcentGrR() As Double: PourcentGrR = m_PourcentGrR: End Property
Public Property Let PourcentGrR(ByVal v As Double): m_PourcentGrR = v: End Property
Public Property Get PourcentGrRTolerance() As Double: PourcentGrRTolerance = m_PourcentGrRTolerance: End Property
Public Property Let PourcentGrRTolerance(ByVal v As Double): m_PourcentGrRTolerance = v: End Property
Public Property Get Ndc() As Long: Ndc = m_Ndc: End Property
Public Property Let Ndc(ByVal v As Long): m_Ndc = v: End Property
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
    m_NbOperateurs = 3
    m_NbPieces = 10
    m_NbEssais = 3
End Sub

Public Sub AssignReference()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("etudemsa")
    yr = Format$(Year(m_DateEtude), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Reference, 9) = "MSA-" & yr Then
            n = CLng(Val(Mid$(coll(k).Reference, 10)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Reference = "MSA-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(m_MoyenId) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Caracteristique)) = 0 Then Validate = False: Exit Function
    If m_NbOperateurs < 2 Or m_NbPieces < 5 Or m_NbEssais < 2 Then
        Validate = False: Exit Function
    End If
    Validate = True
End Function

' Interprétation MSA : <10% acceptable, 10-30% conditionnel, >30% rejeté
Public Function Interpretation() As String
    If m_PourcentGrRTolerance <= 0 Then Interpretation = "Non calculé": Exit Function
    If m_PourcentGrRTolerance < 10 Then
        Interpretation = "Acceptable"
    ElseIf m_PourcentGrRTolerance <= 30 Then
        Interpretation = "Conditionnel"
    Else
        Interpretation = "Rejeté"
    End If
End Function