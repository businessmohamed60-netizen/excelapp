Attribute VB_Name = "CAuditFournisseur"
Option Explicit

' ============================================================
' CAuditFournisseur — Audit fournisseur (IATF §8.4.2.4.1)
' ============================================================

Private m_Id As String
Private m_FournisseurId As String
Private m_Reference As String           ' AF-2026-001
Private m_TypeAudit As String           ' Initial / Surveillance / Renouvellement / Pour cause
Private m_DateAudit As Date
Private m_Auditeur As String
Private m_Score As Double               ' 0-100
Private m_NbMajeurs As Long
Private m_NbMineurs As Long
Private m_NbObservations As Long
Private m_Conclusion As String          ' Conforme / Conforme avec réserves / Non conforme
Private m_PlanAction As String
Private m_Statut As String              ' Planifié / Réalisé / Clôturé
Private m_DateCloture As Date
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get FournisseurId() As String: FournisseurId = m_FournisseurId: End Property
Public Property Let FournisseurId(ByVal v As String): m_FournisseurId = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get TypeAudit() As String: TypeAudit = m_TypeAudit: End Property
Public Property Let TypeAudit(ByVal v As String): m_TypeAudit = v: End Property
Public Property Get DateAudit() As Date: DateAudit = m_DateAudit: End Property
Public Property Let DateAudit(ByVal v As Date): m_DateAudit = v: End Property
Public Property Get Auditeur() As String: Auditeur = m_Auditeur: End Property
Public Property Let Auditeur(ByVal v As String): m_Auditeur = v: End Property
Public Property Get Score() As Double: Score = m_Score: End Property
Public Property Let Score(ByVal v As Double): m_Score = v: End Property
Public Property Get NbMajeurs() As Long: NbMajeurs = m_NbMajeurs: End Property
Public Property Let NbMajeurs(ByVal v As Long): m_NbMajeurs = v: End Property
Public Property Get NbMineurs() As Long: NbMineurs = m_NbMineurs: End Property
Public Property Let NbMineurs(ByVal v As Long): m_NbMineurs = v: End Property
Public Property Get NbObservations() As Long: NbObservations = m_NbObservations: End Property
Public Property Let NbObservations(ByVal v As Long): m_NbObservations = v: End Property
Public Property Get Conclusion() As String: Conclusion = m_Conclusion: End Property
Public Property Let Conclusion(ByVal v As String): m_Conclusion = v: End Property
Public Property Get PlanAction() As String: PlanAction = m_PlanAction: End Property
Public Property Let PlanAction(ByVal v As String): m_PlanAction = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get DateCloture() As Date: DateCloture = m_DateCloture: End Property
Public Property Let DateCloture(ByVal v As Date): m_DateCloture = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateAudit = Date
    m_Statut = "Planifié"
    m_TypeAudit = "Surveillance"
End Sub

Public Sub AssignReference()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("auditfournisseur")
    yr = Format$(Year(m_DateAudit), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Reference, 6) = "AF-" & yr Then
            n = CLng(Val(Mid$(coll(k).Reference, 9)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Reference = "AF-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Function Validate() As Boolean
    If Len(m_FournisseurId) = 0 Then Validate = False: Exit Function
    If m_Score < 0 Or m_Score > 100 Then Validate = False: Exit Function
    Validate = True
End Function