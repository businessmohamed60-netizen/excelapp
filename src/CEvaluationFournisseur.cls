Attribute VB_Name = "CEvaluationFournisseur"
Option Explicit

' ============================================================
' CEvaluationFournisseur — Évaluation périodique (IATF §8.4.2.4)
' ============================================================

Private m_Id As String
Private m_FournisseurId As String
Private m_DateEvaluation As Date
Private m_Periode As String             ' 2026-Q1 / 2026-S1 / 2026
Private m_NoteQualite As Double         ' 0-100
Private m_NoteLivraison As Double
Private m_NoteReactivite As Double
Private m_NotePrix As Double
Private m_NoteGlobale As Double
Private m_NbNonConformites As Long
Private m_NbRetards As Long
Private m_NbReclamations As Long
Private m_Evaluateur As String
Private m_PlanAction As String
Private m_Statut As String              ' En cours / Validée
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get FournisseurId() As String: FournisseurId = m_FournisseurId: End Property
Public Property Let FournisseurId(ByVal v As String): m_FournisseurId = v: End Property
Public Property Get DateEvaluation() As Date: DateEvaluation = m_DateEvaluation: End Property
Public Property Let DateEvaluation(ByVal v As Date): m_DateEvaluation = v: End Property
Public Property Get Periode() As String: Periode = m_Periode: End Property
Public Property Let Periode(ByVal v As String): m_Periode = v: End Property
Public Property Get NoteQualite() As Double: NoteQualite = m_NoteQualite: End Property
Public Property Let NoteQualite(ByVal v As Double): m_NoteQualite = v: End Property
Public Property Get NoteLivraison() As Double: NoteLivraison = m_NoteLivraison: End Property
Public Property Let NoteLivraison(ByVal v As Double): m_NoteLivraison = v: End Property
Public Property Get NoteReactivite() As Double: NoteReactivite = m_NoteReactivite: End Property
Public Property Let NoteReactivite(ByVal v As Double): m_NoteReactivite = v: End Property
Public Property Get NotePrix() As Double: NotePrix = m_NotePrix: End Property
Public Property Let NotePrix(ByVal v As Double): m_NotePrix = v: End Property
Public Property Get NoteGlobale() As Double: NoteGlobale = m_NoteGlobale: End Property
Public Property Let NoteGlobale(ByVal v As Double): m_NoteGlobale = v: End Property
Public Property Get NbNonConformites() As Long: NbNonConformites = m_NbNonConformites: End Property
Public Property Let NbNonConformites(ByVal v As Long): m_NbNonConformites = v: End Property
Public Property Get NbRetards() As Long: NbRetards = m_NbRetards: End Property
Public Property Let NbRetards(ByVal v As Long): m_NbRetards = v: End Property
Public Property Get NbReclamations() As Long: NbReclamations = m_NbReclamations: End Property
Public Property Let NbReclamations(ByVal v As Long): m_NbReclamations = v: End Property
Public Property Get Evaluateur() As String: Evaluateur = m_Evaluateur: End Property
Public Property Let Evaluateur(ByVal v As String): m_Evaluateur = v: End Property
Public Property Get PlanAction() As String: PlanAction = m_PlanAction: End Property
Public Property Let PlanAction(ByVal v As String): m_PlanAction = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateEvaluation = Date
    m_Statut = "En cours"
End Sub

Public Sub RecalculerNoteGlobale()
    m_NoteGlobale = m_NoteQualite * 0.4 + m_NoteLivraison * 0.3 + _
                    m_NoteReactivite * 0.2 + m_NotePrix * 0.1
End Sub

Public Function Validate() As Boolean
    If Len(m_FournisseurId) = 0 Then Validate = False: Exit Function
    If m_NoteQualite < 0 Or m_NoteQualite > 100 Then Validate = False: Exit Function
    Validate = True
End Function