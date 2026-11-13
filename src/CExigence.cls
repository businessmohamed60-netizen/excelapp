Attribute VB_Name = "CExigence"
Option Explicit

' ============================================================
' CExigence — Exigence réglementaire / norme / client
' ISO 9001 §4.2 / IATF §4.2.2
' ============================================================

Private m_Id As String
Private m_Reference As String           ' 8.5.1.1 / CSR-Ford-4.2
Private m_Source As String             ' ISO 9001 / IATF 16949 / Client / Réglementaire
Private m_ClientId As String           ' si source = Client
Private m_Intitule As String
Private m_Chapitre As String
Private m_Applicable As Boolean
Private m_Conformite As String         ' Conforme / Non conforme / Partielle / Non évaluée
Private m_Preuve As String
Private m_Responsable As String
Private m_DateEvaluation As Date
Private m_DateRevue As Date
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Reference() As String: Reference = m_Reference: End Property
Public Property Let Reference(ByVal v As String): m_Reference = v: End Property
Public Property Get Source() As String: Source = m_Source: End Property
Public Property Let Source(ByVal v As String): m_Source = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Intitule() As String: Intitule = m_Intitule: End Property
Public Property Let Intitule(ByVal v As String): m_Intitule = v: End Property
Public Property Get Chapitre() As String: Chapitre = m_Chapitre: End Property
Public Property Let Chapitre(ByVal v As String): m_Chapitre = v: End Property
Public Property Get Applicable() As Boolean: Applicable = m_Applicable: End Property
Public Property Let Applicable(ByVal v As Boolean): m_Applicable = v: End Property
Public Property Get Conformite() As String: Conformite = m_Conformite: End Property
Public Property Let Conformite(ByVal v As String): m_Conformite = v: End Property
Public Property Get Preuve() As String: Preuve = m_Preuve: End Property
Public Property Let Preuve(ByVal v As String): m_Preuve = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get DateEvaluation() As Date: DateEvaluation = m_DateEvaluation: End Property
Public Property Let DateEvaluation(ByVal v As Date): m_DateEvaluation = v: End Property
Public Property Get DateRevue() As Date: DateRevue = m_DateRevue: End Property
Public Property Let DateRevue(ByVal v As Date): m_DateRevue = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Applicable = True
    m_Conformite = "Non évaluée"
    m_Source = "IATF 16949"
End Sub

Public Function EstNonConforme() As Boolean
    EstNonConforme = m_Applicable And (m_Conformite = "Non conforme" Or m_Conformite = "Partielle")
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_Reference)) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Intitule)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function