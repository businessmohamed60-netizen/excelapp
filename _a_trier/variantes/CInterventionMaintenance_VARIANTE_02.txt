Attribute VB_Name = "CInterventionMaintenance"
Option Explicit

' ============================================================
' CInterventionMaintenance — Intervention réalisée
' ============================================================

Private m_Id As String
Private m_EquipementId As String
Private m_MaintenanceId As String       ' optionnel (si issue d'un plan)
Private m_DateIntervention As Date
Private m_TypeIntervention As String    ' Préventive / Corrective / Amélioration
Private m_Description As String
Private m_Technicien As String
Private m_DureeHeures As Double
Private m_PiecesUtilisees As String
Private m_CoutMainOeuvre As Double
Private m_CoutPieces As Double
Private m_CoutTotal As Double
Private m_Resultat As String            ' Résolu / Partiel / Non résolu
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get EquipementId() As String: EquipementId = m_EquipementId: End Property
Public Property Let EquipementId(ByVal v As String): m_EquipementId = v: End Property
Public Property Get MaintenanceId() As String: MaintenanceId = m_MaintenanceId: End Property
Public Property Let MaintenanceId(ByVal v As String): m_MaintenanceId = v: End Property
Public Property Get DateIntervention() As Date: DateIntervention = m_DateIntervention: End Property
Public Property Let DateIntervention(ByVal v As Date): m_DateIntervention = v: End Property
Public Property Get TypeIntervention() As String: TypeIntervention = m_TypeIntervention: End Property
Public Property Let TypeIntervention(ByVal v As String): m_TypeIntervention = v: End Property
Public Property Get Description() As String: Description = m_Description: End Property
Public Property Let Description(ByVal v As String): m_Description = v: End Property
Public Property Get Technicien() As String: Technicien = m_Technicien: End Property
Public Property Let Technicien(ByVal v As String): m_Technicien = v: End Property
Public Property Get DureeHeures() As Double: DureeHeures = m_DureeHeures: End Property
Public Property Let DureeHeures(ByVal v As Double): m_DureeHeures = v: End Property
Public Property Get PiecesUtilisees() As String: PiecesUtilisees = m_PiecesUtilisees: End Property
Public Property Let PiecesUtilisees(ByVal v As String): m_PiecesUtilisees = v: End Property
Public Property Get CoutMainOeuvre() As Double: CoutMainOeuvre = m_CoutMainOeuvre: End Property
Public Property Let CoutMainOeuvre(ByVal v As Double): m_CoutMainOeuvre = v: End Property
Public Property Get CoutPieces() As Double: CoutPieces = m_CoutPieces: End Property
Public Property Let CoutPieces(ByVal v As Double): m_CoutPieces = v: End Property
Public Property Get CoutTotal() As Double: CoutTotal = m_CoutTotal: End Property
Public Property Let CoutTotal(ByVal v As Double): m_CoutTotal = v: End Property
Public Property Get Resultat() As String: Resultat = m_Resultat: End Property
Public Property Let Resultat(ByVal v As String): m_Resultat = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateIntervention = Date
    m_TypeIntervention = "Préventive"
    m_Resultat = "Résolu"
End Sub

Public Sub RecalculerCout()
    m_CoutTotal = m_CoutMainOeuvre + m_CoutPieces
End Sub

Public Function Validate() As Boolean
    If Len(m_EquipementId) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Description)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function