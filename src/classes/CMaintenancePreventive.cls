Attribute VB_Name = "CMaintenancePreventive"
Option Explicit

' ============================================================
' CMaintenancePreventive — Plan de maintenance préventive
' ============================================================

Private m_Id As String
Private m_EquipementId As String
Private m_Designation As String
Private m_TypeMaintenance As String     ' Quotidienne / Hebdomadaire / Mensuelle / Trimestrielle / Annuelle
Private m_PeriodiciteJours As Long
Private m_Operations As String          ' liste des opérations
Private m_Responsable As String
Private m_DureeEstimeeHeures As Double
Private m_PiecesRechange As String
Private m_DateDerniere As Date
Private m_ProchaineDate As Date
Private m_Statut As String              ' Actif / Suspendu
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get EquipementId() As String: EquipementId = m_EquipementId: End Property
Public Property Let EquipementId(ByVal v As String): m_EquipementId = v: End Property
Public Property Get Designation() As String: Designation = m_Designation: End Property
Public Property Let Designation(ByVal v As String): m_Designation = v: End Property
Public Property Get TypeMaintenance() As String: TypeMaintenance = m_TypeMaintenance: End Property
Public Property Let TypeMaintenance(ByVal v As String): m_TypeMaintenance = v: End Property
Public Property Get PeriodiciteJours() As Long: PeriodiciteJours = m_PeriodiciteJours: End Property
Public Property Let PeriodiciteJours(ByVal v As Long): m_PeriodiciteJours = v: End Property
Public Property Get Operations() As String: Operations = m_Operations: End Property
Public Property Let Operations(ByVal v As String): m_Operations = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get DureeEstimeeHeures() As Double: DureeEstimeeHeures = m_DureeEstimeeHeures: End Property
Public Property Let DureeEstimeeHeures(ByVal v As Double): m_DureeEstimeeHeures = v: End Property
Public Property Get PiecesRechange() As String: PiecesRechange = m_PiecesRechange: End Property
Public Property Let PiecesRechange(ByVal v As String): m_PiecesRechange = v: End Property
Public Property Get DateDerniere() As Date: DateDerniere = m_DateDerniere: End Property
Public Property Let DateDerniere(ByVal v As Date): m_DateDerniere = v: End Property
Public Property Get ProchaineDate() As Date: ProchaineDate = m_ProchaineDate: End Property
Public Property Let ProchaineDate(ByVal v As Date): m_ProchaineDate = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "Actif"
    m_TypeMaintenance = "Mensuelle"
    m_PeriodiciteJours = 30
End Sub

Public Function EstEnRetard() As Boolean
    EstEnRetard = (m_Statut = "Actif") And IsDate(m_ProchaineDate) And m_ProchaineDate < Date
End Function

Public Function Validate() As Boolean
    If Len(m_EquipementId) = 0 Then Validate = False: Exit Function
    If m_PeriodiciteJours <= 0 Then Validate = False: Exit Function
    Validate = True
End Function