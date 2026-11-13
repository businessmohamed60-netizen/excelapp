Attribute vb_Name = "CLigneSurveillance"
Option Explicit

' ============================================================
' CLigneSurveillance — Ligne de plan de surveillance
' ============================================================

Private m_Id As String
Private m_PlanId As String
Private m_NumeroLigne As Long
Private m_Caracteristique As String
Private m_TypeCaracteristique As String ' CC / SC / Normale
Private m_Specification As String
Private m_Tolerance As String
Private m_Frequence As String
Private m_TailleEchantillon As Long
Private m_MoyenControle As String
Private m_Enregistrement As String
Private m_ReactionNC As String
Private m_StatutCar As String           ' OK / Alerte / Hors tolérance
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get PlanId() As String: PlanId = m_PlanId: End Property
Public Property Let PlanId(ByVal v As String): m_PlanId = v: End Property
Public Property Get NumeroLigne() As Long: NumeroLigne = m_NumeroLigne: End Property
Public Property Let NumeroLigne(ByVal v As Long): m_NumeroLigne = v: End Property
Public Property Get Caracteristique() As String: Caracteristique = m_Caracteristique: End Property
Public Property Let Caracteristique(ByVal v As String): m_Caracteristique = v: End Property
Public Property Get TypeCaracteristique() As String: TypeCaracteristique = m_TypeCaracteristique: End Property
Public Property Let TypeCaracteristique(ByVal v As String): m_TypeCaracteristique = v: End Property
Public Property Get Specification() As String: Specification = m_Specification: End Property
Public Property Let Specification(ByVal v As String): m_Specification = v: End Property
Public Property Get Tolerance() As String: Tolerance = m_Tolerance: End Property
Public Property Let Tolerance(ByVal v As String): m_Tolerance = v: End Property
Public Property Get Frequence() As String: Frequence = m_Frequence: End Property
Public Property Let Frequence(ByVal v As String): m_Frequence = v: End Property
Public Property Get TailleEchantillon() As Long: TailleEchantillon = m_TailleEchantillon: End Property
Public Property Let TailleEchantillon(ByVal v As Long): m_TailleEchantillon = v: End Property
Public Property Get MoyenControle() As String: MoyenControle = m_MoyenControle: End Property
Public Property Let MoyenControle(ByVal v As String): m_MoyenControle = v: End Property
Public Property Get Enregistrement() As String: Enregistrement = m_Enregistrement: End Property
Public Property Let Enregistrement(ByVal v As String): m_Enregistrement = v: End Property
Public Property Get ReactionNC() As String: ReactionNC = m_ReactionNC: End Property
Public Property Let ReactionNC(ByVal v As String): m_ReactionNC = v: End Property
Public Property Get StatutCar() As String: StatutCar = m_StatutCar: End Property
Public Property Let StatutCar(ByVal v As String): m_StatutCar = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_TypeCaracteristique = "Normale"
    m_TailleEchantillon = 5
    m_StatutCar = "OK"
End Sub

Public Function Validate() As Boolean
    If Len(m_PlanId) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Caracteristique)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function