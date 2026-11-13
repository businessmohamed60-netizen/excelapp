Attribute VB_Name = "CMoyenMesure"
Option Explicit

' ============================================================
' CMoyenMesure — Moyen de mesure (inventaire, étalonnage)
' IATF §7.1.5.1.1 / §7.1.5.2.1
' ============================================================

Private m_Id As String
Private m_Code As String                ' MM-001
Private m_Designation As String
Private m_TypeMoyen As String           ' Pied à coulisse / Micromètre / Comparateur / Balance / Thermo / Autre
Private m_Marque As String
Private m_Modele As String
Private m_NumeroSerie As String
Private m_Localisation As String
Private m_ClientId As String            ' optionnel
Private m_EtendueMesure As String       ' 0-150 mm
Private m_Resolution As String          ' 0.01 mm
Private m_Incertitude As String         ' ±0.02 mm
Private m_PeriodiciteMois As Long       ' 12
Private m_DateMiseEnService As Date
Private m_DernierEtalonnage As Date
Private m_ProchainEtalonnage As Date
Private m_Statut As String              ' En service / Hors service / En étalonnage / Réformé
Private m_Responsable As String
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Code() As String: Code = m_Code: End Property
Public Property Let Code(ByVal v As String): m_Code = v: End Property
Public Property Get Designation() As String: Designation = m_Designation: End Property
Public Property Let Designation(ByVal v As String): m_Designation = v: End Property
Public Property Get TypeMoyen() As String: TypeMoyen = m_TypeMoyen: End Property
Public Property Let TypeMoyen(ByVal v As String): m_TypeMoyen = v: End Property
Public Property Get Marque() As String: Marque = m_Marque: End Property
Public Property Let Marque(ByVal v As String): m_Marque = v: End Property
Public Property Get Modele() As String: Modele = m_Modele: End Property
Public Property Let Modele(ByVal v As String): m_Modele = v: End Property
Public Property Get NumeroSerie() As String: NumeroSerie = m_NumeroSerie: End Property
Public Property Let NumeroSerie(ByVal v As String): m_NumeroSerie = v: End Property
Public Property Get Localisation() As String: Localisation = m_Localisation: End Property
Public Property Let Localisation(ByVal v As String): m_Localisation = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get EtendueMesure() As String: EtendueMesure = m_EtendueMesure: End Property
Public Property Let EtendueMesure(ByVal v As String): m_EtendueMesure = v: End Property
Public Property Get Resolution() As String: Resolution = m_Resolution: End Property
Public Property Let Resolution(ByVal v As String): m_Resolution = v: End Property
Public Property Get Incertitude() As String: Incertitude = m_Incertitude: End Property
Public Property Let Incertitude(ByVal v As String): m_Incertitude = v: End Property
Public Property Get PeriodiciteMois() As Long: PeriodiciteMois = m_PeriodiciteMois: End Property
Public Property Let PeriodiciteMois(ByVal v As Long): m_PeriodiciteMois = v: End Property
Public Property Get DateMiseEnService() As Date: DateMiseEnService = m_DateMiseEnService: End Property
Public Property Let DateMiseEnService(ByVal v As Date): m_DateMiseEnService = v: End Property
Public Property Get DernierEtalonnage() As Date: DernierEtalonnage = m_DernierEtalonnage: End Property
Public Property Let DernierEtalonnage(ByVal v As Date): m_DernierEtalonnage = v: End Property
Public Property Get ProchainEtalonnage() As Date: ProchainEtalonnage = m_ProchainEtalonnage: End Property
Public Property Let ProchainEtalonnage(ByVal v As Date): m_ProchainEtalonnage = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "En service"
    m_PeriodiciteMois = 12
    m_DateMiseEnService = Date
End Sub

Public Sub AssignCode()
    Dim coll As Object, k As Variant, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("moyenmesure")
    For Each k In coll.Keys
        n = CLng(Val(Mid$(coll(k).Code, 4)))
        If n > maxN Then maxN = n
    Next k
    m_Code = "MM-" & Format$(maxN + 1, "000")
End Sub

Public Sub RecalculerProchaineEcheance()
    If IsDate(m_DernierEtalonnage) And m_DernierEtalonnage > 0 And m_PeriodiciteMois > 0 Then
        m_ProchainEtalonnage = DateAdd("m", m_PeriodiciteMois, m_DernierEtalonnage)
    End If
End Sub

Public Function EstEtalonnageEchu() As Boolean
    EstEtalonnageEchu = (m_Statut = "En service") And _
                        IsDate(m_ProchainEtalonnage) And (m_ProchainEtalonnage < Date)
End Function

Public Function EtalonnageProche() As Boolean
    EtalonnageProche = (m_Statut = "En service") And _
                       IsDate(m_ProchainEtalonnage) And _
                       (m_ProchainEtalonnage >= Date) And (m_ProchainEtalonnage <= Date + 30)
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_Designation)) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_TypeMoyen)) = 0 Then Validate = False: Exit Function
    If m_PeriodiciteMois <= 0 Then Validate = False: Exit Function
    Validate = True
End Function

Public Function DisplayName() As String
    DisplayName = m_Code & " — " & m_Designation
End Function