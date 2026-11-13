Attribute VB_Name = "CEquipement"
Option Explicit

' ============================================================
' CEquipement — Équipement de production (IATF §8.5.1.5)
' ============================================================

Private m_Id As String
Private m_Code As String                ' EQ-001
Private m_Designation As String
Private m_TypeEquipement As String      ' Machine / Presse / Four / Convoyeur / Robot / Autre
Private m_Marque As String
Private m_Modele As String
Private m_NumeroSerie As String
Private m_Localisation As String
Private m_DateMiseEnService As Date
Private m_ValeurAcquisition As Double
Private m_Critique As Boolean           ' équipement critique (arrêt = arrêt production)
Private m_Statut As String              ' En service / En panne / En maintenance / Réformé
Private m_Responsable As String
Private m_TempsOuverturePlanifie As Double  ' heures/mois
Private m_TempsArretPlanifie As Double      ' heures/mois (maintenance planifiée)
Private m_TempsArretNonPlanifie As Double   ' heures/mois (pannes)
Private m_TempsProduction As Double         ' heures/mois
Private m_PiecesProduites As Long
Private m_PiecesBonnes As Long
Private m_DateDerniereMaintenance As Date
Private m_ProchaineMaintenance As Date
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Code() As String: Code = m_Code: End Property
Public Property Let Code(ByVal v As String): m_Code = v: End Property
Public Property Get Designation() As String: Designation = m_Designation: End Property
Public Property Let Designation(ByVal v As String): m_Designation = v: End Property
Public Property Get TypeEquipement() As String: TypeEquipement = m_TypeEquipement: End Property
Public Property Let TypeEquipement(ByVal v As String): m_TypeEquipement = v: End Property
Public Property Get Marque() As String: Marque = m_Marque: End Property
Public Property Let Marque(ByVal v As String): m_Marque = v: End Property
Public Property Get Modele() As String: Modele = m_Modele: End Property
Public Property Let Modele(ByVal v As String): m_Modele = v: End Property
Public Property Get NumeroSerie() As String: NumeroSerie = m_NumeroSerie: End Property
Public Property Let NumeroSerie(ByVal v As String): m_NumeroSerie = v: End Property
Public Property Get Localisation() As String: Localisation = m_Localisation: End Property
Public Property Let Localisation(ByVal v As String): m_Localisation = v: End Property
Public Property Get DateMiseEnService() As Date: DateMiseEnService = m_DateMiseEnService: End Property
Public Property Let DateMiseEnService(ByVal v As Date): m_DateMiseEnService = v: End Property
Public Property Get ValeurAcquisition() As Double: ValeurAcquisition = m_ValeurAcquisition: End Property
Public Property Let ValeurAcquisition(ByVal v As Double): m_ValeurAcquisition = v: End Property
Public Property Get Critique() As Boolean: Critique = m_Critique: End Property
Public Property Let Critique(ByVal v As Boolean): m_Critique = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Responsable() As String: Responsable = m_Responsable: End Property
Public Property Let Responsable(ByVal v As String): m_Responsable = v: End Property
Public Property Get TempsOuverturePlanifie() As Double: TempsOuverturePlanifie = m_TempsOuverturePlanifie: End Property
Public Property Let TempsOuverturePlanifie(ByVal v As Double): m_TempsOuverturePlanifie = v: End Property
Public Property Get TempsArretPlanifie() As Double: TempsArretPlanifie = m_TempsArretPlanifie: End Property
Public Property Let TempsArretPlanifie(ByVal v As Double): m_TempsArretPlanifie = v: End Property
Public Property Get TempsArretNonPlanifie() As Double: TempsArretNonPlanifie = m_TempsArretNonPlanifie: End Property
Public Property Let TempsArretNonPlanifie(ByVal v As Double): m_TempsArretNonPlanifie = v: End Property
Public Property Get TempsProduction() As Double: TempsProduction = m_TempsProduction: End Property
Public Property Let TempsProduction(ByVal v As Double): m_TempsProduction = v: End Property
Public Property Get PiecesProduites() As Long: PiecesProduites = m_PiecesProduites: End Property
Public Property Let PiecesProduites(ByVal v As Long): m_PiecesProduites = v: End Property
Public Property Get PiecesBonnes() As Long: PiecesBonnes = m_PiecesBonnes: End Property
Public Property Let PiecesBonnes(ByVal v As Long): m_PiecesBonnes = v: End Property
Public Property Get DateDerniereMaintenance() As Date: DateDerniereMaintenance = m_DateDerniereMaintenance: End Property
Public Property Let DateDerniereMaintenance(ByVal v As Date): m_DateDerniereMaintenance = v: End Property
Public Property Get ProchaineMaintenance() As Date: ProchaineMaintenance = m_ProchaineMaintenance: End Property
Public Property Let ProchaineMaintenance(ByVal v As Date): m_ProchaineMaintenance = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "En service"
    m_TypeEquipement = "Machine"
    m_Critique = False
End Sub

Public Sub AssignCode()
    Dim coll As Object, k As Variant, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("equipement")
    For Each k In coll.Keys
        n = CLng(Val(Mid$(coll(k).Code, 4)))
        If n > maxN Then maxN = n
    Next k
    m_Code = "EQ-" & Format$(maxN + 1, "000")
End Sub

' ----- Indicateurs TPM -----

' Disponibilité = Temps de fonctionnement / Temps d'ouverture planifié
Public Function Disponibilite() As Double
    If m_TempsOuverturePlanifie <= 0 Then Exit Function
    Dim tempsFonctionnement As Double
    tempsFonctionnement = m_TempsOuverturePlanifie - m_TempsArretPlanifie - m_TempsArretNonPlanifie
    Disponibilite = tempsFonctionnement / m_TempsOuverturePlanifie * 100
End Function

' Performance = (Pièces produites × Temps cycle théorique) / Temps de fonctionnement
Public Function Performance() As Double
    If m_TempsProduction <= 0 Or m_PiecesProduites <= 0 Then Exit Function
    ' Approximation : temps cycle = temps production / pièces produites
    Performance = 100 ' à affiner avec temps cycle théorique saisi
End Function

' Qualité = Pièces bonnes / Pièces produites
Public Function Qualite() As Double
    If m_PiecesProduites <= 0 Then Exit Function
    Qualite = m_PiecesBonnes / m_PiecesProduites * 100
End Function

' TRS / OEE = Disponibilité × Performance × Qualité / 10000
Public Function TRS() As Double
    TRS = Disponibilite() * Performance() * Qualite() / 10000
End Function

Public Function EstMaintenanceEchue() As Boolean
    EstMaintenanceEchue = (m_Statut = "En service") And IsDate(m_ProchaineMaintenance) And _
                          m_ProchaineMaintenance < Date
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_Designation)) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_TypeEquipement)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function

Public Function DisplayName() As String
    DisplayName = m_Code & " — " & m_Designation
End Function