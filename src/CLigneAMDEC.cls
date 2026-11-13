Attribute VB_Name = "CLigneAMDEC"
Option Explicit

' ============================================================
' CLigneAMDEC — Ligne d'AMDEC (mode de défaillance)
' RPN = Gravité × Occurrence × Détection (1..10 chacun)
' ============================================================

Private m_Id As String
Private m_AmdecId As String
Private m_NumeroLigne As Long
Private m_EtapeProcessus As String
Private m_Fonction As String
Private m_ModeDefaillance As String
Private m_Effet As String
Private m_Gravite As Long               ' 1..10
Private m_Cause As String
Private m_Occurrence As Long            ' 1..10
Private m_ControlesPreventifs As String
Private m_Detection As Long             ' 1..10
Private m_ControlesDetectifs As String
Private m_Rpn As Long                   ' Gravité × Occurrence × Détection
Private m_Criticite As String           ' Faible / Moyenne / Élevée / Critique
Private m_ActionRecommandee As String
Private m_ResponsableAction As String
Private m_DateEcheanceAction As Date
Private m_DateRealisationAction As Date
Private m_StatutAction As String        ' Ouverte / En cours / Clôturée
Private m_GraviteApres As Long
Private m_OccurrenceApres As Long
Private m_DetectionApres As Long
Private m_RpnApres As Long
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get AmdecId() As String: AmdecId = m_AmdecId: End Property
Public Property Let AmdecId(ByVal v As String): m_AmdecId = v: End Property
Public Property Get NumeroLigne() As Long: NumeroLigne = m_NumeroLigne: End Property
Public Property Let NumeroLigne(ByVal v As Long): m_NumeroLigne = v: End Property
Public Property Get EtapeProcessus() As String: EtapeProcessus = m_EtapeProcessus: End Property
Public Property Let EtapeProcessus(ByVal v As String): m_EtapeProcessus = v: End Property
Public Property Get Fonction() As String: Fonction = m_Fonction: End Property
Public Property Let Fonction(ByVal v As String): m_Fonction = v: End Property
Public Property Get ModeDefaillance() As String: ModeDefaillance = m_ModeDefaillance: End Property
Public Property Let ModeDefaillance(ByVal v As String): m_ModeDefaillance = v: End Property
Public Property Get Effet() As String: Effet = m_Effet: End Property
Public Property Let Effet(ByVal v As String): m_Effet = v: End Property
Public Property Get Gravite() As Long: Gravite = m_Gravite: End Property
Public Property Let Gravite(ByVal v As Long): m_Gravite = v: End Property
Public Property Get Cause() As String: Cause = m_Cause: End Property
Public Property Let Cause(ByVal v As String): m_Cause = v: End Property
Public Property Get Occurrence() As Long: Occurrence = m_Occurrence: End Property
Public Property Let Occurrence(ByVal v As Long): m_Occurrence = v: End Property
Public Property Get ControlesPreventifs() As String: ControlesPreventifs = m_ControlesPreventifs: End Property
Public Property Let ControlesPreventifs(ByVal v As String): m_ControlesPreventifs = v: End Property
Public Property Get Detection() As Long: Detection = m_Detection: End Property
Public Property Let Detection(ByVal v As Long): m_Detection = v: End Property
Public Property Get ControlesDetectifs() As String: ControlesDetectifs = m_ControlesDetectifs: End Property
Public Property Let ControlesDetectifs(ByVal v As String): m_ControlesDetectifs = v: End Property
Public Property Get Rpn() As Long: Rpn = m_Rpn: End Property
Public Property Let Rpn(ByVal v As Long): m_Rpn = v: End Property
Public Property Get Criticite() As String: Criticite = m_Criticite: End Property
Public Property Let Criticite(ByVal v As String): m_Criticite = v: End Property
Public Property Get ActionRecommandee() As String: ActionRecommandee = m_ActionRecommandee: End Property
Public Property Let ActionRecommandee(ByVal v As String): m_ActionRecommandee = v: End Property
Public Property Get ResponsableAction() As String: ResponsableAction = m_ResponsableAction: End Property
Public Property Let ResponsableAction(ByVal v As String): m_ResponsableAction = v: End Property
Public Property Get DateEcheanceAction() As Date: DateEcheanceAction = m_DateEcheanceAction: End Property
Public Property Let DateEcheanceAction(ByVal v As Date): m_DateEcheanceAction = v: End Property
Public Property Get DateRealisationAction() As Date: DateRealisationAction = m_DateRealisationAction: End Property
Public Property Let DateRealisationAction(ByVal v As Date): m_DateRealisationAction = v: End Property
Public Property Get StatutAction() As String: StatutAction = m_StatutAction: End Property
Public Property Let StatutAction(ByVal v As String): m_StatutAction = v: End Property
Public Property Get GraviteApres() As Long: GraviteApres = m_GraviteApres: End Property
Public Property Let GraviteApres(ByVal v As Long): m_GraviteApres = v: End Property
Public Property Get OccurrenceApres() As Long: OccurrenceApres = m_OccurrenceApres: End Property
Public Property Let OccurrenceApres(ByVal v As Long): m_OccurrenceApres = v: End Property
Public Property Get DetectionApres() As Long: DetectionApres = m_DetectionApres: End Property
Public Property Let DetectionApres(ByVal v As Long): m_DetectionApres = v: End Property
Public Property Get RpnApres() As Long: RpnApres = m_RpnApres: End Property
Public Property Let RpnApres(ByVal v As Long): m_RpnApres = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Gravite = 5
    m_Occurrence = 5
    m_Detection = 5
    m_StatutAction = "Ouverte"
    RecalculerRpn
End Sub

Public Sub RecalculerRpn()
    m_Rpn = m_Gravite * m_Occurrence * m_Detection
    If m_GraviteApres > 0 Or m_OccurrenceApres > 0 Or m_DetectionApres > 0 Then
        m_RpnApres = m_GraviteApres * m_OccurrenceApres * m_DetectionApres
    Else
        m_RpnApres = 0
    End If
    m_Criticite = Classifier(m_Rpn, m_Gravite)
End Sub

Private Function Classifier(ByVal rpn As Long, ByVal gravite As Long) As String
    ' Règle de criticité : priorité à la gravité élevée (IATF recommande G ≥ 9 = critique)
    If gravite >= 9 Then
        Classifier = "Critique"
    ElseIf rpn >= 200 Then
        Classifier = "Critique"
    ElseIf rpn >= 100 Then
        Classifier = "Élevée"
    ElseIf rpn >= 50 Then
        Classifier = "Moyenne"
    Else
        Classifier = "Faible"
    End If
End Function

Public Function EstEnRetard() As Boolean
    EstEnRetard = (m_StatutAction <> "Clôturée") And _
                  IsDate(m_DateEcheanceAction) And (m_DateEcheanceAction < Date)
End Function

Public Function Validate() As Boolean
    If Len(m_AmdecId) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_ModeDefaillance)) = 0 Then Validate = False: Exit Function
    If m_Gravite < 1 Or m_Gravite > 10 Then Validate = False: Exit Function
    If m_Occurrence < 1 Or m_Occurrence > 10 Then Validate = False: Exit Function
    If m_Detection < 1 Or m_Detection > 10 Then Validate = False: Exit Function
    Validate = True
End Function