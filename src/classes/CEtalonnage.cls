Attribute VB_Name = "CEtalonnage"
Option Explicit

' ============================================================
' CEtalonnage — Enregistrement d'un étalonnage / vérification
' IATF §7.1.5.2.1
' ============================================================

Private m_Id As String
Private m_MoyenId As String
Private m_DateEtalonnage As Date
Private m_TypeOperation As String       ' Étalonnage / Vérification / Réparation
Private m_Organisme As String           ' Interne / Externe (nom)
Private m_CertificatNumero As String
Private m_Resultat As String            ' Conforme / Non conforme / Conforme avec réserve
Private m_IncertitudeRelevee As String
Private m_ProchaineEcheance As Date
Private m_Cout As Double
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get MoyenId() As String: MoyenId = m_MoyenId: End Property
Public Property Let MoyenId(ByVal v As String): m_MoyenId = v: End Property
Public Property Get DateEtalonnage() As Date: DateEtalonnage = m_DateEtalonnage: End Property
Public Property Let DateEtalonnage(ByVal v As Date): m_DateEtalonnage = v: End Property
Public Property Get TypeOperation() As String: TypeOperation = m_TypeOperation: End Property
Public Property Let TypeOperation(ByVal v As String): m_TypeOperation = v: End Property
Public Property Get Organisme() As String: Organisme = m_Organisme: End Property
Public Property Let Organisme(ByVal v As String): m_Organisme = v: End Property
Public Property Get CertificatNumero() As String: CertificatNumero = m_CertificatNumero: End Property
Public Property Let CertificatNumero(ByVal v As String): m_CertificatNumero = v: End Property
Public Property Get Resultat() As String: Resultat = m_Resultat: End Property
Public Property Let Resultat(ByVal v As String): m_Resultat = v: End Property
Public Property Get IncertitudeRelevee() As String: IncertitudeRelevee = m_IncertitudeRelevee: End Property
Public Property Let IncertitudeRelevee(ByVal v As String): m_IncertitudeRelevee = v: End Property
Public Property Get ProchaineEcheance() As Date: ProchaineEcheance = m_ProchaineEcheance: End Property
Public Property Let ProchaineEcheance(ByVal v As Date): m_ProchaineEcheance = v: End Property
Public Property Get Cout() As Double: Cout = m_Cout: End Property
Public Property Let Cout(ByVal v As Double): m_Cout = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateEtalonnage = Date
    m_TypeOperation = "Étalonnage"
    m_Resultat = "Conforme"
End Sub

Public Function Validate() As Boolean
    If Len(m_MoyenId) = 0 Then Validate = False: Exit Function
    If Not IsDate(m_DateEtalonnage) Then Validate = False: Exit Function
    Validate = True
End Function