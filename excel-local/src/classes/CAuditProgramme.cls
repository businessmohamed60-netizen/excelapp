Attribute VB_Name = "CAuditProgramme"
Option Explicit

' ============================================================
' CAuditProgramme — Ligne du programme annuel d'audits
' ISO 9001 §9.2.2 / IATF §9.2.2.1
' ============================================================

Private m_Id As String
Private m_Annee As Long
Private m_ClientId As String
Private m_Processus As String
Private m_TypeAudit As String          ' Processus / Système / Produit / Fournisseur
Private m_Risque As String             ' Faible / Moyen / Élevé / Critique
Private m_FrequenceMois As Long        ' 3, 6, 12, 24
Private m_DatePrevue As Date
Private m_DateRealisee As Date
Private m_AuditId As String            ' lien vers CAudit si réalisé
Private m_Auditeur As String
Private m_Statut As String             ' Planifié / Réalisé / Reporté / Annulé
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Annee() As Long: Annee = m_Annee: End Property
Public Property Let Annee(ByVal v As Long): m_Annee = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Processus() As String: Processus = m_Processus: End Property
Public Property Let Processus(ByVal v As String): m_Processus = v: End Property
Public Property Get TypeAudit() As String: TypeAudit = m_TypeAudit: End Property
Public Property Let TypeAudit(ByVal v As String): m_TypeAudit = v: End Property
Public Property Get Risque() As String: Risque = m_Risque: End Property
Public Property Let Risque(ByVal v As String): m_Risque = v: End Property
Public Property Get FrequenceMois() As Long: FrequenceMois = m_FrequenceMois: End Property
Public Property Let FrequenceMois(ByVal v As Long): m_FrequenceMois = v: End Property
Public Property Get DatePrevue() As Date: DatePrevue = m_DatePrevue: End Property
Public Property Let DatePrevue(ByVal v As Date): m_DatePrevue = v: End Property
Public Property Get DateRealisee() As Date: DateRealisee = m_DateRealisee: End Property
Public Property Let DateRealisee(ByVal v As Date): m_DateRealisee = v: End Property
Public Property Get AuditId() As String: AuditId = m_AuditId: End Property
Public Property Let AuditId(ByVal v As String): m_AuditId = v: End Property
Public Property Get Auditeur() As String: Auditeur = m_Auditeur: End Property
Public Property Let Auditeur(ByVal v As String): m_Auditeur = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Annee = Year(Date)
    m_Statut = "Planifié"
    m_TypeAudit = "Processus"
    m_Risque = "Moyen"
    m_FrequenceMois = 12
End Sub

Public Function EstEnRetard() As Boolean
    EstEnRetard = (m_Statut = "Planifié") And (m_DatePrevue < Date)
End Function

Public Function EstEchu() As Boolean
    EstEchu = (m_Statut = "Planifié") And (m_DatePrevue <= Date + 30)
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_Processus)) = 0 Then Validate = False: Exit Function
    If m_Annee < 2000 Or m_Annee > 2100 Then Validate = False: Exit Function
    If Not IsDate(m_DatePrevue) Then Validate = False: Exit Function
    If m_Statut = "Réalisé" And Not IsDate(m_DateRealisee) Then
        Validate = False: Exit Function
    End If
    Validate = True
End Function