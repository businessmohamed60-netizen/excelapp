Option Explicit

' ============================================================
' CVersionDoc — Version d'un document (ISO 9001 §7.5.3)
' ============================================================

Private m_Id As String
Private m_DocumentId As String
Private m_NumeroVersion As String        ' 1.0, 1.1, 2.0
Private m_DateCreation As Date
Private m_Auteur As String
Private m_CheminFichier As String        ' chemin local ou UNC
Private m_HashFichier As String          ' SHA-256 pour intégrité
Private m_Commentaire As String
Private m_StatutApprobation As String    ' En attente / Approuvée / Rejetée
Private m_Approbateur As String
Private m_DateApprobation As Date
Private m_MotifRejet As String

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get DocumentId() As String: DocumentId = m_DocumentId: End Property
Public Property Let DocumentId(ByVal v As String): m_DocumentId = v: End Property
Public Property Get NumeroVersion() As String: NumeroVersion = m_NumeroVersion: End Property
Public Property Let NumeroVersion(ByVal v As String): m_NumeroVersion = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property
Public Property Get Auteur() As String: Auteur = m_Auteur: End Property
Public Property Let Auteur(ByVal v As String): m_Auteur = v: End Property
Public Property Get CheminFichier() As String: CheminFichier = m_CheminFichier: End Property
Public Property Let CheminFichier(ByVal v As String): m_CheminFichier = v: End Property
Public Property Get HashFichier() As String: HashFichier = m_HashFichier: End Property
Public Property Let HashFichier(ByVal v As String): m_HashFichier = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get StatutApprobation() As String: StatutApprobation = m_StatutApprobation: End Property
Public Property Let StatutApprobation(ByVal v As String): m_StatutApprobation = v: End Property
Public Property Get Approbateur() As String: Approbateur = m_Approbateur: End Property
Public Property Let Approbateur(ByVal v As String): m_Approbateur = v: End Property
Public Property Get DateApprobation() As Date: DateApprobation = m_DateApprobation: End Property
Public Property Let DateApprobation(ByVal v As Date): m_DateApprobation = v: End Property
Public Property Get MotifRejet() As String: MotifRejet = m_MotifRejet: End Property
Public Property Let MotifRejet(ByVal v As String): m_MotifRejet = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_NumeroVersion = "1.0"
    m_StatutApprobation = "En attente"
    m_Auteur = Environ$("USERNAME")
End Sub

Public Function EstApprouvee() As Boolean
    EstApprouvee = (m_StatutApprobation = "Approuvée")
End Function