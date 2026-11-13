Attribute VB_Name = "CFournisseur"
Option Explicit

' ============================================================
' CFournisseur — Fournisseur (IATF §8.4.2.4)
' ============================================================

Private m_Id As String
Private m_Code As String                ' FOU-001
Private m_RaisonSociale As String
Private m_Siret As String
Private m_Adresse As String
Private m_Contact As String
Private m_Telephone As String
Private m_Email As String
Private m_Categorie As String           ' Matière / Composant / Sous-traitance / Service / Équipement
Private m_Critique As Boolean           ' fournisseur critique (composants de sécurité)
Private m_Statut As String              ' Actif / En développement / Suspendu / Réformé
Private m_NoteGlobale As Double         ' 0-100, calculée
Private m_NoteQualite As Double
Private m_NoteLivraison As Double
Private m_NoteReactivite As Double
Private m_NotePrix As Double
Private m_DateDerniereEvaluation As Date
Private m_ProchaineEvaluation As Date
Private m_EstCertifie As Boolean        ' IATF 16949 ou ISO 9001
Private m_Certification As String       ' IATF 16949 / ISO 9001 / ISO 14001 / Autre
Private m_DateCertification As Date
Private m_DateExpirationCertification As Date
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Code() As String: Code = m_Code: End Property
Public Property Let Code(ByVal v As String): m_Code = v: End Property
Public Property Get RaisonSociale() As String: RaisonSociale = m_RaisonSociale: End Property
Public Property Let RaisonSociale(ByVal v As String): m_RaisonSociale = v: End Property
Public Property Get Siret() As String: Siret = m_Siret: End Property
Public Property Let Siret(ByVal v As String): m_Siret = v: End Property
Public Property Get Adresse() As String: Adresse = m_Adresse: End Property
Public Property Let Adresse(ByVal v As String): m_Adresse = v: End Property
Public Property Get Contact() As String: Contact = m_Contact: End Property
Public Property Let Contact(ByVal v As String): m_Contact = v: End Property
Public Property Get Telephone() As String: Telephone = m_Telephone: End Property
Public Property Let Telephone(ByVal v As String): m_Telephone = v: End Property
Public Property Get Email() As String: Email = m_Email: End Property
Public Property Let Email(ByVal v As String): m_Email = v: End Property
Public Property Get Categorie() As String: Categorie = m_Categorie: End Property
Public Property Let Categorie(ByVal v As String): m_Categorie = v: End Property
Public Property Get Critique() As Boolean: Critique = m_Critique: End Property
Public Property Let Critique(ByVal v As Boolean): m_Critique = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get NoteGlobale() As Double: NoteGlobale = m_NoteGlobale: End Property
Public Property Let NoteGlobale(ByVal v As Double): m_NoteGlobale = v: End Property
Public Property Get NoteQualite() As Double: NoteQualite = m_NoteQualite: End Property
Public Property Let NoteQualite(ByVal v As Double): m_NoteQualite = v: End Property
Public Property Get NoteLivraison() As Double: NoteLivraison = m_NoteLivraison: End Property
Public Property Let NoteLivraison(ByVal v As Double): m_NoteLivraison = v: End Property
Public Property Get NoteReactivite() As Double: NoteReactivite = m_NoteReactivite: End Property
Public Property Let NoteReactivite(ByVal v As Double): m_NoteReactivite = v: End Property
Public Property Get NotePrix() As Double: NotePrix = m_NotePrix: End Property
Public Property Let NotePrix(ByVal v As Double): m_NotePrix = v: End Property
Public Property Get DateDerniereEvaluation() As Date: DateDerniereEvaluation = m_DateDerniereEvaluation: End Property
Public Property Let DateDerniereEvaluation(ByVal v As Date): m_DateDerniereEvaluation = v: End Property
Public Property Get ProchaineEvaluation() As Date: ProchaineEvaluation = m_ProchaineEvaluation: End Property
Public Property Let ProchaineEvaluation(ByVal v As Date): m_ProchaineEvaluation = v: End Property
Public Property Get EstCertifie() As Boolean: EstCertifie = m_EstCertifie: End Property
Public Property Let EstCertifie(ByVal v As Boolean): m_EstCertifie = v: End Property
Public Property Get Certification() As String: Certification = m_Certification: End Property
Public Property Let Certification(ByVal v As String): m_Certification = v: End Property
Public Property Get DateCertification() As Date: DateCertification = m_DateCertification: End Property
Public Property Let DateCertification(ByVal v As Date): m_DateCertification = v: End Property
Public Property Get DateExpirationCertification() As Date: DateExpirationCertification = m_DateExpirationCertification: End Property
Public Property Let DateExpirationCertification(ByVal v As Date): m_DateExpirationCertification = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "Actif"
    m_Categorie = "Composant"
    m_Critique = False
    m_EstCertifie = False
End Sub

Public Sub AssignCode()
    Dim coll As Object, k As Variant, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("fournisseur")
    For Each k In coll.Keys
        n = CLng(Val(Mid$(coll(k).Code, 5)))
        If n > maxN Then maxN = n
    Next k
    m_Code = "FOU-" & Format$(maxN + 1, "000")
End Sub

' Recalcule la note globale (pondération qualité 40%, livraison 30%, réactivité 20%, prix 10%)
Public Sub RecalculerNoteGlobale()
    m_NoteGlobale = m_NoteQualite * 0.4 + m_NoteLivraison * 0.3 + _
                    m_NoteReactivite * 0.2 + m_NotePrix * 0.1
End Sub

Public Function EstCertificationExpiree() As Boolean
    EstCertificationExpiree = m_EstCertifie And IsDate(m_DateExpirationCertification) And _
                              m_DateExpirationCertification < Date
End Function

Public Function EstEvaluationEchue() As Boolean
    EstEvaluationEchue = (m_Statut = "Actif") And IsDate(m_ProchaineEvaluation) And _
                         m_ProchaineEvaluation < Date
End Function

Public Function Validate() As Boolean
    If Len(Trim$(m_RaisonSociale)) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Categorie)) = 0 Then Validate = False: Exit Function
    Validate = True
End Function

Public Function DisplayName() As String
    DisplayName = m_Code & " — " & m_RaisonSociale
End Function