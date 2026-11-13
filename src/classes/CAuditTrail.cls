Attribute VB_Name = "CAuditTrail"
Option Explicit

' ============================================================
' CAuditTrail — Entrée du journal des modifications
' IATF §7.5.3.2 — maîtrise des enregistrements
' ============================================================

Private m_Id As String
Private m_DateHeure As Date
Private m_Utilisateur As String
Private m_Poste As String
Private m_Action As String              ' Création / Modification / Suppression / Signature
Private m_Collection As String          ' clients / audits / nc / capa / ...
Private m_ObjetId As String
Private m_ObjetLibelle As String        ' référence lisible (NC-2026-001, ...)
Private m_ChampsModifies As String      ' liste des champs modifiés (séparés par ;)
Private m_AncienneValeur As String      ' JSON simplifié des anciennes valeurs
Private m_NouvelleValeur As String      ' JSON simplifié des nouvelles valeurs

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get DateHeure() As Date: DateHeure = m_DateHeure: End Property
Public Property Let DateHeure(ByVal v As Date): m_DateHeure = v: End Property
Public Property Get Utilisateur() As String: Utilisateur = m_Utilisateur: End Property
Public Property Let Utilisateur(ByVal v As String): m_Utilisateur = v: End Property
Public Property Get Poste() As String: Poste = m_Poste: End Property
Public Property Let Poste(ByVal v As String): m_Poste = v: End Property
Public Property Get Action() As String: Action = m_Action: End Property
Public Property Let Action(ByVal v As String): m_Action = v: End Property
Public Property Get Collection() As String: Collection = m_Collection: End Property
Public Property Let Collection(ByVal v As String): m_Collection = v: End Property
Public Property Get ObjetId() As String: ObjetId = m_ObjetId: End Property
Public Property Let ObjetId(ByVal v As String): m_ObjetId = v: End Property
Public Property Get ObjetLibelle() As String: ObjetLibelle = m_ObjetLibelle: End Property
Public Property Let ObjetLibelle(ByVal v As String): m_ObjetLibelle = v: End Property
Public Property Get ChampsModifies() As String: ChampsModifies = m_ChampsModifies: End Property
Public Property Let ChampsModifies(ByVal v As String): m_ChampsModifies = v: End Property
Public Property Get AncienneValeur() As String: AncienneValeur = m_AncienneValeur: End Property
Public Property Let AncienneValeur(ByVal v As String): m_AncienneValeur = v: End Property
Public Property Get NouvelleValeur() As String: NouvelleValeur = m_NouvelleValeur: End Property
Public Property Let NouvelleValeur(ByVal v As String): m_NouvelleValeur = v: End Property

Public Sub InitNew(ByVal action As String, ByVal coll As String, _
                   ByVal objetId As String, ByVal libelle As String)
    m_Id = NewGuid()
    m_DateHeure = Now
    m_Utilisateur = Environ$("USERNAME")
    m_Poste = Environ$("COMPUTERNAME")
    m_Action = action
    m_Collection = coll
    m_ObjetId = objetId
    m_ObjetLibelle = libelle
End Sub

Public Function DisplayName() As String
    DisplayName = Format$(m_DateHeure, "yyyy-mm-dd hh:nn") & " | " & _
        m_Utilisateur & " | " & m_Action & " | " & m_ObjetLibelle
End Function