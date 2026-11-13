Attribute VB_Name = "CRejetInterne"
Option Explicit

' ============================================================
' CRejetInterne — Rejet interne (IATF §8.5.6.1)
' ============================================================

Private m_Id As String
Private m_Numero As String              ' RI-2026-001
Private m_DateRejet As Date
Private m_ClientId As String
Private m_Produit As String
Private m_ReferenceProduit As String
Private m_Lot As String
Private m_Operation As String           ' poste / étape où le rejet a été détecté
Private m_TypeRejet As String           ' Dimensionnel / Aspect / Fonctionnel / Matière / Autre
Private m_Description As String
Private m_QuantiteRejetee As Long
Private m_QuantiteTotale As Long
Private m_CoutUnitaire As Double
Private m_CoutTotal As Double
Private m_CauseProbable As String
Private m_ActionImmediate As String
Private m_Statut As String              ' Ouvert / En analyse / Clôturé
Private m_NcId As String                ' lien vers NC si applicable
Private m_DetectePar As String
Private m_DateCloture As Date
Private m_Commentaire As String
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Numero() As String: Numero = m_Numero: End Property
Public Property Let Numero(ByVal v As String): m_Numero = v: End Property
Public Property Get DateRejet() As Date: DateRejet = m_DateRejet: End Property
Public Property Let DateRejet(ByVal v As Date): m_DateRejet = v: End Property
Public Property Get ClientId() As String: ClientId = m_ClientId: End Property
Public Property Let ClientId(ByVal v As String): m_ClientId = v: End Property
Public Property Get Produit() As String: Produit = m_Produit: End Property
Public Property Let Produit(ByVal v As String): m_Produit = v: End Property
Public Property Get ReferenceProduit() As String: ReferenceProduit = m_ReferenceProduit: End Property
Public Property Let ReferenceProduit(ByVal v As String): m_ReferenceProduit = v: End Property
Public Property Get Lot() As String: Lot = m_Lot: End Property
Public Property Let Lot(ByVal v As String): m_Lot = v: End Property
Public Property Get Operation() As String: Operation = m_Operation: End Property
Public Property Let Operation(ByVal v As String): m_Operation = v: End Property
Public Property Get TypeRejet() As String: TypeRejet = m_TypeRejet: End Property
Public Property Let TypeRejet(ByVal v As String): m_TypeRejet = v: End Property
Public Property Get Description() As String: Description = m_Description: End Property
Public Property Let Description(ByVal v As String): m_Description = v: End Property
Public Property Get QuantiteRejetee() As Long: QuantiteRejetee = m_QuantiteRejetee: End Property
Public Property Let QuantiteRejetee(ByVal v As Long): m_QuantiteRejetee = v: End Property
Public Property Get QuantiteTotale() As Long: QuantiteTotale = m_QuantiteTotale: End Property
Public Property Let QuantiteTotale(ByVal v As Long): m_QuantiteTotale = v: End Property
Public Property Get CoutUnitaire() As Double: CoutUnitaire = m_CoutUnitaire: End Property
Public Property Let CoutUnitaire(ByVal v As Double): m_CoutUnitaire = v: End Property
Public Property Get CoutTotal() As Double: CoutTotal = m_CoutTotal: End Property
Public Property Let CoutTotal(ByVal v As Double): m_CoutTotal = v: End Property
Public Property Get CauseProbable() As String: CauseProbable = m_CauseProbable: End Property
Public Property Let CauseProbable(ByVal v As String): m_CauseProbable = v: End Property
Public Property Get ActionImmediate() As String: ActionImmediate = m_ActionImmediate: End Property
Public Property Let ActionImmediate(ByVal v As String): m_ActionImmediate = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get NcId() As String: NcId = m_NcId: End Property
Public Property Let NcId(ByVal v As String): m_NcId = v: End Property
Public Property Get DetectePar() As String: DetectePar = m_DetectePar: End Property
Public Property Let DetectePar(ByVal v As String): m_DetectePar = v: End Property
Public Property Get DateCloture() As Date: DateCloture = m_DateCloture: End Property
Public Property Let DateCloture(ByVal v As Date): m_DateCloture = v: End Property
Public Property Get Commentaire() As String: Commentaire = m_Commentaire: End Property
Public Property Let Commentaire(ByVal v As String): m_Commentaire = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_DateRejet = Date
    m_Statut = "Ouvert"
    m_TypeRejet = "Dimensionnel"
End Sub

Public Sub AssignNumero()
    Dim coll As Object, k As Variant, yr As String, maxN As Long, n As Long
    Set coll = g_CollMgr.LoadCollection("rejetinterne")
    yr = Format$(Year(m_DateRejet), "0000")
    For Each k In coll.Keys
        If Left$(coll(k).Numero, 6) = "RI-" & yr Then
            n = CLng(Val(Mid$(coll(k).Numero, 9)))
            If n > maxN Then maxN = n
        End If
    Next k
    m_Numero = "RI-" & yr & "-" & Format$(maxN + 1, "000")
End Sub

Public Sub RecalculerCout()
    m_CoutTotal = m_QuantiteRejetee * m_CoutUnitaire
End Sub

' Taux de rebut = qté rejetée / qté totale
Public Function TauxRebut() As Double
    If m_QuantiteTotale <= 0 Then Exit Function
    TauxRebut = m_QuantiteRejetee / m_QuantiteTotale * 100
End Function

Public Function Validate() As Boolean
    If m_QuantiteRejetee <= 0 Then Validate = False: Exit Function
    If m_QuantiteTotale <= 0 Then Validate = False: Exit Function
    If m_QuantiteRejetee > m_QuantiteTotale Then Validate = False: Exit Function
    Validate = True
End Function