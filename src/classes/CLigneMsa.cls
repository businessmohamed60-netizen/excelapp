Attribute VB_Name = "CLigneMsa"
Option Explicit

' ============================================================
' CLigneMsa — Mesure individuelle d'une étude MSA
' ============================================================

Private m_Id As String
Private m_EtudeId As String
Private m_Operateur As Long             ' 1..N
Private m_Piece As Long                 ' 1..N
Private m_Essai As Long                 ' 1..N
Private m_Valeur As Double
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get EtudeId() As String: EtudeId = m_EtudeId: End Property
Public Property Let EtudeId(ByVal v As String): m_EtudeId = v: End Property
Public Property Get Operateur() As Long: Operateur = m_Operateur: End Property
Public Property Let Operateur(ByVal v As Long): m_Operateur = v: End Property
Public Property Get Piece() As Long: Piece = m_Piece: End Property
Public Property Let Piece(ByVal v As Long): m_Piece = v: End Property
Public Property Get Essai() As Long: Essai = m_Essai: End Property
Public Property Let Essai(ByVal v As Long): m_Essai = v: End Property
Public Property Get Valeur() As Double: Valeur = m_Valeur: End Property
Public Property Let Valeur(ByVal v As Double): m_Valeur = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
End Sub