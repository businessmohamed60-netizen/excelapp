Attribute VB_Name = "CAnalyseRejet"
Option Explicit

' ============================================================
' CAnalyseRejet — Analyse Pareto des rejets
' ============================================================

Private m_Id As String
Private m_Periode As String             ' 2026-Q1
Private m_DateDebut As Date
Private m_DateFin As Date
Private m_TypeAnalyse As String         ' Type / Poste / Produit / Cause
Private m_Resultats As String           ' JSON : [{libelle, quantite, cout, pourcentage, cumul}]
Private m_DateCreation As Date

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Periode() As String: Periode = m_Periode: End Property
Public Property Let Periode(ByVal v As String): m_Periode = v: End Property
Public Property Get DateDebut() As Date: DateDebut = m_DateDebut: End Property
Public Property Let DateDebut(ByVal v As Date): m_DateDebut = v: End Property
Public Property Get DateFin() As Date: DateFin = m_DateFin: End Property
Public Property Let DateFin(ByVal v As Date): m_DateFin = v: End Property
Public Property Get TypeAnalyse() As String: TypeAnalyse = m_TypeAnalyse: End Property
Public Property Let TypeAnalyse(ByVal v As String): m_TypeAnalyse = v: End Property
Public Property Get Resultats() As String: Resultats = m_Resultats: End Property
Public Property Let Resultats(ByVal v As String): m_Resultats = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_TypeAnalyse = "Type"
End Sub