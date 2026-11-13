Option Explicit

' ============================================================
' CDocument — Document qualité (ISO 9001 §7.5 / IATF §7.5.3)
' ============================================================

Private m_Id As String
Private m_Code As String                 ' PR-QSE-001, IN-PROD-012, FO-ACH-003
Private m_Titre As String
Private m_TypeDoc As String              ' Procédure / Instruction / Formulaire / Enregistrement / Politique / Manuel
Private m_Processus As String
Private m_Proprietaire As String         ' Responsable du document
Private m_Statut As String               ' Brouillon / En revue / Approuvé / Publié / Obsolète / Archivé
Private m_VersionCouranteId As String
Private m_PeriodiciteRevueMois As Long   ' 0 = pas de revue, sinon 6/12/24/36
Private m_DateDerniereRevue As Date
Private m_DateProchaineRevue As Date
Private m_MotsCles As String             ' séparés par des virgules
Private m_Description As String
Private m_DateCreation As Date
Private m_DatePublication As Date
Private m_DateArchivage As Date
Private m_EstCritique As Boolean         ' document critique pour la sécurité/qualité

Public Property Get Id() As String: Id = m_Id: End Property
Public Property Let Id(ByVal v As String): m_Id = v: End Property
Public Property Get Code() As String: Code = m_Code: End Property
Public Property Let Code(ByVal v As String): m_Code = v: End Property
Public Property Get Titre() As String: Titre = m_Titre: End Property
Public Property Let Titre(ByVal v As String): m_Titre = v: End Property
Public Property Get TypeDoc() As String: TypeDoc = m_TypeDoc: End Property
Public Property Let TypeDoc(ByVal v As String): m_TypeDoc = v: End Property
Public Property Get Processus() As String: Processus = m_Processus: End Property
Public Property Let Processus(ByVal v As String): m_Processus = v: End Property
Public Property Get Proprietaire() As String: Proprietaire = m_Proprietaire: End Property
Public Property Let Proprietaire(ByVal v As String): m_Proprietaire = v: End Property
Public Property Get Statut() As String: Statut = m_Statut: End Property
Public Property Let Statut(ByVal v As String): m_Statut = v: End Property
Public Property Get VersionCouranteId() As String: VersionCouranteId = m_VersionCouranteId: End Property
Public Property Let VersionCouranteId(ByVal v As String): m_VersionCouranteId = v: End Property
Public Property Get PeriodiciteRevueMois() As Long: PeriodiciteRevueMois = m_PeriodiciteRevueMois: End Property
Public Property Let PeriodiciteRevueMois(ByVal v As Long): m_PeriodiciteRevueMois = v: End Property
Public Property Get DateDerniereRevue() As Date: DateDerniereRevue = m_DateDerniereRevue: End Property
Public Property Let DateDerniereRevue(ByVal v As Date): m_DateDerniereRevue = v: End Property
Public Property Get DateProchaineRevue() As Date: DateProchaineRevue = m_DateProchaineRevue: End Property
Public Property Let DateProchaineRevue(ByVal v As Date): m_DateProchaineRevue = v: End Property
Public Property Get MotsCles() As String: MotsCles = m_MotsCles: End Property
Public Property Let MotsCles(ByVal v As String): m_MotsCles = v: End Property
Public Property Get Description() As String: Description = m_Description: End Property
Public Property Let Description(ByVal v As String): m_Description = v: End Property
Public Property Get DateCreation() As Date: DateCreation = m_DateCreation: End Property
Public Property Let DateCreation(ByVal v As Date): m_DateCreation = v: End Property
Public Property Get DatePublication() As Date: DatePublication = m_DatePublication: End Property
Public Property Let DatePublication(ByVal v As Date): m_DatePublication = v: End Property
Public Property Get DateArchivage() As Date: DateArchivage = m_DateArchivage: End Property
Public Property Let DateArchivage(ByVal v As Date): m_DateArchivage = v: End Property
Public Property Get EstCritique() As Boolean: EstCritique = m_EstCritique: End Property
Public Property Let EstCritique(ByVal v As Boolean): m_EstCritique = v: End Property

Public Sub InitNew()
    m_Id = NewGuid()
    m_DateCreation = Now
    m_Statut = "Brouillon"
    m_TypeDoc = "Procédure"
    m_PeriodiciteRevueMois = 12
End Sub

' ---- Numérotation automatique du code ----
' Format : <Préfixe>-<Processus3>-<NNN>
' Préfixes : PR=Procédure, IN=Instruction, FO=Formulaire, EN=Enregistrement, PO=Politique, MA=Manuel
Public Sub AssignCode()
    Dim prefixe As String
    Select Case m_TypeDoc
        Case "Procédure": prefixe = "PR"
        Case "Instruction": prefixe = "IN"
        Case "Formulaire": prefixe = "FO"
        Case "Enregistrement": prefixe = "EN"
        Case "Politique": prefixe = "PO"
        Case "Manuel": prefixe = "MA"
        Case Else: prefixe = "DC"
    End Select
    Dim proc As String
    proc = UCase$(Left$(m_Processus, 3))
    If Len(proc) < 3 Then proc = proc & String$(3 - Len(proc), "X")

    Dim dom As Object, root As Object, yr As Long
    Dim fname As String, item As Object, codeNode As Object
    Dim maxN As Long, parts() As String, n As Long
    fname = DataPath() & "documents.xml"
    If Dir(fname) <> "" Then
        Set dom = CreateObject("MSXML2.DOMDocument.6.0")
        dom.async = False
        If dom.Load(fname) Then
            Set root = dom.documentElement
            For Each item In root.ChildNodes
                Set codeNode = item.selectSingleNode("Code")
                If Not codeNode Is Nothing Then
                    parts = Split(codeNode.Text, "-")
                    If UBound(parts) >= 2 Then
                        If parts(0) = prefixe And parts(1) = proc Then
                            n = CLng(Val(parts(2)))
                            If n > maxN Then maxN = n
                        End If
                    End If
                End If
            Next item
        End If
    End If
    m_Code = prefixe & "-" & proc & "-" & Format$(maxN + 1, "000")
End Sub

' ---- Calcul de la prochaine revue ----
Public Function CalculerProchaineRevue() As Date
    If m_PeriodiciteRevueMois <= 0 Then Exit Function
    Dim base As Date
    If m_DateDerniereRevue > 0 Then base = m_DateDerniereRevue Else base = m_DateCreation
    m_DateProchaineRevue = DateAdd("m", m_PeriodiciteRevueMois, base)
    CalculerProchaineRevue = m_DateProchaineRevue
End Function

' ---- Jours avant la prochaine revue (négatif si en retard) ----
Public Function JoursAvantRevue() As Long
    If m_DateProchaineRevue = 0 Then Exit Function
    JoursAvantRevue = DateDiff("d", Date, m_DateProchaineRevue)
End Function

' ---- Indicateur de statut de revue ----
Public Function StatutRevue() As String
    If m_PeriodiciteRevueMois <= 0 Then
        StatutRevue = "—"
        Exit Function
    End If
    Dim j As Long: j = JoursAvantRevue()
    If j < 0 Then
        StatutRevue = "🔴 En retard (" & (-j) & " j)"
    ElseIf j <= 30 Then
        StatutRevue = "🟠 À revoir (" & j & " j)"
    Else
        StatutRevue = "🟢 OK (" & j & " j)"
    End If
End Function

' ---- Affichage ----
Public Function DisplayName() As String
    DisplayName = m_Code & " — " & m_Titre & " [" & m_Statut & "]"
End Function

' ---- Validation ----
Public Function Validate() As Boolean
    If Len(Trim$(m_Code)) = 0 Then Exit Function
    If Len(Trim$(m_Titre)) = 0 Then Exit Function
    If Len(Trim$(m_Proprietaire)) = 0 Then Exit Function
    Validate = True
End Function

' ---- Cycle de vie ----
Public Function PeutPasserA(ByVal nouveauStatut As String) As Boolean
    Select Case m_Statut
        Case "Brouillon": PeutPasserA = (nouveauStatut = "En revue")
        Case "En revue": PeutPasserA = (nouveauStatut = "Approuvé" Or nouveauStatut = "Brouillon")
        Case "Approuvé": PeutPasserA = (nouveauStatut = "Publié" Or nouveauStatut = "En revue")
        Case "Publié": PeutPasserA = (nouveauStatut = "Obsolète" Or nouveauStatut = "En revue")
        Case "Obsolète": PeutPasserA = (nouveauStatut = "Archivé" Or nouveauStatut = "Publié")
        Case "Archivé": PeutPasserA = False
    End Select
End Function