Option Explicit

Public Enum NotifSeverite
    NotifInfo = 0
    NotifAvertissement = 1
    NotifCritique = 2
End Enum

Private Function CreerNotif(ByVal sev As NotifSeverite, ByVal categorie As String, _
                            ByVal titre As String, ByVal detail As String, _
                            ByVal action As String, ByVal objetId As String) As Object
    Dim n As Object
    Set n = CreateObject("Scripting.Dictionary")
    n("Severite") = sev
    n("Categorie") = categorie
    n("Titre") = titre
    n("Detail") = detail
    n("Action") = action
    n("ObjetId") = objetId
    n("Date") = Now
    Set CreerNotif = n
End Function

Public Function CollecterNotifications() As Collection
    Dim notifs As New Collection
    On Error GoTo ErrHandler
    CollecterAuditsAVenir notifs
    CollecterCapaEnRetard notifs
    CollecterNcMajeuresOuvertes notifs
    CollecterActionsAmdecH notifs
    CollecterPpapASoumettre notifs
    CollecterApqpEnDerive notifs
    CollecterMsaARenouveler notifs
    Set CollecterNotifications = notifs
    Exit Function
ErrHandler:
    LogEx "M_Notifications", "CollecterNotifications", Err
    Set CollecterNotifications = notifs
End Function

Private Sub CollecterAuditsAVenir(ByRef notifs As Collection)
    Dim audits As Object, k As Variant, jours As Long
    Set audits = g_CollMgr.LoadCollection("audits")
    For Each k In audits.Keys
        If audits(k).Statut = "Planifié" And audits(k).DateDebut > 0 Then
            jours = DateDiff("d", Date, audits(k).DateDebut)
            If jours >= 0 And jours <= 7 Then
                notifs.Add CreerNotif(NotifAvertissement, "Audit", _
                    "Audit dans " & jours & " jour(s)", _
                    "Processus : " & audits(k).Processus & " — Auditeur : " & audits(k).Auditeur, _
                    "OuvrirAudit", audits(k).Id)
            End If
        End If
    Next k
End Sub

Private Sub CollecterCapaEnRetard(ByRef notifs As Collection)
    Dim capas As Object, k As Variant
    Set capas = g_CollMgr.LoadCollection("capa")
    For Each k In capas.Keys
        If capas(k).EstEnRetard() Then
            notifs.Add CreerNotif(NotifCritique, "CAPA", _
                "CAPA en retard : " & capas(k).Numero, _
                "Échéance : " & FormatIsoDate(capas(k).DateEcheance) & _
                " — Responsable : " & capas(k).Responsable, _
                "OuvrirCapa", capas(k).Id)
        End If
    Next k
End Sub

Private Sub CollecterNcMajeuresOuvertes(ByRef notifs As Collection)
    Dim ncs As Object, k As Variant
    Set ncs = g_CollMgr.LoadCollection("nc")
    For Each k In ncs.Keys
        If ncs(k).Statut <> "Clôturée" And _
           (ncs(k).Classification = "Majeur" Or ncs(k).Classification = "Critique") Then
            notifs.Add CreerNotif(NotifCritique, "NC", _
                "NC " & ncs(k).Classification & " ouverte : " & ncs(k).Numero, _
                Left$(ncs(k).Description, 80), _
                "OuvrirNc", ncs(k).Id)
        End If
    Next k
End Sub

Private Sub CollecterActionsAmdecH(ByRef notifs As Collection)
    Dim amdecs As Object, k As Variant, l As CLigneAmdec
    Set amdecs = g_CollMgr.LoadCollection("amdec")
    For Each k In amdecs.Keys
        For Each l In amdecs(k).Lignes
            If l.ActionPriority() = "H" And Len(Trim$(l.ActionsRealisees)) = 0 Then
                notifs.Add CreerNotif(NotifAvertissement, "AMDEC", _
                    "Action H non traitée — " & amdecs(k).Reference, _
                    l.ModeDefaillance & " (RPN " & l.RPN() & ", S=" & l.S & ")", _
                    "OuvrirAmdec", amdecs(k).Id)
                Exit For
            End If
        Next l
    Next k
End Sub

Private Sub CollecterPpapASoumettre(ByRef notifs As Collection)
    Dim ppaps As Object, k As Variant
    Set ppaps = g_CollMgr.LoadCollection("ppap")
    For Each k In ppaps.Keys
        If ppaps(k).Statut = "En préparation" Then
            notifs.Add CreerNotif(NotifAvertissement, "PPAP", _
                "PPAP à soumettre : " & ppaps(k).Reference, _
                "Pièce : " & ppaps(k).Piece & " — " & ppaps(k).NbElementsFournis & "/18 éléments", _
                "OuvrirPpap", ppaps(k).Id)
        End If
    Next k
End Sub

Private Sub CollecterApqpEnDerive(ByRef notifs As Collection)
    Dim apqps As Object, k As Variant, jours As Long
    Set apqps = g_CollMgr.LoadCollection("apqp")
    For Each k In apqps.Keys
        If apqps(k).Statut = "En cours" And apqps(k).DateCibleSOP > 0 Then
            jours = DateDiff("d", Date, apqps(k).DateCibleSOP)
            If jours < 30 And apqps(k).Avancement < 80 Then
                notifs.Add CreerNotif(NotifCritique, "APQP", _
                    "APQP en dérive : " & apqps(k).Reference, _
                    "SOP dans " & jours & " j — avancement " & apqps(k).Avancement & "%", _
                    "OuvrirApqp", apqps(k).Id)
            End If
        End If
    Next k
End Sub

Private Sub CollecterMsaARenouveler(ByRef notifs As Collection)
    Dim msas As Object, k As Variant, jours As Long
    Set msas = g_CollMgr.LoadCollection("msa")
    For Each k In msas.Keys
        jours = DateDiff("d", msas(k).DateCreation, Date)
        If jours > 330 Then
            notifs.Add CreerNotif(NotifInfo, "MSA", _
                "MSA à renouveler : " & msas(k).Reference, _
                "Créé il y a " & jours & " jours — " & msas(k).Caracteristique, _
                "OuvrirMsa", msas(k).Id)
        End If
    Next k
End Sub

Public Function CompterParSeverite(ByVal notifs As Collection, ByVal sev As NotifSeverite) As Long
    Dim i As Long
    For i = 1 To notifs.Count
        If notifs(i)("Severite") = sev Then CompterParSeverite = CompterParSeverite + 1
    Next i
End Function

Public Sub OuvrirObjetDepuisNotif(ByVal action As String, ByVal objetId As String)
    Select Case action
        Case "OuvrirAudit": FAudit.Show vbModeless
        Case "OuvrirCapa": FNonConformite.Show vbModeless
        Case "OuvrirNc": FNonConformite.Show vbModeless
        Case "OuvrirAmdec": FAmdec.Show vbModeless
        Case "OuvrirPpap": FPpap.Show vbModeless
        Case "OuvrirApqp": FApqa.Show vbModeless
        Case "OuvrirMsa": FMsa.Show vbModeless
    End Select
End Sub