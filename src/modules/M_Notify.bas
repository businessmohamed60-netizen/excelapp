Attribute VB_Name = "M_Notify"
Option Explicit

' ============================================================
' M_Notify — Notifications Outlook pour échéances
' ============================================================

Public Type TEcheance
    Categorie As String
    Reference As String
    Description As String
    DateEcheance As Date
    JoursRestants As Long
    Responsable As String
    Niveau As String           ' Critique / Urgent / Attention / Info
End Type

' Collecte toutes les échéances à venir dans les N jours
Public Function CollecterEcheances(Optional ByVal horizonJours As Long = 30) As Collection
    Dim result As New Collection
    Dim k As Variant, d As Date, j As Long
    Dim coll As Object

    ' CAPA
    Set coll = g_CollMgr.LoadCollection("capa")
    For Each k In coll.Keys
        If coll(k).Statut <> "Clôturée" And IsDate(coll(k).DateEcheance) Then
            d = coll(k).DateEcheance
            j = DateDiff("d", Date, d)
            If j <= horizonJours Then
                result.Add ConstruireEcheance("CAPA", coll(k).Numero, _
                    Left$(coll(k).Description, 60), d, j, coll(k).Responsable)
            End If
        End If
    Next k

    ' Étalonnages
    Set coll = g_CollMgr.LoadCollection("moyenmesure")
    For Each k In coll.Keys
        If coll(k).Statut = "En service" And IsDate(coll(k).ProchainEtalonnage) Then
            d = coll(k).ProchainEtalonnage
            j = DateDiff("d", Date, d)
            If j <= horizonJours Then
                result.Add ConstruireEcheance("Étalonnage", coll(k).Code, _
                    coll(k).Designation, d, j, coll(k).Responsable)
            End If
        End If
    Next k

    ' Revues documentaires
    Set coll = g_CollMgr.LoadCollection("document")
    For Each k In coll.Keys
        If coll(k).Statut = "Approuvé" And IsDate(coll(k).DateRevuePrevue) Then
            d = coll(k).DateRevuePrevue
            j = DateDiff("d", Date, d)
            If j <= horizonJours Then
                result.Add ConstruireEcheance("Revue documentaire", coll(k).Code, _
                    coll(k).Titre, d, j, coll(k).Approbateur)
            End If
        End If
    Next k

    ' Actions AMDEC
    Set coll = g_CollMgr.LoadCollection("ligneamdec")
    Dim amdecColl As Object, amdecRef As String
    Set amdecColl = g_CollMgr.LoadCollection("amdec")
    For Each k In coll.Keys
        If coll(k).StatutAction <> "Clôturée" And IsDate(coll(k).DateEcheanceAction) Then
            d = coll(k).DateEcheanceAction
            j = DateDiff("d", Date, d)
            If j <= horizonJours Then
                amdecRef = ""
                If amdecColl.Exists(coll(k).AmdecId) Then amdecRef = amdecColl(coll(k).AmdecId).Reference
                result.Add ConstruireEcheance("Action AMDEC", amdecRef, _
                    coll(k).ModeDefaillance, d, j, coll(k).ResponsableAction)
            End If
        End If
    Next k

    ' Audits programmés
    Set coll = g_CollMgr.LoadCollection("auditprogramme")
    For Each k In coll.Keys
        If coll(k).Statut = "Planifié" And IsDate(coll(k).DatePrevue) Then
            d = coll(k).DatePrevue
            j = DateDiff("d", Date, d)
            If j <= horizonJours Then
                result.Add ConstruireEcheance("Audit programmé", CStr(coll(k).Annee), _
                    coll(k).Processus & " (" & coll(k).TypeAudit & ")", d, j, coll(k).Auditeur)
            End If
        End If
    Next k

    Set CollecterEcheances = result
End Function

Private Function ConstruireEcheance(ByVal cat As String, ByVal ref As String, _
                                    ByVal desc As String, ByVal d As Date, _
                                    ByVal j As Long, ByVal resp As String) As TEcheance
    Dim e As TEcheance
    e.Categorie = cat: e.Reference = ref: e.Description = desc
    e.DateEcheance = d: e.JoursRestants = j: e.Responsable = resp
    If j < 0 Then
        e.Niveau = "Critique"
    ElseIf j <= 7 Then
        e.Niveau = "Urgent"
    ElseIf j <= 30 Then
        e.Niveau = "Attention"
    Else
        e.Niveau = "Info"
    End If
    ConstruireEcheance = e
End Function

' Génère le corps HTML de l'email récapitulatif
Public Function GenererCorpsEmail(echeances As Collection) As String
    Dim html As String, i As Long, e As TEcheance
    html = "<html><body><h2>Rappel des échéances IATF</h2>"
    html = html & "<p>Bonjour,</p><p>Voici les échéances à venir dans les 30 prochains jours :</p>"
    html = html & "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-family:Arial;font-size:12px'>"
    html = html & "<tr style='background:#003366;color:white'><th>Niveau</th><th>Catégorie</th><th>Référence</th><th>Description</th><th>Échéance</th><th>Jours</th><th>Responsable</th></tr>"

    ' Tri par niveau puis par échéance
    Dim niveaux(3) As String, n As Long, k As Long
    niveaux(0) = "Critique": niveaux(1) = "Urgent": niveaux(2) = "Attention": niveaux(3) = "Info"
    For k = 0 To 3
        For i = 1 To echeances.Count
            e = echeances(i)
            If e.Niveau = niveaux(k) Then
                Dim couleur As String
                Select Case e.Niveau
                    Case "Critique": couleur = "#FFCCCC"
                    Case "Urgent": couleur = "#FFE5CC"
                    Case "Attention": couleur = "#FFF5CC"
                    Case Else: couleur = "#FFFFFF"
                End Select
                html = html & "<tr style='background:" & couleur & "'>"
                html = html & "<td><b>" & e.Niveau & "</b></td>"
                html = html & "<td>" & e.Categorie & "</td>"
                html = html & "<td>" & e.Reference & "</td>"
                html = html & "<td>" & e.Description & "</td>"
                html = html & "<td>" & Format$(e.DateEcheance, "yyyy-mm-dd") & "</td>"
                html = html & "<td>" & e.JoursRestants & "</td>"
                html = html & "<td>" & e.Responsable & "</td>"
                html = html & "</tr>"
            End If
        Next i
    Next k

    html = html & "</table><p>Cordialement,<br/>IATF Consultant Assistant</p></body></html>"
    GenererCorpsEmail = html
End Function

' Envoie l'email récapitulatif via Outlook
Public Function EnvoyerNotification(Optional ByVal destinataire As String = "", _
                                    Optional ByVal horizonJours As Long = 30) As Boolean
    On Error GoTo ErrHandler
    Dim echeances As Collection
    Set echeances = CollecterEcheances(horizonJours)
    If echeances.Count = 0 Then
        Log LogInfo, "M_Notify", "Aucune échéance à notifier."
        EnvoyerNotification = True
        Exit Function
    End If

    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)
    mail.Subject = "IATF — " & echeances.Count & " échéance(s) à venir"
    mail.HTMLBody = GenererCorpsEmail(echeances)
    If Len(destinataire) > 0 Then mail.To = destinataire
    mail.Display ' Ouvre l'email pour relecture (au lieu d'envoyer directement)
    EnvoyerNotification = True
    Log LogInfo, "M_Notify", echeances.Count & " échéance(s) notifiée(s)."
    Exit Function
ErrHandler:
    LogEx "M_Notify", "EnvoyerNotification", Err
    EnvoyerNotification = False
End Function

' Vérifie au démarrage si des échéances critiques existent
Public Sub VerifierEcheancesAuDemarrage()
    Dim echeances As Collection, i As Long, e As TEcheance, nbCritiques As Long
    Set echeances = CollecterEcheances(7)
    For i = 1 To echeances.Count
        e = echeances(i)
        If e.Niveau = "Critique" Then nbCritiques = nbCritiques + 1
    Next i
    If nbCritiques > 0 Then
        If PromptYesNo("Échéances critiques", _
            nbCritiques & " échéance(s) critique(s) détectée(s). Voulez-vous ouvrir le centre de notifications ?") Then
            FNotifications.Show vbModeless
        End If
    End If
End Sub