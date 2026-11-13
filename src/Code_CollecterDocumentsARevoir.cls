' À ajouter dans CollecterNotifications (M_Notifications.bas), juste avant Set CollecterNotifications = notifs
    CollecterDocumentsARevoir notifs

' Nouvelle méthode privée
Private Sub CollecterDocumentsARevoir(ByRef notifs As Collection)
    On Error Resume Next
    ContribuerNotifications notifs
End Sub