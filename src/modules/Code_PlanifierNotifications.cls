Option Explicit

Public Sub PlanifierNotifications()
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:30:00"), "RafraichirNotifications"
End Sub

Public Sub RafraichirNotifications()
    On Error Resume Next
    Dim notifs As Collection
    Set notifs = CollecterNotifications()
    If notifs.Count > 0 Then
        Dim nCrit As Long
        nCrit = CompterParSeverite(notifs, NotifCritique)
        If nCrit > 0 Then Application.StatusBar = "IATF Assistant — " & nCrit & " alerte(s) critique(s)"
    End If
    PlanifierNotifications
End Sub

Public Sub AnnulerPlanification()
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:30:00"), "RafraichirNotifications", , False
End Sub