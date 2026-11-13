Option Explicit

' ============================================================
' M_Heatmap — Matrice d'activité 12 semaines x 7 jours
' ============================================================

' Retourne un Dictionary : clé = "yyyy-mm-dd", valeur = nombre d'événements
Public Function CalculerActiviteParJour(ByVal semaines As Long) As Object
    Dim activite As Object
    Set activite = CreateObject("Scripting.Dictionary")
    Dim dateDebut As Date
    dateDebut = Date - (semaines * 7) + 1

    AjouterDatesCollection activite, "audits", "DateCreation", dateDebut
    AjouterDatesCollection activite, "nc", "DateCreation", dateDebut
    AjouterDatesCollection activite, "capa", "DateCreation", dateDebut
    AjouterDatesCollection activite, "constats", "DateCreation", dateDebut

    Set CalculerActiviteParJour = activite
End Function

Private Sub AjouterDatesCollection(ByVal activite As Object, ByVal collection As String, _
                                   ByVal champ As String, ByVal dateDebut As Date)
    Dim d As Object, k As Variant, dte As Date
    Set d = g_CollMgr.LoadCollection(collection)
    For Each k In d.Keys
        On Error Resume Next
        dte = CallByName(d(k), champ, VbGet)
        If Err.Number = 0 And IsDate(dte) And dte >= dateDebut Then
            Dim cle As String: cle = FormatIsoDate(dte)
            If activite.Exists(cle) Then
                activite(cle) = activite(cle) + 1
            Else
                activite.Add cle, 1
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next k
End Sub

' Retourne la couleur de la case selon l'intensité (0 = très clair, max = très foncé)
Public Function CouleurHeatmap(ByVal valeur As Long, ByVal maximum As Long) As Long
    If maximum = 0 Or valeur = 0 Then CouleurHeatmap = RGB(235, 237, 240): Exit Function
    Dim ratio As Double: ratio = valeur / maximum
    Dim r As Long, g As Long, b As Long
    ' Gradient du vert clair (220, 245, 220) vers le vert foncé (25, 135, 84)
    r = CLng(220 - (220 - 25) * ratio)
    g = CLng(245 - (245 - 135) * ratio)
    b = CLng(220 - (220 - 84) * ratio)
    CouleurHeatmap = RGB(r, g, b)
End Function

' Retourne le maximum d'activité sur la période
Public Function ActiviteMax(ByVal activite As Object) As Long
    Dim k As Variant, mx As Long
    For Each k In activite.Keys
        If activite(k) > mx Then mx = activite(k)
    Next k
    ActiviteMax = mx
End Function