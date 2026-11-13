Attribute VB_Name = "M_Amdec"
Option Explicit

' ============================================================
' M_Amdec — Calculs AMDEC, statistiques, tri par RPN
' ============================================================

' Charge toutes les lignes d'une AMDEC, triées par RPN décroissant
Public Function LignesAmdecTriees(ByVal amdecId As String) As Collection
    Dim allLignes As Object, k As Variant, filtered As New Collection
    Dim lignes As Object, arr() As CLigneAMDEC, i As Long, j As Long, tmp As CLigneAMDEC

    Set allLignes = g_CollMgr.LoadCollection("ligneamdec")
    For Each k In allLignes.Keys
        If allLignes(k).AmdecId = amdecId Then filtered.Add allLignes(k)
    Next k

    ' Tri à bulles par RPN décroissant
    If filtered.Count > 1 Then
        ReDim arr(1 To filtered.Count)
        For i = 1 To filtered.Count: Set arr(i) = filtered(i): Next i
        For i = 1 To UBound(arr) - 1
            For j = i + 1 To UBound(arr)
                If arr(j).Rpn > arr(i).Rpn Then
                    Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
                End If
            Next j
        Next i
        Set LignesAmdecTriees = New Collection
        For i = 1 To UBound(arr): LignesAmdecTriees.Add arr(i): Next i
    Else
        Set LignesAmdecTriees = filtered
    End If
End Function

' Statistiques d'une AMDEC
Public Type TStatAmdec
    NbLignes As Long
    NbCritiques As Long
    NbElevees As Long
    NbMoyennes As Long
    NbFaibles As Long
    RpnMax As Long
    RpnMoyen As Double
    NbActionsOuvertes As Long
    NbActionsRetard As Long
    GainRpnMoyen As Double
End Type

Public Function StatistiquesAmdec(ByVal amdecId As String) As TStatAmdec
    Dim s As TStatAmdec
    Dim allLignes As Object, k As Variant, ligne As CLigneAMDEC
    Dim sumRpn As Double, sumGain As Double, nGain As Long

    Set allLignes = g_CollMgr.LoadCollection("ligneamdec")
    For Each k In allLignes.Keys
        If allLignes(k).AmdecId = amdecId Then
            Set ligne = allLignes(k)
            s.NbLignes = s.NbLignes + 1
            Select Case ligne.Criticite
                Case "Critique": s.NbCritiques = s.NbCritiques + 1
                Case "Élevée":   s.NbElevees = s.NbElevees + 1
                Case "Moyenne":  s.NbMoyennes = s.NbMoyennes + 1
                Case "Faible":   s.NbFaibles = s.NbFaibles + 1
            End Select
            If ligne.Rpn > s.RpnMax Then s.RpnMax = ligne.Rpn
            sumRpn = sumRpn + ligne.Rpn
            If ligne.StatutAction <> "Clôturée" Then s.NbActionsOuvertes = s.NbActionsOuvertes + 1
            If ligne.EstEnRetard() Then s.NbActionsRetard = s.NbActionsRetard + 1
            If ligne.RpnApres > 0 And ligne.Rpn > 0 Then
                sumGain = sumGain + (ligne.Rpn - ligne.RpnApres)
                nGain = nGain + 1
            End If
        End If
    Next k

    If s.NbLignes > 0 Then s.RpnMoyen = sumRpn / s.NbLignes
    If nGain > 0 Then s.GainRpnMoyen = sumGain / nGain
    StatistiquesAmdec = s
End Function

' Évolution du RPN avant/après actions, pour une AMDEC
Public Function TableauAvantApres(ByVal amdecId As String) As Variant
    Dim lignes As Collection, i As Long, arr() As Variant, ligne As CLigneAMDEC
    Set lignes = LignesAmdecTriees(amdecId)
    ReDim arr(1 To lignes.Count, 1 To 4)
    For i = 1 To lignes.Count
        Set ligne = lignes(i)
        arr(i, 1) = ligne.NumeroLigne
        arr(i, 2) = ligne.ModeDefaillance
        arr(i, 3) = ligne.Rpn
        arr(i, 4) = ligne.RpnApres
    Next i
    TableauAvantApres = arr
End Function