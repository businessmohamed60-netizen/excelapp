Attribute VB_Name = "M_Stats"
Option Explicit

' ============================================================
' M_Stats — Fonctions statistiques (MSA, capabilité, SPC)
' ============================================================

' ----- Statistiques descriptives -----

Public Function Moyenne(arr As Variant) As Double
    Dim i As Long, s As Double
    For i = LBound(arr) To UBound(arr): s = s + arr(i): Next i
    If UBound(arr) >= LBound(arr) Then Moyenne = s / (UBound(arr) - LBound(arr) + 1)
End Function

Public Function EcartTypeGlobal(arr As Variant) As Double
    Dim i As Long, m As Double, s As Double, n As Long
    m = Moyenne(arr)
    n = UBound(arr) - LBound(arr) + 1
    If n < 2 Then Exit Function
    For i = LBound(arr) To UBound(arr): s = s + (arr(i) - m) ^ 2: Next i
    EcartTypeGlobal = Sqr(s / (n - 1))
End Function

Public Function Etendue(arr As Variant) As Double
    Dim i As Long, vMin As Double, vMax As Double
    vMin = arr(LBound(arr)): vMax = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If arr(i) < vMin Then vMin = arr(i)
        If arr(i) > vMax Then vMax = arr(i)
    Next i
    Etendue = vMax - vMin
End Function

' ----- Constantes d2, d3, A2, D3, D4 pour cartes X̄-R (n = 2..10) -----

Public Function ConstanteD2(ByVal n As Long) As Double
    Select Case n
        Case 2: ConstanteD2 = 1.128
        Case 3: ConstanteD2 = 1.693
        Case 4: ConstanteD2 = 2.059
        Case 5: ConstanteD2 = 2.326
        Case 6: ConstanteD2 = 2.534
        Case 7: ConstanteD2 = 2.704
        Case 8: ConstanteD2 = 2.847
        Case 9: ConstanteD2 = 2.97
        Case 10: ConstanteD2 = 3.078
        Case Else: ConstanteD2 = 0
    End Select
End Function

Public Function ConstanteA2(ByVal n As Long) As Double
    Select Case n
        Case 2: ConstanteA2 = 1.88
        Case 3: ConstanteA2 = 1.023
        Case 4: ConstanteA2 = 0.729
        Case 5: ConstanteA2 = 0.577
        Case 6: ConstanteA2 = 0.483
        Case 7: ConstanteA2 = 0.419
        Case 8: ConstanteA2 = 0.373
        Case 9: ConstanteA2 = 0.337
        Case 10: ConstanteA2 = 0.308
        Case Else: ConstanteA2 = 0
    End Select
End Function

Public Function ConstanteD3(ByVal n As Long) As Double
    Select Case n
        Case 2: ConstanteD3 = 0
        Case 3: ConstanteD3 = 0
        Case 4: ConstanteD3 = 0
        Case 5: ConstanteD3 = 0
        Case 6: ConstanteD3 = 0
        Case 7: ConstanteD3 = 0.076
        Case 8: ConstanteD3 = 0.136
        Case 9: ConstanteD3 = 0.184
        Case 10: ConstanteD3 = 0.223
        Case Else: ConstanteD3 = 0
    End Select
End Function

Public Function ConstanteD4(ByVal n As Long) As Double
    Select Case n
        Case 2: ConstanteD4 = 3.267
        Case 3: ConstanteD4 = 2.574
        Case 4: ConstanteD4 = 2.282
        Case 5: ConstanteD4 = 2.114
        Case 6: ConstanteD4 = 2.004
        Case 7: ConstanteD4 = 1.924
        Case 8: ConstanteD4 = 1.864
        Case 9: ConstanteD4 = 1.816
        Case 10: ConstanteD4 = 1.777
        Case Else: ConstanteD4 = 0
    End Select
End Function

' ----- MSA / R&R (méthode ANOVA simplifiée, moyenne des étendues) -----

Public Sub CalculerGrR(ByRef etude As CEtudeMsa)
    ' Charge les mesures
    Dim lignes As Object, k As Variant
    Dim nOp As Long, nPiece As Long, nEssai As Long
    Dim mesures() As Double
    Dim idx As Long, i As Long, o As Long, p As Long, e As Long

    nOp = etude.NbOperateurs: nPiece = etude.NbPieces: nEssai = etude.NbEssais
    If nOp < 2 Or nPiece < 2 Or nEssai < 2 Then Exit Sub

    Set lignes = g_CollMgr.LoadCollection("lignemsa")
    ReDim mesures(1 To nOp, 1 To nPiece, 1 To nEssai)
    Dim nbMesures As Long
    For Each k In lignes.Keys
        If lignes(k).EtudeId = etude.Id Then
            o = lignes(k).Operateur: p = lignes(k).Piece: e = lignes(k).Essai
            If o >= 1 And o <= nOp And p >= 1 And p <= nPiece And e >= 1 And e <= nEssai Then
                mesures(o, p, e) = lignes(k).Valeur
                nbMesures = nbMesures + 1
            End If
        End If
    Next k
    If nbMesures < nOp * nPiece * nEssai Then Exit Sub

    ' Moyenne par (opérateur, pièce)
    Dim moyOpPiece() As Double
    ReDim moyOpPiece(1 To nOp, 1 To nPiece)
    For o = 1 To nOp: For p = 1 To nPiece
        Dim s As Double
        For e = 1 To nEssai: s = s + mesures(o, p, e): Next e
        moyOpPiece(o, p) = s / nEssai
    Next p: Next o

    ' Étendue par (opérateur, pièce) et étendue moyenne
    Dim rBar As Double, sumR As Double
    Dim etendueOpPiece() As Double
    ReDim etendueOpPiece(1 To nOp, 1 To nPiece)
    For o = 1 To nOp: For p = 1 To nPiece
        Dim vMin As Double, vMax As Double
        vMin = mesures(o, p, 1): vMax = mesures(o, p, 1)
        For e = 2 To nEssai
            If mesures(o, p, e) < vMin Then vMin = mesures(o, p, e)
            If mesures(o, p, e) > vMax Then vMax = mesures(o, p, e)
        Next e
        etendueOpPiece(o, p) = vMax - vMin
        sumR = sumR + etendueOpPiece(o, p)
    Next p: Next o
    rBar = sumR / (nOp * nPiece)

    ' EV = R̄ / d2(n)
    Dim ev As Double
    ev = rBar / ConstanteD2(nEssai)
    If ConstanteD2(nEssai) = 0 Then Exit Sub

    ' Moyenne par opérateur
    Dim moyOp() As Double
    ReDim moyOp(1 To nOp)
    Dim sumAll As Double, nAll As Long
    For o = 1 To nOp
        Dim s2 As Double
        For p = 1 To nPiece: s2 = s2 + moyOpPiece(o, p): Next p
        moyOp(o) = s2 / nPiece
        sumAll = sumAll + s2: nAll = nAll + nPiece
    Next o

    ' Étendue des moyennes opérateurs
    Dim rOp As Double, vMinOp As Double, vMaxOp As Double
    vMinOp = moyOp(1): vMaxOp = moyOp(1)
    For o = 2 To nOp
        If moyOp(o) < vMinOp Then vMinOp = moyOp(o)
        If moyOp(o) > vMaxOp Then vMaxOp = moyOp(o)
    Next o
    rOp = vMaxOp - vMinOp

    ' AV = sqrt((R̄_op / d2(nOp))² - EV²/(nPiece*nEssai))
    Dim av2 As Double, av As Double
    If ConstanteD2(nOp) > 0 Then
        av2 = (rOp / ConstanteD2(nOp)) ^ 2 - (ev ^ 2) / (nPiece * nEssai)
        If av2 > 0 Then av = Sqr(av2)
    End If

    ' GRR = sqrt(EV² + AV²)
    Dim grr As Double
    grr = Sqr(ev ^ 2 + av ^ 2)

    ' Écart-type total (TV) — toutes les mesures
    Dim sumAllMes As Double, nMes As Long
    nMes = nOp * nPiece * nEssai
    Dim tv As Double, sTV As Double
    For o = 1 To nOp: For p = 1 To nPiece: For e = 1 To nEssai
        sumAllMes = sumAllMes + mesures(o, p, e)
    Next e: Next p: Next o
    Dim mAll As Double: mAll = sumAllMes / nMes
    For o = 1 To nOp: For p = 1 To nPiece: For e = 1 To nEssai
        sTV = sTV + (mesures(o, p, e) - mAll) ^ 2
    Next e: Next p: Next o
    tv = Sqr(sTV / (nMes - 1))

    ' Remplissage
    etude.Moyenne = mAll
    etude.EtendueMoyenne = rBar
    etude.EcartTypeRepetabilite = ev
    etude.EcartTypeReproductibilite = av
    etude.GrR = grr
    If tv > 0 Then etude.PourcentGrR = grr / tv * 100
    If etude.Tolerance > 0 Then etude.PourcentGrRTolerance = grr / etude.Tolerance * 100
    If etude.PourcentGrRTolerance > 0 Then
        etude.Ndc = Int(1.41 * (etude.PourcentGrRTolerance / 100) ^ -1)
    End If
End Sub

' ----- Capabilité -----

Public Sub CalculerCapabilite(ByRef etude As CEtudeCapabilite, valeurs As Variant)
    Dim m As Double, sLong As Double, n As Long
    m = Moyenne(valeurs)
    sLong = EcartTypeGlobal(valeurs)
    n = UBound(valeurs) - LBound(valeurs) + 1

    ' Écart-type court terme approximé par la moyenne des étendues de sous-groupes
    ' (simplification : on utilise l'écart-type global si pas de sous-groupes)
    Dim sCourt As Double
    sCourt = sLong

    etude.Moyenne = m
    etude.NbValeurs = n
    etude.EcartTypeLong = sLong
    etude.EcartTypeCourt = sCourt

    If sLong > 0 And etude.USL > etude.LSL Then
        etude.Cp = (etude.USL - etude.LSL) / (6 * sCourt)
        etude.Cpk = Application.Min((etude.USL - m) / (3 * sCourt), (m - etude.LSL) / (3 * sCourt))
        etude.Pp = (etude.USL - etude.LSL) / (6 * sLong)
        etude.Ppk = Application.Min((etude.USL - m) / (3 * sLong), (m - etude.LSL) / (3 * sLong))
    End If
End Sub

' ----- Carte de contrôle X̄-R -----

Public Sub CalculerCarteXbarR(ByRef carte As CCarteControle, sousGroupes As Variant)
    ' sousGroupes : tableau 2D (1..k, 1..n)
    Dim k As Long, n As Long, i As Long, j As Long
    k = UBound(sousGroupes, 1)
    n = UBound(sousGroupes, 2)
    carte.NbSousGroupes = k
    carte.SousGroupeTaille = n

    Dim moyennes() As Double, etendues() As Double
    ReDim moyennes(1 To k): ReDim etendues(1 To k)
    Dim sumMoy As Double, sumR As Double
    For i = 1 To k
        Dim s As Double, vMin As Double, vMax As Double
        vMin = sousGroupes(i, 1): vMax = sousGroupes(i, 1)
        For j = 1 To n
            s = s + sousGroupes(i, j)
            If sousGroupes(i, j) < vMin Then vMin = sousGroupes(i, j)
            If sousGroupes(i, j) > vMax Then vMax = sousGroupes(i, j)
        Next j
        moyennes(i) = s / n
        etendues(i) = vMax - vMin
        sumMoy = sumMoy + moyennes(i)
        sumR = sumR + etendues(i)
    Next i
    Dim xBarBar As Double, rBar As Double
    xBarBar = sumMoy / k
    rBar = sumR / k

    carte.LigneCentre = xBarBar
    carte.LigneCentreR = rBar
    carte.LimiteSup = xBarBar + ConstanteA2(n) * rBar
    carte.LimiteInf = xBarBar - ConstanteA2(n) * rBar
    carte.LimiteSupR = ConstanteD4(n) * rBar
    carte.LimiteInfR = ConstanteD3(n) * rBar
End Sub

' ----- Règles de Nelson (détection de dérives) -----

Public Function AppliquerReglesNelson(valeurs As Variant, _
                                      ByVal ligneCentre As Double, _
                                      ByVal ucl As Double, _
                                      ByVal lcl As Double) As Collection
    ' Retourne une collection de chaînes : "Point i : règle N"
    Dim result As New Collection
    Dim i As Long, n As Long, sigma As Double
    n = UBound(valeurs) - LBound(valeurs) + 1
    If n < 2 Then Set AppliquerReglesNelson = result: Exit Function
    sigma = (ucl - ligneCentre) / 3

    For i = LBound(valeurs) To UBound(valeurs)
        ' Règle 1 : point hors limites
        If valeurs(i) > ucl Or valeurs(i) < lcl Then
            result.Add "Point " & i & " : hors limites (R1)"
        End If
    Next i

    ' Règle 2 : 9 points consécutifs du même côté de la ligne centre
    Dim count As Long, side As Long
    For i = LBound(valeurs) To UBound(valeurs)
        Dim s As Long
        If valeurs(i) > ligneCentre Then s = 1 ElseIf valeurs(i) < ligneCentre Then s = -1 Else s = 0
        If s = side And s <> 0 Then
            count = count + 1
            If count >= 9 Then result.Add "Point " & i & " : 9 points du même côté (R2)"
        Else
            side = s: count = 1
        End If
    Next i

    ' Règle 3 : 6 points consécutifs en tendance croissante ou décroissante
    Dim trend As Long, trendCount As Long
    For i = LBound(valeurs) + 1 To UBound(valeurs)
        Dim t As Long
        If valeurs(i) > valeurs(i - 1) Then t = 1 ElseIf valeurs(i) < valeurs(i - 1) Then t = -1 Else t = 0
        If t = trend And t <> 0 Then
            trendCount = trendCount + 1
            If trendCount >= 6 Then result.Add "Point " & i & " : 6 points en tendance (R3)"
        Else
            trend = t: trendCount = 1
        End If
    Next i

    Set AppliquerReglesNelson = result
End Function