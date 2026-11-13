Option Explicit

Public Function Moyenne(ByRef arr() As Double) As Double
    Dim i As Long, s As Double
    For i = LBound(arr) To UBound(arr): s = s + arr(i): Next i
    If UBound(arr) >= LBound(arr) Then Moyenne = s / (UBound(arr) - LBound(arr) + 1)
End Function

Public Function EcartType(ByRef arr() As Double, Optional ByVal echantillon As Boolean = True) As Double
    Dim m As Double, s As Double, i As Long, n As Long
    n = UBound(arr) - LBound(arr) + 1
    If n < 2 Then Exit Function
    m = Moyenne(arr)
    For i = LBound(arr) To UBound(arr): s = s + (arr(i) - m) ^ 2: Next i
    If echantillon Then EcartType = Sqr(s / (n - 1)) Else EcartType = Sqr(s / n)
End Function

Public Function Etendue(ByRef arr() As Double) As Double
    Dim i As Long, mn As Double, mx As Double
    mn = arr(LBound(arr)): mx = mn
    For i = LBound(arr) To UBound(arr)
        If arr(i) < mn Then mn = arr(i)
        If arr(i) > mx Then mx = arr(i)
    Next i
    Etendue = mx - mn
End Function

Public Function Minimum(ByRef arr() As Double) As Double
    Dim i As Long, mn As Double
    mn = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr): If arr(i) < mn Then mn = arr(i): Next i
    Minimum = mn
End Function

Public Function Maximum(ByRef arr() As Double) As Double
    Dim i As Long, mx As Double
    mx = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr): If arr(i) > mx Then mx = arr(i): Next i
    Maximum = mx
End Function

Public Function C4(ByVal n As Long) As Double
    Select Case n
        Case 2: C4 = 0.7979: Case 3: C4 = 0.8862: Case 4: C4 = 0.9213
        Case 5: C4 = 0.94: Case 6: C4 = 0.9515: Case 7: C4 = 0.9594
        Case 8: C4 = 0.965: Case 9: C4 = 0.9693: Case 10: C4 = 0.9727
        Case Else: C4 = 1
    End Select
End Function

Public Function D3(ByVal n As Long) As Double
    Select Case n
        Case 2, 3, 4, 5, 6: D3 = 0
        Case 7: D3 = 0.076: Case 8: D3 = 0.136
        Case 9: D3 = 0.184: Case 10: D3 = 0.223
        Case Else: D3 = 0
    End Select
End Function

Public Function D4(ByVal n As Long) As Double
    Select Case n
        Case 2: D4 = 3.267: Case 3: D4 = 2.574: Case 4: D4 = 2.282
        Case 5: D4 = 2.114: Case 6: D4 = 2.004: Case 7: D4 = 1.924
        Case 8: D4 = 1.864: Case 9: D4 = 1.816: Case 10: D4 = 1.777
        Case Else: D4 = 3.267
    End Select
End Function

Public Function A2(ByVal n As Long) As Double
    Select Case n
        Case 2: A2 = 1.88: Case 3: A2 = 1.023: Case 4: A2 = 0.729
        Case 5: A2 = 0.577: Case 6: A2 = 0.483: Case 7: A2 = 0.419
        Case 8: A2 = 0.373: Case 9: A2 = 0.337: Case 10: A2 = 0.308
        Case Else: A2 = 1.88
    End Select
End Function

Public Function NormSDist(ByVal z As Double) As Double
    Dim t As Double, p As Double
    t = 1 / (1 + 0.2316419 * Abs(z))
    p = 1 - 0.3989423 * Exp(-z * z / 2) * _
        (t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274)))))
    If z >= 0 Then NormSDist = p Else NormSDist = 1 - p
End Function