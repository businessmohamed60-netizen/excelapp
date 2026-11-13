Private Sub RefreshActivites(ByVal audits As Object, ByVal constats As Object)
    lstActivites.Clear
    Dim items As Object
    Set items = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In audits.Keys
        items.Add Format$(audits(k).DateDebut, "yyyymmddhhnnss") & "A" & k, _
            Format$(audits(k).DateDebut, "yyyy-mm-dd") & " | Audit | " & _
            audits(k).Processus & " (" & audits(k).Statut & ")"
    Next k
    For Each k In constats.Keys
        items.Add Format$(constats(k).DateCreation, "yyyymmddhhnnss") & "C" & k, _
            Format$(constats(k).DateCreation, "yyyy-mm-dd") & " | Constat | " & _
            constats(k).Classification & " — " & Left$(constats(k).Description, 50)
    Next k

    Dim keys As Variant, i As Long, j As Long, tmp As String
    keys = items.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If items(keys(j)) > items(keys(i)) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    Dim n As Long
    For i = LBound(keys) To UBound(keys)
        lstActivites.AddItem items(keys(i))
        n = n + 1
        If n >= 5 Then Exit For
    Next i
    If lstActivites.ListCount = 0 Then lstActivites.AddItem "(aucune activité enregistrée)"
End Sub