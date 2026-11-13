Public Sub AssignNumero()
    Dim dom As Object, root As Object, yr As String
    Dim fname As String, item As Object, numeroNode As Object
    Dim maxN As Long, parts() As String, n As Long
    fname = DataPath() & "nc.xml"
    yr = Format$(Year(m_DateDetection), "0000")

    If Dir(fname) <> "" Then
        Set dom = CreateObject("MSXML2.DOMDocument.6.0")
        dom.async = False
        If dom.Load(fname) Then
            Set root = dom.documentElement
            For Each item In root.ChildNodes
                Set numeroNode = item.selectSingleNode("Numero")
                If Not numeroNode Is Nothing Then
                    parts = Split(numeroNode.Text, "-")
                    If UBound(parts) >= 2 Then
                        If parts(1) = yr Then
                            n = CLng(Val(parts(2)))
                            If n > maxN Then maxN = n
                        End If
                    End If
                End If
            Next item
        End If
    End If

    ' Boucle d'incrément : garantir l'unicité
    Dim candidate As String
    Do
        maxN = maxN + 1
        candidate = "NC-" & yr & "-" & Format$(maxN, "000")
    Loop While NumeroExiste(candidate, fname)
    m_Numero = candidate
End Sub

Private Function NumeroExiste(ByVal numero As String, ByVal fname As String) As Boolean
    If Dir(fname) = "" Then Exit Function
    Dim dom As Object, item As Object, node As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(fname) Then Exit Function
    For Each item In dom.documentElement.ChildNodes
        Set node = item.selectSingleNode("Numero")
        If Not node Is Nothing Then
            If node.Text = numero Then NumeroExiste = True: Exit Function
        End If
    Next item
End Function