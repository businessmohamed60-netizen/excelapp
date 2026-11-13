Private Sub ExporterCollectionJson(ByVal collection As String, ByVal apiDir As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open apiDir & collection & ".json" For Output As #f
    Print #f, "["
    Dim first As Boolean: first = True

    If g_Config.DataMode = "SQL" Then
        Dim rs As Object
        Set rs = DbQuery("SELECT * FROM " & collection)
        Do While Not rs.EOF
            If Not first Then Print #f, ","
            Print #f, RecordsetVersJson(rs, collection);
            first = False
            rs.MoveNext
        Loop
        rs.Close
    Else
        ' Fallback XML (code existant)
        Dim coll As Object, k As Variant
        Set coll = g_CollMgr.LoadCollection(collection)
        For Each k In coll.Keys
            If Not first Then Print #f, ","
            Print #f, ObjetVersJson(coll(k), collection);
            first = False
        Next k
    End If

    Print #f, vbCrLf & "]"
    Close #f
End Sub

Private Function RecordsetVersJson(ByVal rs As Object, ByVal collection As String) As String
    Dim props As Variant, i As Long, s As String
    props = GetDomainProperties(collection)
    s = "{"
    For i = LBound(props) To UBound(props)
        If i > LBound(props) Then s = s & ", "
        Dim colName As String
        colName = IIf(CStr(props(i)) = "Collection", "CollectionName", CStr(props(i)))
        On Error Resume Next
        Dim v As Variant: v = rs(colName)
        Err.Clear
        s = s & """" & props(i) & """: " & ValeurVersJson(v)
    Next i
    s = s & "}"
    RecordsetVersJson = s
End Function