Public Function Load(ByVal collection As String) As Object
    If g_Config.DataMode = "SQL" Then
        Set Load = LoadSql(collection)
    Else
        Set Load = LoadXml(collection)
    End If
End Function

Private Function LoadSql(ByVal collection As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = DbQuery("SELECT * FROM " & collection)
    Do While Not rs.EOF
        Dim obj As Object
        Set obj = g_CollMgr.CreateEmpty(collection)
        RemplirDepuisRecordset obj, rs, collection
        result.Add obj.Id, obj
        rs.MoveNext
    Loop
    rs.Close
    Set LoadSql = result
End Function

Private Sub RemplirDepuisRecordset(ByVal obj As Object, ByVal rs As Object, ByVal collection As String)
    Dim props As Variant, i As Long
    props = GetDomainProperties(collection)
    For i = LBound(props) To UBound(props)
        On Error Resume Next
        Dim v As Variant
        v = rs(CStr(props(i)))
        If Not IsNull(v) Then CallByName obj, CStr(props(i)), VbLet, v
        Err.Clear
    Next i
End Sub

Public Sub Save(ByVal obj As Object, ByVal collection As String)
    If g_Config.DataMode = "SQL" Then
        SaveSql obj, collection
    Else
        SaveXml obj, collection
    End If
End Sub

Private Sub SaveSql(ByVal obj As Object, ByVal collection As String)
    Dim props As Variant, i As Long, setClause As String
    Dim params() As Variant
    props = GetDomainProperties(collection)
    ReDim params(LBound(props) To UBound(props) + 1)

    ' Détection : INSERT ou UPDATE
    Dim rs As Object
    Set rs = DbQuery("SELECT 1 FROM " & collection & " WHERE Id = ?", Array(obj.Id))
    Dim exists As Boolean: exists = Not rs.EOF
    rs.Close

    If exists Then
        For i = LBound(props) To UBound(props)
            Dim v As Variant: v = GetPropertyValue(obj, CStr(props(i)))
            If IsDate(v) Then v = Format$(v, "yyyy-mm-dd hh:nn:ss")
            setClause = setClause & IIf(Len(setClause) > 0, ", ", "") & props(i) & " = ?"
            params(i) = v
        Next i
        DbExecute "UPDATE " & collection & " SET " & setClause & " WHERE Id = ?", params
    Else
        Dim cols As String, vals As String
        For i = LBound(props) To UBound(props)
            Dim v2 As Variant: v2 = GetPropertyValue(obj, CStr(props(i)))
            If IsDate(v2) Then v2 = Format$(v2, "yyyy-mm-dd hh:nn:ss")
            cols = cols & IIf(Len(cols) > 0, ", ", "") & props(i)
            vals = vals & IIf(Len(vals) > 0, ", ", "") & "?"
            params(i) = v2
        Next i
        DbExecute "INSERT INTO " & collection & " (" & cols & ") VALUES (" & vals & ")", params
    End If

    Journaliser IIf(exists, "Modification", "Création"), collection, obj.Id, GetLibelle(obj, collection)
End Sub