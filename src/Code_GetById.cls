Public Function GetById(ByVal id As String, ByVal collName As String) As Object
    Dim coll As Object, k As Variant
    Set coll = LoadCollection(collName)
    If coll.Exists(id) Then Set GetById = coll(id)
End Function