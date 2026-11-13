Attribute VB_Name = "M_Migration"
Option Explicit

' ============================================================
' M_Migration — Migration XML → SQL
' ============================================================

Public Function MigrerCollection(ByVal collection As String) As Long
    On Error GoTo ErrHandler
    If Not DbOpen() Then Exit Function

    Dim obj As Object, k As Variant, n As Long
    Set obj = g_CollMgr.LoadCollection(collection)
    If obj Is Nothing Then MigrerCollection = 0: Exit Function

    DbBegin
    For Each k In obj.Keys
        InsererObjet obj(k), collection
        n = n + 1
    Next k
    DbCommit

    Log LogInfo, "M_Migration", collection & " : " & n & " enregistrements migrés"
    MigrerCollection = n
    Exit Function
ErrHandler:
    DbRollback
    LogEx "M_Migration", "MigrerCollection", Err
    MigrerCollection = -1
End Function

Private Sub InsererObjet(ByVal obj As Object, ByVal collection As String)
    ' Construire dynamiquement l'INSERT à partir des propriétés
    Dim props As Variant, i As Long, cols As String, vals As String
    Dim params() As Variant
    props = GetDomainProperties(collection)
    If IsEmpty(props) Then Exit Sub

    ReDim params(LBound(props) To UBound(props))
    For i = LBound(props) To UBound(props)
        Dim v As Variant
        v = GetPropertyValue(obj, CStr(props(i)))
        If IsDate(v) Then
            v = Format$(v, "yyyy-mm-dd hh:nn:ss")
        ElseIf IsEmpty(v) Then
            v = Null
        End If
        cols = cols & IIf(Len(cols) > 0, ", ", "") & CStr(props(i))
        vals = vals & IIf(Len(vals) > 0, ", ", "") & "?"
        params(i) = v
    Next i
    DbExecute "INSERT INTO " & collection & " (" & cols & ") VALUES (" & vals & ")", params
End Sub

' Récupère la valeur d'une propriété par son nom (reflection VBA)
Private Function GetPropertyValue(ByVal obj As Object, ByVal propName As String) As Variant
    On Error Resume Next
    GetPropertyValue = CallByName(obj, propName, VbGet)
    If Err.Number <> 0 Then GetPropertyValue = Empty
    Err.Clear
End Function

Public Sub MigrerTout()
    Dim tables As Variant, i As Long, total As Long, nb As Long
    tables = Array("clients", "audits", "constats", "nc", "capa", "huitd", _
        "auditprogramme", "document", "exigence", _
        "amdec", "ligneamdec", "plancontrole", "ligneplan", "caracteristique", _
        "moyenmesure", "etalonnage", "etudemsa", "lignemsa", _
        "etudecapabilite", "cartecontrole", "lignecarte", "audittrail")
    For i = LBound(tables) To UBound(tables)
        nb = MigrerCollection(CStr(tables(i)))
        If nb >= 0 Then total = total + nb
    Next i
    ShowInfo "Migration", total & " enregistrements migrés vers " & g_Config.DbPath
End Sub