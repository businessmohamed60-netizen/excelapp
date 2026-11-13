Option Explicit

Private m_cache As Object
Private m_index As Object
Private m_repo As CRepository

Private Sub Class_Initialize()
    Set m_cache = CreateObject("Scripting.Dictionary")
    Set m_index = CreateObject("Scripting.Dictionary")
    Set m_repo = New CRepository
End Sub

Public Property Get Repo() As CRepository
    Set Repo = m_repo
End Property

Public Function CreateEmpty(ByVal collection As String) As Object
    Select Case LCase$(collection)
        Case "clients": Set CreateEmpty = New CClient
        Case "audits": Set CreateEmpty = New CAudit
        Case "constats": Set CreateEmpty = New CConstat
        Case "nc": Set CreateEmpty = New CNonConformite
        Case "capa": Set CreateEmpty = New CCapa
        Case "huitd": Set CreateEmpty = New C8D
        Case "apqp": Set CreateEmpty = New CProjetAPQP
        Case "ppap": Set CreateEmpty = New CPpapDossier
        Case "amdec": Set CreateEmpty = New CAmdec
        Case "spc": Set CreateEmpty = New CSpcEtude
        Case "msa": Set CreateEmpty = New CMsaEtude
        Case Else: Set CreateEmpty = New CClient
    End Select
End Function

Public Function LoadCollection(ByVal collection As String) As Object
    If m_cache.Exists(collection) Then Set LoadCollection = m_cache(collection): Exit Function
    Dim dict As Object
    m_repo.LoadAll collection, Me, dict
    m_cache.Add collection, dict
    Set LoadCollection = dict
End Function

Public Sub InvalidateCache(ByVal collection As String)
    If m_cache.Exists(collection) Then m_cache.Remove collection
End Sub

Public Sub InvalidateAll()
    Set m_cache = CreateObject("Scripting.Dictionary")
    Set m_index = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Save(ByVal obj As Object, ByVal collection As String)
    m_repo.Save obj, collection
    Dim dict As Object
    Set dict = LoadCollection(collection)
    If dict.Exists(obj.Id) Then Set dict(obj.Id) = obj Else dict.Add obj.Id, obj
    If m_index.Exists(collection) Then m_index.Remove collection
End Sub

Public Function DeleteById(ByVal id As String, ByVal collection As String) As Boolean
    Dim ok As Boolean
    ok = m_repo.DeleteById(id, collection)
    If ok Then
        Dim dict As Object
        Set dict = LoadCollection(collection)
        If dict.Exists(id) Then dict.Remove id
        If m_index.Exists(collection) Then m_index.Remove collection
    End If
    DeleteById = ok
End Function

Public Function FindByKey(ByVal collection As String, ByVal cleMetier As String, _
                          ByVal valeur As String) As Object
    Dim dict As Object, k As Variant
    Set dict = LoadCollection(collection)
    If Not m_index.Exists(collection) Then
        Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
        m_index.Add collection, idx
    End If
    Dim idxColl As Object: Set idxColl = m_index(collection)
    Dim indexKey As String: indexKey = LCase$(cleMetier & "=" & valeur)
    If idxColl.Exists(indexKey) Then
        Dim id As String: id = idxColl(indexKey)
        If dict.Exists(id) Then Set FindByKey = dict(id): Exit Function
    End If
    For Each k In dict.Keys
        On Error Resume Next
        Dim v As Variant
        v = CallByName(dict(k), cleMetier, VbGet)
        If Err.Number = 0 Then
            If LCase$(CStr(v)) = LCase$(valeur) Then
                idxColl(indexKey) = k
                Set FindByKey = dict(k)
                Exit Function
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next k
End Function