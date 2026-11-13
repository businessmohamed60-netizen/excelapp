Attribute VB_Name = "M_Db"
Option Explicit

' ============================================================
' M_Db — Couche d'accès données (SQLite / Access via ADODB)
' ============================================================

Private m_Conn As Object            ' ADODB.Connection
Private m_Transaction As Boolean

Public Function DbOpen() As Boolean
    On Error GoTo ErrHandler
    If Not m_Conn Is Nothing Then
        If m_Conn.State = 1 Then DbOpen = True: Exit Function
    End If

    Set m_Conn = CreateObject("ADODB.Connection")
    Dim connStr As String
    Select Case g_Config.DbEngine
        Case "SQLite"
            connStr = "DRIVER={SQLite3 ODBC Driver};Database=" & g_Config.DbPath & ";"
        Case "Access"
            connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & g_Config.DbPath & ";"
        Case Else
            Err.Raise vbObjectError + 5001, "M_Db", "Moteur non supporté : " & g_Config.DbEngine
    End Select
    m_Conn.Open connStr
    m_Conn.CommandTimeout = 30
    DbOpen = True
    Exit Function
ErrHandler:
    LogEx "M_Db", "DbOpen", Err
    DbOpen = False
End Function

Public Sub DbClose()
    On Error Resume Next
    If Not m_Conn Is Nothing Then
        If m_Conn.State = 1 Then m_Conn.Close
    End If
    Set m_Conn = Nothing
End Sub

Public Function DbExecute(ByVal sql As String, Optional ByVal params As Variant) As Long
    On Error GoTo ErrHandler
    If Not DbOpen() Then Exit Function
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = m_Conn
    cmd.CommandText = sql
    cmd.CommandType = 1 ' adCmdText
    If Not IsMissing(params) Then
        Dim i As Long
        For i = LBound(params) To UBound(params)
            cmd.Parameters.Append cmd.CreateParameter("p" & i, 200, 1, 255, params(i)) ' adVarChar
        Next i
    End If
    cmd.Execute , , 128 ' adExecuteNoRecords
    DbExecute = cmd.Parameters.Count
    Exit Function
ErrHandler:
    LogEx "M_Db", "DbExecute", Err
    Err.Raise Err.Number, "M_Db.DbExecute", Err.Description & " SQL: " & sql
End Function

Public Function DbQuery(ByVal sql As String, Optional ByVal params As Variant) As Object
    On Error GoTo ErrHandler
    If Not DbOpen() Then Exit Function
    Dim rs As Object, cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = m_Conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    If Not IsMissing(params) Then
        Dim i As Long
        For i = LBound(params) To UBound(params)
            cmd.Parameters.Append cmd.CreateParameter("p" & i, 200, 1, 255, params(i))
        Next i
    End If
    Set rs = cmd.Execute
    Set DbQuery = rs
    Exit Function
ErrHandler:
    LogEx "M_Db", "DbQuery", Err
    Err.Raise Err.Number, "M_Db.DbQuery", Err.Description & " SQL: " & sql
End Function

Public Sub DbBegin()
    If Not DbOpen() Then Exit Sub
    m_Conn.BeginTrans
    m_Transaction = True
End Sub

Public Sub DbCommit()
    If m_Transaction And Not m_Conn Is Nothing Then
        m_Conn.CommitTrans
        m_Transaction = False
    End If
End Sub

Public Sub DbRollback()
    If m_Transaction And Not m_Conn Is Nothing Then
        m_Conn.RollbackTrans
        m_Transaction = False
    End If
End Sub

' Test de connectivité
Public Function DbTest() As Boolean
    On Error GoTo ErrHandler
    If Not DbOpen() Then DbTest = False: Exit Function
    Dim rs As Object
    Set rs = DbQuery("SELECT 1 AS test")
    If Not rs.EOF Then DbTest = True
    rs.Close
    Exit Function
ErrHandler:
    DbTest = False
End Function

' Retourne la liste des tables
Public Function DbTables() As Collection
    Dim result As New Collection, rs As Object
    Set rs = DbQuery("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    Do While Not rs.EOF
        result.Add rs("name")
        rs.MoveNext
    Loop
    rs.Close
    Set DbTables = result
End Function