Option Explicit

Private Const MAX_SIZE_KB As Long = 2048

Public Enum LogLevel
    LogDebug = 0
    LogInfo = 1
    LogWarn = 2
    LogError = 3
End Enum

Public Sub Log(ByVal niveau As LogLevel, ByVal module As String, ByVal message As String)
    On Error Resume Next
    Dim f As Integer, fname As String, line As String
    fname = LogPath() & "iatf_" & Format$(Date, "yyyymmdd") & ".log"
    If Dir(fname) <> "" Then
        If FileLen(fname) > MAX_SIZE_KB * 1024 Then
            Name fname As Replace(fname, ".log", "_" & Format$(Now, "hhmmss") & ".log")
        End If
    End If
    f = FreeFile
    Open fname For Append As #f
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & LevelName(niveau) & " | " & module & " | " & message
    Print #f, line
    Close #f
End Sub

Private Function LevelName(ByVal n As LogLevel) As String
    Select Case n
        Case LogDebug: LevelName = "DEBUG"
        Case LogInfo:  LevelName = "INFO "
        Case LogWarn:  LevelName = "WARN "
        Case LogError: LevelName = "ERROR"
    End Select
End Function

Public Sub LogEx(ByVal module As String, ByVal contexte As String, ByVal ex As Object)
    Log LogError, module, contexte & " :: " & ex.Number & " - " & ex.Description
End Sub