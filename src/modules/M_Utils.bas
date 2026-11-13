Attribute VB_Name = "M_Utils"
Option Explicit

' ============================================================
' M_Utils — Helpers transverses (chemins, GUID, dates, XML)
' ============================================================

' ---- Chemins applicatifs ----
Public Function AppDataPath() As String
    Dim p As String
    p = Environ$("APPDATA") & "\IATFConsultantAssistant\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    AppDataPath = p
End Function

Public Function DataPath() As String
    Dim p As String
    p = AppDataPath() & "data\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    DataPath = p
End Function

Public Function TemplatePath() As String
    Dim p As String
    p = AppDataPath() & "templates\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    TemplatePath = p
End Function

Public Function LogPath() As String
    Dim p As String
    p = AppDataPath() & "logs\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    LogPath = p
End Function

Public Function BackupPath() As String
    Dim p As String
    p = AppDataPath() & "backups\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    BackupPath = p
End Function

Public Function ReportsPath() As String
    Dim p As String
    p = AppDataPath() & "reports\"
    If Dir(p, vbDirectory) = "" Then MkDir p
    ReportsPath = p
End Function

' ---- GUID v4 simplifié ----
Public Function NewGuid() As String
    Dim s As String, i As Long, c As String
    Randomize
    For i = 1 To 32
        If i = 9 Or i = 13 Or i = 17 Or i = 21 Then s = s & "-"
        c = Hex(Int(Rnd() * 16))
        s = s & LCase(c)
    Next i
    NewGuid = s
End Function

' ---- Dates ISO 8601 ----
Public Function FormatIsoDate(ByVal d As Date) As String
    FormatIsoDate = Format$(d, "yyyy-mm-dd")
End Function

Public Function ParseIsoDate(ByVal s As String) As Date
    On Error GoTo Fallback
    Dim p() As String
    p = Split(s, "-")
    If UBound(p) < 2 Then GoTo Fallback
    ParseIsoDate = DateSerial(CInt(p(0)), CInt(p(1)), CInt(p(2)))
    Exit Function
Fallback:
    ParseIsoDate = 0
End Function

' ---- Échappement XML ----
Public Function EscapeXml(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    r = Replace(r, "'", "&apos;")
    EscapeXml = r
End Function

' ---- Boîtes de dialogue ----
Public Sub ShowError(ByVal titre As String, ByVal message As String)
    MsgBox message, vbCritical, titre
End Sub

Public Sub ShowInfo(ByVal titre As String, ByVal message As String)
    MsgBox message, vbInformation, titre
End Sub

Public Function PromptYesNo(ByVal titre As String, ByVal message As String) As Boolean
    PromptYesNo = (MsgBox(message, vbYesNo + vbQuestion, titre) = vbYes)
End Function