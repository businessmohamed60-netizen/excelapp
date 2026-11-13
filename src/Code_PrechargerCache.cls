Option Explicit

Public Sub PrechargerCache()
    Dim t0 As Single: t0 = Timer
    Dim cols As Variant, i As Long
    cols = Array("clients", "audits", "constats", "nc", "capa", "huitd", _
                 "amdec", "ppap", "apqp", "spc", "msa")
    For i = LBound(cols) To UBound(cols)
        On Error Resume Next
        g_CollMgr.LoadCollection CStr(cols(i))
        Err.Clear
        On Error GoTo 0
    Next i
    Log LogInfo, "M_Perf", "Préchargement cache en " & Format$(Timer - t0, "0.000") & " s"
End Sub

Public Function DiagnostiquerCache() As String
    Dim s As String, cols As Variant, i As Long
    cols = Array("clients", "audits", "constats", "nc", "capa", "huitd", _
                 "amdec", "ppap", "apqp", "spc", "msa")
    s = "=== Diagnostic du cache ===" & vbCrLf
    For i = LBound(cols) To UBound(cols)
        Dim d As Object
        On Error Resume Next
        Set d = g_CollMgr.LoadCollection(CStr(cols(i)))
        If Err.Number = 0 Then s = s & cols(i) & " : " & d.Count & " objet(s)" & vbCrLf Else s = s & cols(i) & " : non chargé" & vbCrLf
        Err.Clear
        On Error GoTo 0
    Next i
    DiagnostiquerCache = s
End Function