Attribute VB_Name = "M_Storage"
Option Explicit

' ============================================================
' M_Storage — Abstraction du stockage et verrouillage optimiste
' ============================================================

' Chemin racine des données, selon le mode configuré
Public Function DataPath() As String
    Dim mode As String, p As String
    mode = g_Config.StorageMode
    Select Case mode
        Case "OneDrive"
            p = Environ$("OneDrive")
            If Len(p) = 0 Then p = Environ$("USERPROFILE") & "\OneDrive"
            p = p & "\IATFConsultantAssistant\data\"
        Case "SharePoint"
            p = g_Config.StoragePath
            If Right$(p, 1) <> "\" Then p = p & "\"
        Case "Reseau"
            p = g_Config.StoragePath
            If Right$(p, 1) <> "\" Then p = p & "\"
        Case Else ' Local
            p = Environ$("APPDATA") & "\IATFConsultantAssistant\data\"
    End Select
    If Dir(p, vbDirectory) = "" Then MkDir p
    DataPath = p
End Function

' Chemin des templates (toujours local, jamais partagé)
Public Function TemplatePath() As String
    TemplatePath = Environ$("APPDATA") & "\IATFConsultantAssistant\templates\"
End Function

' Chemin des rapports (toujours local)
Public Function AppDataPath() As String
    AppDataPath = Environ$("APPDATA") & "\IATFConsultantAssistant\"
End Function

' ----- Verrouillage optimiste -----

' Calcule un hash SHA-256 (hex) d'un fichier
Public Function HashFichier(ByVal path As String) As String
    If Dir(path) = "" Then HashFichier = "": Exit Function
    Dim stream As Object, bytes() As Byte, i As Long, s As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binary
    stream.Open
    stream.LoadFromFile path
    bytes = stream.Read
    stream.Close
    ' Hash simple (pas cryptographique) — suffisant pour détecter un conflit
    Dim h As Long: h = 5381
    For i = LBound(bytes) To UBound(bytes)
        h = ((h * 33) Xor bytes(i)) And &H7FFFFFFF
    Next i
    HashFichier = Format$(h, "00000000")
End Function

' Vérifie qu'un fichier n'a pas été modifié depuis le dernier chargement
Public Function VerifierConflit(ByVal path As String, ByVal hashAttendu As String) As Boolean
    Dim h As String
    h = HashFichier(path)
    VerifierConflit = (Len(hashAttendu) > 0) And (h <> hashAttendu)
End Function

' Enregistre le hash courant d'un fichier dans le cache de session
Public Sub EnregistrerHash(ByVal collection As String, ByVal path As String)
    Dim h As String: h = HashFichier(path)
    g_SessionHashes(collection) = h
End Sub

' Récupère le hash attendu pour une collection
Public Function HashAttendu(ByVal collection As String) As String
    If g_SessionHashes.Exists(collection) Then
        HashAttendu = g_SessionHashes(collection)
    End If
End Function

' ----- Test de connectivité -----

Public Function VerifierStockage() As Boolean
    On Error GoTo ErrHandler
    Dim p As String, testFile As String
    p = DataPath()
    testFile = p & "_test_" & Format$(Now, "yyyymmddhhnnss") & ".tmp"
    Dim f As Integer: f = FreeFile
    Open testFile For Output As #f
    Print #f, "test"
    Close #f
    Kill testFile
    VerifierStockage = True
    Exit Function
ErrHandler:
    VerifierStockage = False
End Function