Option Explicit

Private Const MANIFEST_FILE As String = "manifest.sig"

Private Function HashFichiersData() As String
    Dim fso As Object, folder As Object, file As Object
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(DataPath()) Then HashFichiersData = "EMPTY": Exit Function
    Set folder = fso.GetFolder(DataPath())
    Dim noms As Object: Set noms = CreateObject("Scripting.Dictionary")
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "xml" Then noms.Add file.Name, file.Path
    Next file
    Dim keys As Variant, i As Long, j As Long, tmp As String
    keys = noms.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i
    For i = LBound(keys) To UBound(keys)
        s = s & keys(i) & ":" & SHA256FichierHex(noms(keys(i))) & vbLf
    Next i
    HashFichiersData = SHA256Hex(s)
End Function

Public Sub SignerFichiersData()
    On Error GoTo ErrHandler
    Dim manifestPath As String, signature As String, f As Integer
    manifestPath = DataPath() & MANIFEST_FILE
    signature = HMACSHA256Hex(MachineFingerprint(), HashFichiersData())
    f = FreeFile
    Open manifestPath For Output Lock Write As #f
    Print #f, "version=1.0"
    Print #f, "date=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, "hash=" & HashFichiersData()
    Print #f, "signature=" & signature
    Close #f
    Log LogInfo, "M_Security", "Manifeste signé"
    Exit Sub
ErrHandler:
    LogEx "M_Security", "SignerFichiersData", Err
End Sub

Public Function VerifierIntegrite() As Boolean
    On Error GoTo ErrHandler
    Dim manifestPath As String
    manifestPath = DataPath() & MANIFEST_FILE
    If Dir(manifestPath) = "" Then
        SignerFichiersData
        VerifierIntegrite = True
        Exit Function
    End If
    Dim f As Integer, ligne As String, hashAttendu As String, sigAttendue As String
    f = FreeFile
    Open manifestPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, ligne
        If Left$(ligne, 5) = "hash=" Then hashAttendu = Mid$(ligne, 6)
        If Left$(ligne, 10) = "signature=" Then sigAttendue = Mid$(ligne, 11)
    Loop
    Close #f
    Dim hashActuel As String
    hashActuel = HashFichiersData()
    If hashActuel <> hashAttendu Then
        Log LogWarn, "M_Security", "Hash modifié"
        VerifierIntegrite = False
        Exit Function
    End If
    Dim sigCalculee As String
    sigCalculee = HMACSHA256Hex(MachineFingerprint(), hashActuel)
    If sigCalculee <> sigAttendue Then
        Log LogWarn, "M_Security", "Signature invalide"
        VerifierIntegrite = False
        Exit Function
    End If
    VerifierIntegrite = True
    Exit Function
ErrHandler:
    LogEx "M_Security", "VerifierIntegrite", Err
    VerifierIntegrite = False
End Function

Public Function MachineFingerprint() As String
    MachineFingerprint = SHA256Hex(Environ$("COMPUTERNAME") & "|" & Environ$("USERNAME") & "|" & Environ$("USERDOMAIN"))
End Function