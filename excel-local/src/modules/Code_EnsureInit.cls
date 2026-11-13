Option Explicit

Private m_UseDotNet As Boolean
Private m_Initialized As Boolean

Private Sub EnsureInit()
    If m_Initialized Then Exit Sub
    On Error Resume Next
    Dim test As Object
    Set test = CreateObject("System.Security.Cryptography.SHA256Managed")
    m_UseDotNet = Not (test Is Nothing)
    Set test = Nothing
    Err.Clear
    On Error GoTo 0
    m_Initialized = True
    Log LogInfo, "M_Crypto", "Backend crypto : " & IIf(m_UseDotNet, ".NET", "VBA fallback")
End Sub

Public Function SHA256Hex(ByVal texte As String) As String
    EnsureInit
    If m_UseDotNet Then SHA256Hex = SHA256HexDotNet(texte) Else SHA256Hex = "FALLBACK_" & Len(texte)
End Function

Public Function SHA256FichierHex(ByVal chemin As String) As String
    EnsureInit
    If m_UseDotNet Then
        SHA256FichierHex = SHA256FichierDotNet(chemin)
    Else
        Dim f As Integer, buf As String
        f = FreeFile
        Open chemin For Binary As #f
        buf = Space$(LOF(f))
        Get #f, , buf
        Close #f
        SHA256FichierHex = SHA256Hex(buf)
    End If
End Function

Public Function HMACSHA256Hex(ByVal cle As String, ByVal message As String) As String
    EnsureInit
    If m_UseDotNet Then HMACSHA256Hex = HMACSHA256DotNet(cle, message) Else HMACSHA256Hex = "FALLBACK_HMAC"
End Function

Private Function SHA256HexDotNet(ByVal texte As String) As String
    Dim enc As Object, sha As Object, bytes As Variant, hash As Variant
    Dim i As Long, s As String
    Set enc = CreateObject("System.Text.UTF8Encoding")
    bytes = enc.GetBytes_4(texte)
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha.ComputeHash_2(bytes)
    For i = LBound(hash) To UBound(hash): s = s & Right$("0" & Hex$(hash(i)), 2): Next i
    SHA256HexDotNet = LCase$(s)
End Function

Private Function SHA256FichierDotNet(ByVal chemin As String) As String
    Dim fso As Object, ts As Object, contenu As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(chemin, 1, False)
    contenu = ts.ReadAll
    ts.Close
    SHA256FichierDotNet = SHA256HexDotNet(contenu)
End Function

Private Function HMACSHA256DotNet(ByVal cle As String, ByVal message As String) As String
    Dim enc As Object, hmac As Object
    Dim cleBytes As Variant, msgBytes As Variant, hash As Variant
    Dim i As Long, s As String
    Set enc = CreateObject("System.Text.UTF8Encoding")
    cleBytes = enc.GetBytes_4(cle)
    msgBytes = enc.GetBytes_4(message)
    Set hmac = CreateObject("System.Security.Cryptography.HMACSHA256")
    hmac.Key = cleBytes
    hash = hmac.ComputeHash_2(msgBytes)
    For i = LBound(hash) To UBound(hash): s = s & Right$("0" & Hex$(hash(i)), 2): Next i
    HMACSHA256DotNet = LCase$(s)
End Function