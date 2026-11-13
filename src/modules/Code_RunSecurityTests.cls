Option Explicit

Public Sub RunSecurityTests()
    ResetCounters
    Dim h As String: h = SHA256Hex("abc")
    AssertTrue Len(h) = 64, "SHA-256 longueur 64"
    Dim h1 As String, h2 As String
    h1 = HMACSHA256Hex("cle", "message"): h2 = HMACSHA256Hex("cle", "message")
    AssertEquals h1, h2, "HMAC déterministe"
    SignerFichiersData
    AssertTrue VerifierIntegrite(), "intégrité OK après signature"
    AssertTrue SauvegarderMaintenant(), "sauvegarde OK"
    PrintSummary "Sécurité"
End Sub