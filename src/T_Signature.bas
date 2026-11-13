Attribute VB_Name = "T_Signature"
Option Explicit

Public Sub T_Signature()
    ResetCounters

    ' Crée un fichier factice, le signe, vérifie
    Dim tempFile As String
    tempFile = Environ$("TEMP") & "\test_sig_" & Format$(Now, "hhnnss") & ".xlsx"
    Dim f As Integer: f = FreeFile
    Open tempFile For Output As #f: Print #f, "faux xlsx": Close #f

    AssertTrue SignerRapport(tempFile, "Auteur"), "signature OK"
    AssertTrue Dir(tempFile & ".sig") <> "", "fichier .sig créé"
    AssertTrue VerifierSignature(tempFile), "signature vérifiée (intégrité)"

    ' Modification du fichier → signature invalide
    Dim f2 As Integer: f2 = FreeFile
    Open tempFile For Output As #f2: Print #f2, "modifié": Close #f2
    AssertFalse VerifierSignature(tempFile), "signature invalide après modification"

    ' Lecture de la signature
    Dim s As TSignature
    ' Re-signer d'abord
    SignerRapport tempFile, "Approbateur", "Validé"
    s = LireSignature(tempFile)
    AssertEquals "Approbateur", s.Role, "rôle lu"
    AssertEquals Environ$("USERNAME"), s.Utilisateur, "utilisateur lu"

    Kill tempFile: Kill tempFile & ".sig"
    PrintSummary "M_Signature"
End Sub