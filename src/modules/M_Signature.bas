Attribute VB_Name = "M_Signature"
Option Explicit

' ============================================================
' M_Signature — Signature électronique des rapports
' IATF §7.5.3.2 — maîtrise des enregistrements
' ============================================================

Public Type TSignature
    Utilisateur As String
    Poste As String
    DateHeure As Date
    HashFichier As String
    Role As String               ' Auteur / Vérificateur / Approbateur
    Commentaire As String
End Type

' Signe un rapport : enregistre le hash + métadonnées
Public Function SignerRapport(ByVal path As String, ByVal role As String, _
                              Optional ByVal commentaire As String = "") As Boolean
    On Error GoTo ErrHandler
    If Dir(path) = "" Then
        ShowError "Signature", "Fichier introuvable : " & path
        Exit Function
    End If

    Dim s As TSignature
    s.Utilisateur = Environ$("USERNAME")
    s.Poste = Environ$("COMPUTERNAME")
    s.DateHeure = Now
    s.HashFichier = HashFichier(path)
    s.Role = role
    s.Commentaire = commentaire

    ' Écriture du fichier .sig (JSON-like)
    Dim sigPath As String
    sigPath = path & ".sig"
    Dim f As Integer: f = FreeFile
    Open sigPath For Output As #f
    Print #f, "{"
    Print #f, "  ""utilisateur"": """ & s.Utilisateur & ""","
    Print #f, "  ""poste"": """ & s.Poste & ""","
    Print #f, "  ""dateHeure"": """ & Format$(s.DateHeure, "yyyy-mm-dd hh:nn:ss") & ""","
    Print #f, "  ""hash"": """ & s.HashFichier & ""","
    Print #f, "  ""role"": """ & s.Role & ""","
    Print #f, "  ""commentaire"": """ & Replace(s.Commentaire, """", "'") & """"
    Print #f, "}"
    Close #f

    ' Ajout d'une propriété personnalisée dans le classeur (si ouvert ou à ouvrir)
    AjouterProprieteSignature path, s

    ' Journalisation
    Journaliser "Signature", "rapport", path, ExtractFileName(path), "", "", role

    Log LogInfo, "M_Signature", "Rapport signé : " & path & " par " & s.Utilisateur
    SignerRapport = True
    Exit Function
ErrHandler:
    LogEx "M_Signature", "SignerRapport", Err
    SignerRapport = False
End Function

' Vérifie l'intégrité : le hash du fichier correspond-il à celui de la signature ?
Public Function VerifierSignature(ByVal path As String) As Boolean
    Dim sigPath As String, ligne As String, hashSig As String
    sigPath = path & ".sig"
    If Dir(sigPath) = "" Then VerifierSignature = False: Exit Function

    Dim f As Integer: f = FreeFile
    Open sigPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, ligne
        If InStr(ligne, """hash"":") > 0 Then
            hashSig = Mid$(ligne, InStr(ligne, """: """) + 5)
            hashSig = Replace(hashSig, """", "")
            hashSig = Replace(hashSig, ",", "")
            hashSig = Trim$(hashSig)
        End If
    Loop
    Close #f

    Dim hashActuel As String
    hashActuel = HashFichier(path)
    VerifierSignature = (hashSig = hashActuel)
End Function

' Lit la signature d'un rapport
Public Function LireSignature(ByVal path As String) As TSignature
    Dim s As TSignature, sigPath As String, ligne As String
    sigPath = path & ".sig"
    If Dir(sigPath) = "" Then Exit Function

    Dim f As Integer: f = FreeFile
    Open sigPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, ligne
        If InStr(ligne, """utilisateur"":") > 0 Then
            s.Utilisateur = ExtraireValeur(ligne)
        ElseIf InStr(ligne, """poste"":") > 0 Then
            s.Poste = ExtraireValeur(ligne)
        ElseIf InStr(ligne, """dateHeure"":") > 0 Then
            s.DateHeure = CDate(ExtraireValeur(ligne))
        ElseIf InStr(ligne, """hash"":") > 0 Then
            s.HashFichier = ExtraireValeur(ligne)
        ElseIf InStr(ligne, """role"":") > 0 Then
            s.Role = ExtraireValeur(ligne)
        ElseIf InStr(ligne, """commentaire"":") > 0 Then
            s.Commentaire = ExtraireValeur(ligne)
        End If
    Loop
    Close #f
    LireSignature = s
End Function

Private Function ExtraireValeur(ByVal ligne As String) As String
    Dim p As Long, v As String
    p = InStr(ligne, """: """)
    If p = 0 Then Exit Function
    v = Mid$(ligne, p + 5)
    v = Replace(v, """", "")
    v = Replace(v, ",", "")
    ExtraireValeur = Trim$(v)
End Function

' Ajoute une propriété personnalisée dans le classeur
Private Sub AjouterProprieteSignature(ByVal path As String, ByRef s As TSignature)
    On Error Resume Next
    Dim wb As Workbook, etaitOuvert As Boolean
    etaitOuvert = False
    ' Vérifier si le classeur est déjà ouvert
    Dim w As Workbook
    For Each w In Application.Workbooks
        If w.FullName = path Then
            Set wb = w
            etaitOuvert = True
            Exit For
        End If
    Next w
    If wb Is Nothing Then Set wb = Workbooks.Open(path, ReadOnly:=False)
    ' Supprimer les propriétés existantes
    On Error Resume Next
    wb.CustomDocumentProperties("IATF_Signature").Delete
    wb.CustomDocumentProperties("IATF_Signataire").Delete
    wb.CustomDocumentProperties("IATF_DateSignature").Delete
    wb.CustomDocumentProperties("IATF_Hash").Delete
    On Error GoTo 0
    wb.CustomDocumentProperties.Add "IATF_Signature", False, msoPropertyTypeString, s.Role
    wb.CustomDocumentProperties.Add "IATF_Signataire", False, msoPropertyTypeString, s.Utilisateur & " (" & s.Poste & ")"
    wb.CustomDocumentProperties.Add "IATF_DateSignature", False, msoPropertyTypeString, Format$(s.DateHeure, "yyyy-mm-dd hh:nn:ss")
    wb.CustomDocumentProperties.Add "IATF_Hash", False, msoPropertyTypeString, s.HashFichier
    wb.Save
    If Not etaitOuvert Then wb.Close False
End Sub

Private Function ExtractFileName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, "\")
    If p > 0 Then ExtractFileName = Mid$(path, p + 1) Else ExtractFileName = path
End Function