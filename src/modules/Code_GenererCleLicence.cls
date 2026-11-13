Option Explicit

Private Const TRIAL_DAYS As Long = 30
Private Const GRACE_DAYS As Long = 30

Public Enum TypeLicence
    LicenceTrial = 0
    LicenceStandard = 1
    LicencePremium = 2
    LicenceExpiree = 3
End Enum

Public Function GenererCleLicence(ByVal machineId As String, ByVal dateExpiration As Date, ByVal edition As String) As String
    Dim payload As String, sig As String, core As String, res As String, i As Long
    payload = machineId & "|" & Format$(dateExpiration, "yyyymmdd") & "|" & edition
    sig = UCase$(HMACSHA256Hex(LicenseMasterKey(), payload))
    core = Left$(sig, 20)
    res = "IATF"
    For i = 1 To 20 Step 5: res = res & "-" & Mid$(core, i, 5): Next i
    GenererCleLicence = res
End Function

Public Function VerifierCleLicence(ByVal cle As String, ByVal machineId As String, _
                                   ByRef dateExpiration As Date, ByRef edition As String) As TypeLicence
    On Error GoTo ErrHandler
    Dim parts() As String
    parts = Split(Replace(UCase$(Trim$(cle)), " ", ""), "-")
    If UBound(parts) < 4 Or parts(0) <> "IATF" Then VerifierCleLicence = LicenceExpiree: Exit Function
    Dim core As String
    core = parts(1) & parts(2) & parts(3) & parts(4)
    Dim d As Date, ed As String, edList As Variant
    edList = Array("STANDARD", "PREMIUM")
    Dim annee As Long, mois As Long, jour As Long
    For Each ed In edList
        For annee = Year(Date) To Year(Date) + 5
            For mois = 1 To 12
                For jour = 1 To 28 Step 7
                    d = DateSerial(annee, mois, jour)
                    Dim payload As String
                    payload = machineId & "|" & Format$(d, "yyyymmdd") & "|" & ed
                    If Left$(UCase$(HMACSHA256Hex(LicenseMasterKey(), payload)), 20) = core Then
                        dateExpiration = d: edition = ed
                        If d < Date Then VerifierCleLicence = LicenceExpiree
                        ElseIf ed = "PREMIUM" Then VerifierCleLicence = LicencePremium
                        Else VerifierCleLicence = LicenceStandard
                        Exit Function
                    End If
                Next jour
            Next mois
        Next annee
    Next ed
    VerifierCleLicence = LicenceExpiree
    Exit Function
ErrHandler:
    LogEx "M_License", "VerifierCleLicence", Err
    VerifierCleLicence = LicenceExpiree
End Function

Private Function LicenseMasterKey() As String
    LicenseMasterKey = "IATF" & Chr$(72) & "Consult" & Chr$(64) & "nt2026" & "Premium" & Chr$(33)
End Function

Public Function EtatLicence() As TypeLicence
    If g_Config Is Nothing Then EtatLicence = LicenceExpiree: Exit Function
    If Len(g_Config.LicenseKey) > 0 Then
        Dim d As Date, ed As String, t As TypeLicence
        t = VerifierCleLicence(g_Config.LicenseKey, MachineFingerprint(), d, ed)
        If t <> LicenceExpiree Then EtatLicence = t: Exit Function
    End If
    Dim joursRestants As Long
    joursRestants = DateDiff("d", Date, g_Config.TrialExpiry)
    If joursRestants >= 0 Then EtatLicence = LicenceTrial: Exit Function
    If joursRestants >= -GRACE_DAYS Then EtatLicence = LicenceTrial: Exit Function
    EtatLicence = LicenceExpiree
End Function

Public Function JoursRestantsTrial() As Long
    If g_Config Is Nothing Then Exit Function
    JoursRestantsTrial = DateDiff("d", Date, g_Config.TrialExpiry)
End Function

Public Function EnregistrerCleLicence(ByVal cle As String) As Boolean
    On Error GoTo ErrHandler
    Dim d As Date, ed As String
    Dim t As TypeLicence
    t = VerifierCleLicence(cle, MachineFingerprint(), d, ed)
    If t = LicenceExpiree Then EnregistrerCleLicence = False: Exit Function
    g_Config.LicenseKey = UCase$(Trim$(cle))
    g_Config.Save
    Log LogInfo, "M_License", "Licence enregistrée : " & ed & " jusqu'au " & FormatIsoDate(d)
    EnregistrerCleLicence = True
    Exit Function
ErrHandler:
    LogEx "M_License", "EnregistrerCleLicence", Err
    EnregistrerCleLicence = False
End Function

Public Function MessageEtatLicence() As String
    Select Case EtatLicence()
        Case LicencePremium: MessageEtatLicence = "Licence Premium — active"
        Case LicenceStandard: MessageEtatLicence = "Licence Standard — active"
        Case LicenceTrial
            Dim j As Long: j = JoursRestantsTrial()
            If j >= 0 Then MessageEtatLicence = "Période d'essai — " & j & " jour(s) restant(s)"
            Else MessageEtatLicence = "Période d'essai expirée — grâce (" & (-j) & " j)"
        Case LicenceExpiree: MessageEtatLicence = "Licence expirée — enregistrez une clé valide"
    End Select
End Function