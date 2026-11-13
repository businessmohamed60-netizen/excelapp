Option Explicit

Private m_passed As Long
Private m_failed As Long
Private m_failures As String

Public Function GetFailedCount() As Long: GetFailedCount = m_failed: End Function
Public Function GetPassedCount() As Long: GetPassedCount = m_passed: End Function

Public Sub AssertTrue(ByVal cond As Boolean, ByVal label As String)
    If cond Then
        m_passed = m_passed + 1
    Else
        m_failed = m_failed + 1
        m_failures = m_failures & vbCrLf & "  X " & label
        Debug.Print "FAIL: " & label
    End If
End Sub

Public Sub AssertFalse(ByVal cond As Boolean, ByVal label As String)
    AssertTrue Not cond, label
End Sub

Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim ok As Boolean
    If IsNull(expected) Or IsNull(actual) Then
        ok = (IsNull(expected) And IsNull(actual))
    ElseIf IsObject(expected) Or IsObject(actual) Then
        ok = (expected Is actual)
    Else
        ok = (CStr(expected) = CStr(actual))
    End If
    If ok Then
        m_passed = m_passed + 1
    Else
        m_failed = m_failed + 1
        Dim msg As String
        msg = label & " (attendu=" & CStr(expected) & ", obtenu=" & CStr(actual) & ")"
        m_failures = m_failures & vbCrLf & "  X " & msg
        Debug.Print "FAIL: " & msg
    End If
End Sub

Public Sub AssertBetween(ByVal v As Double, ByVal lo As Double, ByVal hi As Double, ByVal label As String)
    AssertTrue (v >= lo And v <= hi), label & " (valeur=" & v & ")"
End Sub

Public Sub AssertNotNothing(ByVal obj As Object, ByVal label As String)
    AssertTrue Not (obj Is Nothing), label
End Sub

Public Sub ResetCounters()
    m_passed = 0
    m_failed = 0
    m_failures = ""
End Sub

Public Sub PrintSummary(ByVal suiteName As String)
    Debug.Print String(60, "=")
    Debug.Print "Suite : " & suiteName
    Debug.Print "  Passés  : " & m_passed
    Debug.Print "  Échoués : " & m_failed
    If m_failed > 0 Then Debug.Print "  Détails :" & m_failures
    Debug.Print String(60, "=")
End Sub

Public Sub RunAllTests()
    ResetCounters
    Debug.Print vbCrLf & "DEMARRAGE DES TESTS — " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    T_Repository
    T_Client
    T_Audit
    T_Constat
    RunUtilsTests
    RunQualiteTests
    RunCoreToolsTests
    RunUxTests
    RunSecurityTests
    PrintSummary "ALL"
    If m_failed = 0 Then
        MsgBox "Tous les tests sont passés (" & m_passed & " assertions).", vbInformation, "Tests OK"
    Else
        MsgBox m_failed & " test(s) échoué(s) sur " & (m_passed + m_failed) & ". Voir la fenêtre Exécution.", vbExclamation, "Tests KO"
    End If
End Sub

' ---- Suites de base ----
Public Sub T_Repository()
    ResetCounters
    On Error Resume Next
    Kill DataPath() & "test_repo.xml"
    On Error GoTo 0
    Dim c As CClient: Set c = New CClient: c.InitNew
    c.Code = "TST": c.RaisonSociale = "Client de test": c.Actif = True
    Dim repo As CRepository: Set repo = New CRepository
    repo.Save c, "test_repo"
    AssertTrue Dir(DataPath() & "test_repo.xml") <> "", "fichier XML créé"
    Dim c2 As CClient: Set c2 = New CClient
    AssertTrue repo.LoadById(c.Id, "test_repo", c2), "LoadById True"
    AssertEquals c.Code, c2.Code, "code restitué"
    AssertTrue repo.DeleteById(c.Id, "test_repo"), "DeleteById True"
    On Error Resume Next
    Kill DataPath() & "test_repo.xml"
    Kill DataPath() & "test_repo.xml.bak"
    On Error GoTo 0
    PrintSummary "CRepository"
End Sub

Public Sub T_Client()
    ResetCounters
    Dim c As CClient: Set c = New CClient: c.InitNew
    c.Code = "ABC": c.RaisonSociale = "Société ABC"
    AssertTrue c.Validate(), "validation OK"
    AssertTrue Len(c.Id) > 0, "Id généré"
    AssertEquals "IATF 16949", c.Norme, "norme par défaut"
    Dim c2 As CClient: Set c2 = New CClient: c2.InitNew
    c2.Code = "XYZ": c2.RaisonSociale = ""
    AssertFalse c2.Validate(), "validation KO si RS vide"
    PrintSummary "CClient"
End Sub

Public Sub T_Audit()
    ResetCounters
    Dim a As CAudit: Set a = New CAudit: a.InitNew
    AssertEquals "Planifié", a.Statut, "statut par défaut"
    a.ClientId = NewGuid(): a.DateDebut = Date: a.DateFin = Date + 3
    AssertTrue a.Validate(), "audit valide"
    PrintSummary "CAudit"
End Sub

Public Sub T_Constat()
    ResetCounters
    Dim k As CConstat: Set k = New CConstat: k.InitNew
    AssertTrue Len(k.Id) > 0, "Id généré"
    AssertEquals "Observation", k.Classification, "classification par défaut"
    PrintSummary "CConstat"
End Sub