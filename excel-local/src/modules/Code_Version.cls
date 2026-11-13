Option Explicit

Private m_Version As String
Private m_Language As String
Private m_InstallDate As Date
Private m_LastUser As String
Private m_TrialExpiry As Date
Private m_LicenseKey As String

Public Property Get Version() As String: Version = m_Version: End Property
Public Property Let Version(ByVal v As String): m_Version = v: End Property
Public Property Get Language() As String: Language = m_Language: End Property
Public Property Let Language(ByVal v As String): m_Language = v: End Property
Public Property Get InstallDate() As Date: InstallDate = m_InstallDate: End Property
Public Property Let InstallDate(ByVal v As Date): m_InstallDate = v: End Property
Public Property Get LastUser() As String: LastUser = m_LastUser: End Property
Public Property Let LastUser(ByVal v As String): m_LastUser = v: End Property
Public Property Get TrialExpiry() As Date: TrialExpiry = m_TrialExpiry: End Property
Public Property Let TrialExpiry(ByVal v As Date): m_TrialExpiry = v: End Property
Public Property Get LicenseKey() As String: LicenseKey = m_LicenseKey: End Property
Public Property Let LicenseKey(ByVal v As String): m_LicenseKey = v: End Property

Public Sub InitDefault()
    m_Version = "0.1.0"
    m_Language = "FR"
    m_InstallDate = Date
    m_LastUser = Environ$("USERNAME")
    m_TrialExpiry = Date + 30
    m_LicenseKey = ""
End Sub

Public Sub Save()
    Dim dom As Object, root As Object, cfgFile As String
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.loadXML "<config/>"
    Set root = dom.documentElement
    AppendText dom, root, "version", m_Version
    AppendText dom, root, "language", m_Language
    AppendText dom, root, "installDate", FormatIsoDate(m_InstallDate)
    AppendText dom, root, "lastUser", m_LastUser
    AppendText dom, root, "trialExpiry", FormatIsoDate(m_TrialExpiry)
    AppendText dom, root, "licenseKey", m_LicenseKey
    cfgFile = DataPath() & "config.xml"
    dom.save cfgFile
End Sub

Public Function Load() As Boolean
    Dim dom As Object, cfgFile As String
    cfgFile = DataPath() & "config.xml"
    If Dir(cfgFile) = "" Then Load = False: Exit Function
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(cfgFile) Then Load = False: Exit Function
    m_Version = GetText(dom, "version")
    m_Language = GetText(dom, "language")
    m_InstallDate = ParseIsoDate(GetText(dom, "installDate"))
    m_LastUser = GetText(dom, "lastUser")
    m_TrialExpiry = ParseIsoDate(GetText(dom, "trialExpiry"))
    m_LicenseKey = GetText(dom, "licenseKey")
    Load = True
End Function

Private Sub AppendText(ByVal dom As Object, ByVal parent As Object, ByVal tag As String, ByVal val As String)
    Dim e As Object
    Set e = dom.createElement(tag)
    e.Text = val
    parent.appendChild e
End Sub

Private Function GetText(ByVal dom As Object, ByVal tag As String) As String
    Dim n As Object
    Set n = dom.getElementsByTagName(tag)
    If n.Length > 0 Then GetText = n.Item(0).Text Else GetText = ""
End Function