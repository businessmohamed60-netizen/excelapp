Attribute VB_Name = "M_PowerBi"
Option Explicit

' ============================================================
' M_PowerBi — Connecteur Power BI
' ============================================================

Public Function GenererPbids() As String
    Dim path As String
    path = AppDataPath() & "reports\IATF_Data.pbids"
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, "{"
    Print #f, "  ""version"": ""0.1"","
    Print #f, "  ""connections"": [{"
    Print #f, "    ""details"": {"
    Print #f, "      ""protocol"": ""http"","
    Print #f, "      ""address"": {"
    Print #f, "        ""host"": ""localhost"","
    Print #f, "        ""port"": 8080"
    Print #f, "      },"
    Print #f, "      ""path"": ""/index.json"""
    Print #f, "    },"
    Print #f, "    ""mode"": ""Import"""
    Print #f, "  }]"
    Print #f, "}"
    Close #f
    GenererPbids = path
End Function

Public Function GenererPowerQuery() As String
    ' Génère un script M (Power Query) qui charge tous les endpoints
    Dim m As String
    m = "let" & vbCrLf
    m = m & "  BaseUrl = ""http://localhost:8080""," & vbCrLf
    m = m & "  Endpoints = {""clients"", ""audits"", ""nc"", ""capa"", ""huitd"", ""amdec"", ""moyenmesure"", ""etudemsa""}," & vbCrLf
    m = m & "  Fetch = (name) => Json.Document(Web.Contents(BaseUrl, [RelativePath=name & "".json""]))," & vbCrLf
    m = m & "  Result = List.Accumulate(Endpoints, [], (state, current) => " & vbCrLf
    m = m & "    state & {[Name=current, Data=Fetch(current)]})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "  Result"
    Dim path As String
    path = AppDataPath() & "reports\IATF_PowerQuery.txt"
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, m
    Close #f
    GenererPowerQuery = path
End Function