Attribute VB_Name = "FApi"
Option Explicit

Private Sub UserForm_Initialize()
    Caption = "API REST locale"
    cmdExporter.Caption = "Exporter JSON"
    cmdOuvrirDossier.Caption = "Ouvrir le dossier"
    cmdTester.Caption = "Tester"
    cmdClose.Caption = T("btn.close")
    txtPort.Text = "8080"
    lblInfo.Caption = "Les fichiers JSON sont générés dans :" & vbCrLf & _
        DataPath() & "api\" & vbCrLf & vbCrLf & _
        "Servez ce dossier avec le serveur Python fourni par install.ps1 :" & vbCrLf & _
        "python -m http.server 8080 --directory """ & DataPath() & "api"""
End Sub

Private Sub cmdExporter_Click()
    ExporterApiJson
    lblStatus.Caption = "Export JSON terminé."
End Sub

Private Sub cmdOuvrirDossier_Click()
    Shell "explorer.exe """ & DataPath() & "api""", vbNormalFocus
End Sub

Private Sub cmdTester_Click()
    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", "http://localhost:" & txtPort.Text & "/index.json", False
    http.Send
    If http.Status = 200 Then
        txtResultat.Text = http.responseText
        lblStatus.Caption = "Serveur accessible (HTTP 200)."
    Else
        lblStatus.Caption = "Serveur inaccessible (HTTP " & http.Status & ")."
    End If
    Exit Sub
ErrHandler:
    lblStatus.Caption = "Erreur : " & Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub