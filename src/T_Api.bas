Attribute VB_Name = "T_Api"
Option Explicit

Public Sub T_Api()
    ResetCounters

    ' Export JSON
    ExporterApiJson

    Dim apiDir As String
    apiDir = DataPath() & "api\"
    AssertTrue Dir(apiDir & "index.json") <> "", "index.json généré"
    AssertTrue Dir(apiDir & "clients.json") <> "", "clients.json généré"
    AssertTrue Dir(apiDir & "nc.json") <> "", "nc.json généré"

    ' Lire le JSON et vérifier qu'il est bien formé (première ligne = "[")
    Dim f As Integer: f = FreeFile
    Dim ligne As String
    Open apiDir & "clients.json" For Input As #f
    Line Input #f, ligne
    Close #f
    AssertEquals "[", Trim$(ligne), "clients.json commence par ["

    ' Nettoyage
    Kill apiDir & "*.json"

    PrintSummary "M_Api"
End Sub