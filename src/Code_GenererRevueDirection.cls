Public Function GenererRevueDirection(ByVal annee As Long) As String
    Dim wb As Workbook, ws As Worksheet
    Dim tplPath As String
    tplPath = TemplatePath() & "Revue_Direction.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3003, "M_Report", "Template introuvable : " & tplPath
    End If

    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("RevueDirection")

    Dim kpi As Collection, i As Long, k As TKpi
    Set kpi = CalculerTousKpi(annee)

    ws.Range("RD_Annee").Value = annee
    ws.Range("RD_Date").Value = FormatIsoDate(Date)

    ' Remplissage KPI par nom
    For i = 1 To kpi.Count
        k = kpi(i)
        Select Case k.Nom
            Case "NC ouvertes": ws.Range("RD_NbNcOuvertes").Value = k.Valeur
            Case "CAPA en retard": ws.Range("RD_NbCapaRetard").Value = k.Valeur
            Case "8D ouverts": ws.Range("RD_Nb8dOuverts").Value = k.Valeur
            Case "Taux de clôture NC": ws.Range("RD_TauxClotureNc").Value = k.Valeur
            Case "Taux de clôture CAPA": ws.Range("RD_TauxClotureCapa").Value = k.Valeur
            Case "Programme : réalisés": ws.Range("RD_ProgRealises").Value = k.Valeur
            Case "Revues documentaires échues": ws.Range("RD_DocRevueEchue").Value = k.Valeur
            Case "Exigences non conformes": ws.Range("RD_ExigNonConf").Value = k.Valeur
        End Select
    Next i

    ' Tableau complet
    ws.Cells(20, 1).Value = "KPI"
    ws.Cells(20, 2).Value = "Valeur"
    ws.Cells(20, 3).Value = "Cible"
    ws.Cells(20, 4).Value = "Statut"
    ws.Cells(20, 5).Value = "Tendance"
    For i = 1 To kpi.Count
        k = kpi(i)
        ws.Cells(20 + i, 1).Value = k.Nom
        ws.Cells(20 + i, 2).Value = k.Valeur
        ws.Cells(20 + i, 3).Value = k.Cible
        ws.Cells(20 + i, 4).Value = k.Statut
        ws.Cells(20 + i, 5).Value = k.Tendance
    Next i

    Dim outPath As String
    outPath = AppDataPath() & "reports\Revue_Direction_" & annee & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererRevueDirection = outPath
End Function