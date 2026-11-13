Public Function GenererFicheMoyen(ByVal moyen As CMoyenMesure) As String
    Dim wb As Workbook, ws As Worksheet, tplPath As String
    tplPath = TemplatePath() & "Fiche_Moyen_Mesure.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3007, "M_Report", "Template introuvable : " & tplPath
    End If
    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("Fiche")
    ws.Range("FM_Code").Value = moyen.Code
    ws.Range("FM_Designation").Value = moyen.Designation
    ws.Range("FM_Type").Value = moyen.TypeMoyen
    ws.Range("FM_Marque").Value = moyen.Marque
    ws.Range("FM_Modele").Value = moyen.Modele
    ws.Range("FM_NumeroSerie").Value = moyen.NumeroSerie
    ws.Range("FM_Localisation").Value = moyen.Localisation
    ws.Range("FM_Etendue").Value = moyen.EtendueMesure
    ws.Range("FM_Resolution").Value = moyen.Resolution
    ws.Range("FM_Incertitude").Value = moyen.Incertitude
    ws.Range("FM_Periodicite").Value = moyen.PeriodiciteMois
    If IsDate(moyen.DernierEtalonnage) Then ws.Range("FM_DernierEtal").Value = FormatIsoDate(moyen.DernierEtalonnage)
    If IsDate(moyen.ProchainEtalonnage) Then ws.Range("FM_ProchainEtal").Value = FormatIsoDate(moyen.ProchainEtalonnage)
    ws.Range("FM_Statut").Value = moyen.Statut
    ws.Range("FM_Responsable").Value = moyen.Responsable

    Dim etals As Object, k As Variant, row As Long
    Set etals = g_CollMgr.LoadCollection("etalonnage")
    row = 23
    For Each k In etals.Keys
        If etals(k).MoyenId = moyen.Id Then
            ws.Cells(row, 1).Value = FormatIsoDate(etals(k).DateEtalonnage)
            ws.Cells(row, 2).Value = etals(k).TypeOperation
            ws.Cells(row, 3).Value = etals(k).Organisme
            ws.Cells(row, 4).Value = etals(k).CertificatNumero
            ws.Cells(row, 5).Value = etals(k).Resultat
            ws.Cells(row, 6).Value = etals(k).IncertitudeRelevee
            If IsDate(etals(k).ProchaineEcheance) Then ws.Cells(row, 7).Value = FormatIsoDate(etals(k).ProchaineEcheance)
            ws.Cells(row, 8).Value = etals(k).Cout
            row = row + 1
        End If
    Next k

    Dim outPath As String
    outPath = AppDataPath() & "reports\FicheMoyen_" & moyen.Code & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererFicheMoyen = outPath
End Function