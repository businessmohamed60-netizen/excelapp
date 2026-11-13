Public Function GenererRapportAmdec(ByVal amdec As CAMDEC) As String
    Dim wb As Workbook, ws As Worksheet, tplPath As String
    tplPath = TemplatePath() & "AMDEC_Processus.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3005, "M_Report", "Template introuvable : " & tplPath
    End If

    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("AMDEC")

    ws.Range("AMD_Reference").Value = amdec.Reference
    ws.Range("AMD_Type").Value = amdec.TypeAmdec
    ws.Range("AMD_Processus").Value = amdec.Processus
    ws.Range("AMD_Produit").Value = amdec.Produit
    ws.Range("AMD_Responsable").Value = amdec.Responsable
    ws.Range("AMD_Equipe").Value = amdec.Equipe
    ws.Range("AMD_Version").Value = amdec.Version
    ws.Range("AMD_Statut").Value = amdec.Statut
    If IsDate(amdec.DateRevue) Then ws.Range("AMD_DateRevue").Value = FormatIsoDate(amdec.DateRevue)

    Dim s As TStatAmdec
    s = StatistiquesAmdec(amdec.Id)
    ws.Range("AMD_NbLignes").Value = s.NbLignes
    ws.Range("AMD_NbCritiques").Value = s.NbCritiques
    ws.Range("AMD_RpnMax").Value = s.RpnMax
    ws.Range("AMD_RpnMoyen").Value = s.RpnMoyen

    Dim lignes As Collection, i As Long, row As Long, l As CLigneAMDEC
    Set lignes = LignesAmdecTriees(amdec.Id)
    row = 23
    For i = 1 To lignes.Count
        Set l = lignes(i)
        ws.Cells(row, 1).Value = l.NumeroLigne
        ws.Cells(row, 2).Value = l.EtapeProcessus
        ws.Cells(row, 3).Value = l.Fonction
        ws.Cells(row, 4).Value = l.ModeDefaillance
        ws.Cells(row, 5).Value = l.Effet
        ws.Cells(row, 6).Value = l.Gravite
        ws.Cells(row, 7).Value = l.Cause
        ws.Cells(row, 8).Value = l.Occurrence
        ws.Cells(row, 9).Value = l.Detection
        ws.Cells(row, 10).Value = l.Rpn
        ws.Cells(row, 11).Value = l.Criticite
        ws.Cells(row, 12).Value = l.ActionRecommandee
        ws.Cells(row, 13).Value = l.ResponsableAction
        If IsDate(l.DateEcheanceAction) Then ws.Cells(row, 14).Value = FormatIsoDate(l.DateEcheanceAction)
        If l.RpnApres > 0 Then ws.Cells(row, 15).Value = l.RpnApres
        row = row + 1
    Next i

    Dim outPath As String
    outPath = AppDataPath() & "reports\AMDEC_" & amdec.Reference & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererRapportAmdec = outPath
End Function