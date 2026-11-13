Public Function GenererRapportMsa(ByVal etude As CEtudeMsa) As String
    Dim wb As Workbook, ws As Worksheet, tplPath As String
    tplPath = TemplatePath() & "Rapport_MSA.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3008, "M_Report", "Template introuvable : " & tplPath
    End If
    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("MSA")
    ws.Range("MSA_Reference").Value = etude.Reference
    ws.Range("MSA_Caracteristique").Value = etude.Caracteristique
    ws.Range("MSA_NbOperateurs").Value = etude.NbOperateurs
    ws.Range("MSA_NbPieces").Value = etude.NbPieces
    ws.Range("MSA_NbEssais").Value = etude.NbEssais
    ws.Range("MSA_USL").Value = etude.USL
    ws.Range("MSA_LSL").Value = etude.LSL
    ws.Range("MSA_Tolerance").Value = etude.Tolerance
    ws.Range("MSA_EV").Value = etude.EcartTypeRepetabilite
    ws.Range("MSA_AV").Value = etude.EcartTypeReproductibilite
    ws.Range("MSA_GrR").Value = etude.GrR
    ws.Range("MSA_PctGrR").Value = etude.PourcentGrR
    ws.Range("MSA_PctGrRTol").Value = etude.PourcentGrRTolerance
    ws.Range("MSA_Ndc").Value = etude.Ndc
    ws.Range("MSA_Interpretation").Value = etude.Interpretation()

    Dim lignes As Object, k As Variant, row As Long
    Set lignes = g_CollMgr.LoadCollection("lignemsa")
    row = 29
    For Each k In lignes.Keys
        If lignes(k).EtudeId = etude.Id Then
            ws.Cells(row, 1).Value = lignes(k).Operateur
            ws.Cells(row, 2).Value = lignes(k).Piece
            ws.Cells(row, 3).Value = lignes(k).Essai
            ws.Cells(row, 4).Value = lignes(k).Valeur
            row = row + 1
        End If
    Next k

    Dim outPath As String
    outPath = AppDataPath() & "reports\MSA_" & etude.Reference & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererRapportMsa = outPath
End Function