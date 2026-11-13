Public Function GenererRapportPlanControle(ByVal plan As CPlanControle) As String
    Dim wb As Workbook, ws As Worksheet, tplPath As String
    tplPath = TemplatePath() & "Plan_Controle.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3006, "M_Report", "Template introuvable : " & tplPath
    End If

    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("Plan")

    ws.Range("PC_Reference").Value = plan.Reference
    ws.Range("PC_Produit").Value = plan.Produit
    ws.Range("PC_Processus").Value = plan.Processus
    ws.Range("PC_Responsable").Value = plan.Responsable
    ws.Range("PC_Version").Value = plan.Version
    ws.Range("PC_Statut").Value = plan.Statut

    Dim allLignes As Object, k As Variant, row As Long, l As CLignePlan
    Set allLignes = g_CollMgr.LoadCollection("ligneplan")
    row = 19
    For Each k In allLignes.Keys
        If allLignes(k).PlanId = plan.Id Then
            Set l = allLignes(k)
            ws.Cells(row, 1).Value = l.NumeroLigne
            ws.Cells(row, 2).Value = l.Etape
            ws.Cells(row, 3).Value = l.Caracteristique
            ws.Cells(row, 4).Value = l.TypeCaracteristique
            ws.Cells(row, 5).Value = l.Specification
            ws.Cells(row, 6).Value = l.MoyenControle
            ws.Cells(row, 7).Value = l.Frequence
            ws.Cells(row, 8).Value = l.TailleEchantillon
            ws.Cells(row, 9).Value = l.CritereAcceptation
            ws.Cells(row, 10).Value = l.Responsable
            ws.Cells(row, 11).Value = l.Enregistrement
            ws.Cells(row, 12).Value = l.ReactionNonConformite
            row = row + 1
        End If
    Next k

    Dim outPath As String
    outPath = AppDataPath() & "reports\PlanControle_" & plan.Reference & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererRapportPlanControle = outPath
End Function