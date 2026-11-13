Public Function GenererMatriceConformite() As String
    Dim wb As Workbook, ws As Worksheet, row As Long
    Dim tplPath As String
    tplPath = TemplatePath() & "Matrice_Conformite.xlsx"
    If Dir(tplPath) = "" Then
        Err.Raise vbObjectError + 3004, "M_Report", "Template introuvable : " & tplPath
    End If

    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("Matrice")

    Dim exigences As Object, k As Variant
    Set exigences = g_CollMgr.LoadCollection("exigence")
    row = 2
    For Each k In exigences.Keys
        ws.Cells(row, 1).Value = exigences(k).Reference
        ws.Cells(row, 2).Value = exigences(k).Source
        ws.Cells(row, 3).Value = exigences(k).Intitule
        ws.Cells(row, 4).Value = IIf(exigences(k).Applicable, "Oui", "Non")
        ws.Cells(row, 5).Value = exigences(k).Conformite
        ws.Cells(row, 6).Value = exigences(k).Preuve
        row = row + 1
    Next k

    Dim outPath As String
    outPath = AppDataPath() & "reports\Matrice_Conformite_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    GenererMatriceConformite = outPath
End Function