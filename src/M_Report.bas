Option Explicit

' ============================================================
' M_Report — Générateurs de rapports .xlsx
' ============================================================

Public Function GenererRapportAudit(ByVal audit As CAudit, ByVal client As CClient, _
                                    ByRef constats As Object) As String
    Dim wb As Workbook, ws As Worksheet, tplPath As String
    tplPath = TemplatePath() & "Rapport_Audit_Processus.xlsx"
    If Dir(tplPath) = "" Then Err.Raise vbObjectError + 3001, "M_Report", "Template introuvable : " & tplPath
    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("Rapport")
    ws.Range("R_ClientNom").Value = client.RaisonSociale
    ws.Range("R_ClientCode").Value = client.Code
    ws.Range("R_Norme").Value = audit.Norme
    ws.Range("R_Processus").Value = audit.Processus
    ws.Range("R_TypeAudit").Value = audit.TypeAudit
    ws.Range("R_Auditeur").Value = audit.Auditeur
    ws.Range("R_DateDebut").Value = FormatIsoDate(audit.DateDebut)
    ws.Range("R_DateFin").Value = FormatIsoDate(audit.DateFin)
    ws.Range("R_Statut").Value = audit.Statut
    ws.Range("R_Score").Value = audit.Score
    Dim nMajeurs As Long, nMineurs As Long, nObs As Long, nOfi As Long
    Dim k As Variant
    For Each k In constats.Keys
        Select Case constats(k).Classification
            Case "Majeur": nMajeurs = nMajeurs + 1
            Case "Mineur": nMineurs = nMineurs + 1
            Case "Observation": nObs = nObs + 1
            Case "OFI": nOfi = nOfi + 1
        End Select
    Next k
    ws.Range("R_NbMajeurs").Value = nMajeurs
    ws.Range("R_NbMineurs").Value = nMineurs
    ws.Range("R_NbObs").Value = nObs
    ws.Range("R_NbOFI").Value = nOfi
    Dim row As Long: row = 21
    For Each k In constats.Keys
        ws.Cells(row, 1).Value = constats(k).Chapitre
        ws.Cells(row, 2).Value = constats(k).Classification
        ws.Cells(row, 3).Value = constats(k).Description
        ws.Cells(row, 4).Value = constats(k).Preuve
        ws.Cells(row, 5).Value = FormatIsoDate(constats(k).DateCreation)
        row = row + 1
    Next k
    Dim outPath As String
    outPath = ReportsPath() & "Rapport_Audit_" & client.Code & "_" & Format$(audit.DateDebut, "yyyymmdd") & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Log LogInfo, "M_Report", "Rapport généré : " & outPath
    GenererRapportAudit = outPath
End Function

Public Function GenererRapport8D(ByVal d As C8D) As String
    Dim tplPath As String, outPath As String
    tplPath = TemplatePath() & "Rapport_8D.xlsx"
    If Dir(tplPath) = "" Then Err.Raise vbObjectError + 3002, "M_Report", "Template 8D introuvable"
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("8D")
    ws.Range("D8_Numero").Value = d.Numero
    ws.Range("D8_NcNumero").Value = d.NcId
    ws.Range("D8_Statut").Value = d.Statut
    ws.Range("D1_Champion").Value = d.D1_Champion
    ws.Range("D1_Leader").Value = d.D1_Leader
    ws.Range("D1_Membres").Value = d.D1_Membres
    ws.Range("D2_Description").Value = d.D2_Description
    If IsDate(d.D2_Date) Then ws.Range("D2_Date").Value = FormatIsoDate(d.D2_Date)
    ws.Range("D2_5W2H").Value = d.D2_5W2H
    ws.Range("D3_Actions").Value = d.D3_Actions
    If IsDate(d.D3_Date) Then ws.Range("D3_Date").Value = FormatIsoDate(d.D3_Date)
    ws.Range("D3_Verification").Value = d.D3_Verification
    ws.Range("D4_Ishikawa").Value = d.D4_Ishikawa
    ws.Range("D4_5Pourquoi").Value = d.D4_5Pourquoi
    ws.Range("D4_CauseRacine").Value = d.D4_CauseRacine
    ws.Range("D4_CauseEchappee").Value = d.D4_CauseEchappee
    If IsDate(d.D4_Date) Then ws.Range("D4_Date").Value = FormatIsoDate(d.D4_Date)
    ws.Range("D5_Actions").Value = d.D5_Actions
    ws.Range("D5_Responsable").Value = d.D5_Responsable
    If IsDate(d.D5_Date) Then ws.Range("D5_Date").Value = FormatIsoDate(d.D5_Date)
    ws.Range("D6_Actions").Value = d.D6_Actions
    If IsDate(d.D6_Date) Then ws.Range("D6_Date").Value = FormatIsoDate(d.D6_Date)
    ws.Range("D6_Validation").Value = d.D6_Validation
    ws.Range("D7_Actions").Value = d.D7_Actions
    If IsDate(d.D7_Date) Then ws.Range("D7_Date").Value = FormatIsoDate(d.D7_Date)
    If IsDate(d.D8_Date) Then ws.Range("D8_Date").Value = FormatIsoDate(d.D8_Date)
    ws.Range("D8_Bilan").Value = d.D8_Bilan
    outPath = ReportsPath() & "Rapport_8D_" & d.Numero & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenererRapport8D = outPath
End Function

Public Function GenererPSW(ByVal ppap As CPpapDossier, ByVal client As CClient) As String
    Dim tplPath As String, outPath As String
    tplPath = TemplatePath() & "PSW.xlsx"
    If Dir(tplPath) = "" Then Err.Raise vbObjectError + 3003, "M_Report", "Template PSW introuvable"
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(tplPath, ReadOnly:=False)
    Set ws = wb.Worksheets("PSW")
    ws.Range("PSW_Reference").Value = ppap.Reference
    ws.Range("PSW_Client").Value = client.RaisonSociale
    ws.Range("PSW_ClientCode").Value = client.Code
    ws.Range("PSW_Piece").Value = ppap.Piece
    ws.Range("PSW_RefPiece").Value = ppap.ReferencePiece
    ws.Range("PSW_NumSoumission").Value = ppap.NumeroSoumission
    ws.Range("PSW_Niveau").Value = ppap.NiveauPpap
    ws.Range("PSW_Statut").Value = ppap.Statut
    If IsDate(ppap.DateSoumission) And ppap.DateSoumission > 0 Then ws.Range("PSW_DateSoumission").Value = FormatIsoDate(ppap.DateSoumission)
    Dim i As Long
    For i = 1 To 18: ws.Range("PSW_El" & i).Value = IIf(ppap.Element(i), "X", ""): Next i
    outPath = ReportsPath() & "PSW_" & ppap.Reference & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenererPSW = outPath
End Function

Public Function ExporterAmdecAIAG(ByVal a As CAmdec) As String
    Dim wb As Workbook, ws As Worksheet, row As Long
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1): ws.Name = "AMDEC"
    ws.Cells(1, 1).Value = "Référence AMDEC": ws.Cells(1, 2).Value = a.Reference
    ws.Cells(2, 1).Value = "Pièce": ws.Cells(2, 2).Value = a.Piece
    ws.Cells(3, 1).Value = "Type": ws.Cells(3, 2).Value = a.TypeAmdec
    ws.Cells(4, 1).Value = "Processus": ws.Cells(4, 2).Value = a.Processus
    ws.Cells(5, 1).Value = "Équipe": ws.Cells(5, 2).Value = a.Equipe
    row = 7
    Dim headers As Variant, i As Long
    headers = Array("Élément", "Fonction", "Mode de défaillance", "Effets", "Causes", _
                    "Contrôles prévention", "Contrôles détection", "S", "O", "D", "RPN", "AP", _
                    "Actions recommandées", "Responsable", "Échéance", "S'", "O'", "D'", "RPN'", "AP'")
    For i = LBound(headers) To UBound(headers): ws.Cells(row, i + 1).Value = headers(i): Next i
    ws.Rows(row).Font.Bold = True
    ws.Rows(row).Interior.Color = RGB(200, 200, 200)
    Dim l As CLigneAmdec
    row = row + 1
    For Each l In a.Lignes
        ws.Cells(row, 1).Value = l.Element: ws.Cells(row, 2).Value = l.Fonction
        ws.Cells(row, 3).Value = l.ModeDefaillance: ws.Cells(row, 4).Value = l.Effets
        ws.Cells(row, 5).Value = l.Causes: ws.Cells(row, 6).Value = l.ControlesPrevention
        ws.Cells(row, 7).Value = l.ControlesDetection
        ws.Cells(row, 8).Value = l.S: ws.Cells(row, 9).Value = l.O: ws.Cells(row, 10).Value = l.D
        ws.Cells(row, 11).Value = l.RPN(): ws.Cells(row, 12).Value = l.ActionPriority()
        ws.Cells(row, 13).Value = l.ActionsRecommandees: ws.Cells(row, 14).Value = l.Responsable
        If l.DateEcheance > 0 Then ws.Cells(row, 15).Value = FormatIsoDate(l.DateEcheance)
        If l.NouveauS > 0 Then
            ws.Cells(row, 16).Value = l.NouveauS: ws.Cells(row, 17).Value = l.NouveauO
            ws.Cells(row, 18).Value = l.NouveauD: ws.Cells(row, 19).Value = l.NouveauRPN()
            ws.Cells(row, 20).Value = l.NouvelleAP()
        End If
        row = row + 1
    Next l
    ws.Columns.AutoFit
    Dim outPath As String
    outPath = ReportsPath() & "AMDEC_" & a.Reference & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ExporterAmdecAIAG = outPath
End Function

Public Function ExporterCarteControle(ByVal e As CSpcEtude) As String
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1): ws.Name = "Données"
    Set sh = wb.Worksheets.Add(After:=ws): sh.Name = "Carte"
    ws.Cells(1, 1).Value = "Référence": ws.Cells(1, 2).Value = e.Reference
    ws.Cells(2, 1).Value = "Caractéristique": ws.Cells(2, 2).Value = e.Caracteristique
    ws.Cells(3, 1).Value = "LSI": ws.Cells(3, 2).Value = e.LimiteSpecInf
    ws.Cells(4, 1).Value = "LSS": ws.Cells(4, 2).Value = e.LimiteSpecSup
    ws.Cells(5, 1).Value = "n": ws.Cells(5, 2).Value = e.TailleSousGroupe
    ws.Cells(7, 1).Value = "Sous-groupe": ws.Cells(7, 2).Value = "Moyenne"
    ws.Cells(7, 3).Value = "Étendue": ws.Cells(7, 4).Value = "LCS X"
    ws.Cells(7, 5).Value = "LCI X": ws.Cells(7, 6).Value = "LCS R": ws.Cells(7, 7).Value = "LCI R"
    Dim moyennes As Variant, etendues As Variant
    Dim mg As Double, rb As Double, n As Long, i As Long
    moyennes = e.MoyennesSousGroupes(): etendues = e.EtenduesSousGroupes()
    mg = e.MoyenneGenerale(): rb = e.EtendueMoyenne(): n = e.TailleSousGroupe
    Dim lcsX As Double, lciX As Double, lcsR As Double, lciR As Double
    lcsX = mg + A2(n) * rb: lciX = mg - A2(n) * rb
    lcsR = D4(n) * rb: lciR = D3(n) * rb
    Dim row As Long: row = 8
    For i = 0 To UBound(moyennes)
        ws.Cells(row, 1).Value = i + 1
        ws.Cells(row, 2).Value = moyennes(i)
        ws.Cells(row, 3).Value = etendues(i)
        ws.Cells(row, 4).Value = lcsX: ws.Cells(row, 5).Value = lciX
        ws.Cells(row, 6).Value = lcsR: ws.Cells(row, 7).Value = lciR
        row = row + 1
    Next i
    sh.Cells(1, 1).Value = "Étude : " & e.Reference & " — " & e.Caracteristique
    sh.Cells(2, 1).Value = "Cp = " & Format$(e.Cp(), "0.000")
    sh.Cells(3, 1).Value = "Cpk = " & Format$(e.Cpk(), "0.000")
    sh.Cells(4, 1).Value = "Pp = " & Format$(e.Pp(), "0.000")
    sh.Cells(5, 1).Value = "Ppk = " & Format$(e.Ppk(), "0.000")
    Dim chartObj As ChartObject
    Set chartObj = sh.ChartObjects.Add(Left:=20, Top:=100, Width:=600, Height:=300)
    With chartObj.Chart
        .ChartType = xlXYScatterLines
        .SetSourceData Source:=ws.Range(ws.Cells(7, 1), ws.Cells(7 + UBound(moyennes) + 1, 5))
        .HasTitle = True: .ChartTitle.Text = "Carte X"
    End With
    Dim outPath As String
    outPath = ReportsPath() & "SPC_" & e.Reference & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ExporterCarteControle = outPath
End Function

Public Function GenererRapportGRR(ByVal e As CMsaEtude) As String
    Dim wb As Workbook, ws As Worksheet, row As Long
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1): ws.Name = "GRR"
    ws.Cells(1, 1).Value = "Rapport GRR — " & e.Reference
    ws.Cells(2, 1).Value = "Caractéristique": ws.Cells(2, 2).Value = e.Caracteristique
    ws.Cells(3, 1).Value = "Moyen de mesure": ws.Cells(3, 2).Value = e.MoyenMesure
    ws.Cells(4, 1).Value = "Tolérance": ws.Cells(4, 2).Value = e.Tolerance
    row = 6
    ws.Cells(row, 1).Value = "Indicateur": ws.Cells(row, 2).Value = "Valeur"
    ws.Cells(row + 1, 1).Value = "%EV": ws.Cells(row + 1, 2).Value = e.PourcentEV()
    ws.Cells(row + 2, 1).Value = "%AV": ws.Cells(row + 2, 2).Value = e.PourcentAV()
    ws.Cells(row + 3, 1).Value = "%GRR": ws.Cells(row + 3, 2).Value = e.PourcentGRR()
    ws.Cells(row + 4, 1).Value = "%PV": ws.Cells(row + 4, 2).Value = e.PourcentPV()
    ws.Cells(row + 5, 1).Value = "ndc": ws.Cells(row + 5, 2).Value = e.NombreDistinctCategorie()
    ws.Cells(row + 6, 1).Value = "Verdict": ws.Cells(row + 6, 2).Value = e.Verdict()
    row = 17
    ws.Cells(row, 1).Value = "Opérateur": ws.Cells(row, 2).Value = "Pièce"
    Dim ess As Long
    For ess = 1 To e.NbEssais: ws.Cells(row, 2 + ess).Value = "Essai " & ess: Next ess
    Dim op As Long, p As Long
    row = row + 1
    For op = 1 To e.NbOperateurs
        For p = 1 To e.NbPieces
            ws.Cells(row, 1).Value = "Op " & op: ws.Cells(row, 2).Value = "P" & p
            For ess = 1 To e.NbEssais: ws.Cells(row, 2 + ess).Value = e.GetMesure(op, p, ess): Next ess
            row = row + 1
        Next p
    Next op
    ws.Columns.AutoFit
    Dim outPath As String
    outPath = ReportsPath() & "MSA_GRR_" & e.Reference & "_" & Format$(Now, "hhnnss") & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    GenererRapportGRR = outPath
End Function