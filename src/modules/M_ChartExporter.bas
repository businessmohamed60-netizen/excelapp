Option Explicit

' ============================================================
' M_ChartExporter — Génération et export PNG des graphiques
' ============================================================

Private Const CHART_SHEET As String = "_DashboardCharts"

' Point d'entrée : génère les 4 graphiques et retourne les chemins PNG
'   - Graph1 : Constats par classification (camembert)
'   - Graph2 : Audits par statut (barres)
'   - Graph3 : Activité 6 mois (courbe)
'   - Graph4 : Top 10 RPN AMDEC (barres horizontales)
Public Function GenererGraphiques() As Object
    Dim chemins As Object: Set chemins = CreateObject("Scripting.Dictionary")
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = PreparerFeuille()

    chemins("Graph1") = GenererGraph1(ws)
    chemins("Graph2") = GenererGraph2(ws)
    chemins("Graph3") = GenererGraph3(ws)
    chemins("Graph4") = GenererGraph4(ws)

    ' Masquer la feuille de travail
    ws.Visible = xlSheetVeryHidden

    Set GenererGraphiques = chemins
    Exit Function
ErrHandler:
    LogEx "M_ChartExporter", "GenererGraphiques", Err
    Set GenererGraphiques = chemins
End Function

Private Function PreparerFeuille() As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(CHART_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CHART_SHEET
    End If
    ws.Cells.Clear
    ws.Visible = xlSheetVisible
    Set PreparerFeuille = ws
End Function

Private Function GenererGraph1(ByVal ws As Worksheet) As String
    ' Constats par classification
    Dim constats As Object: Set constats = g_CollMgr.LoadCollection("constats")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    counts("Majeur") = 0: counts("Mineur") = 0: counts("Observation") = 0: counts("OFI") = 0
    Dim k As Variant
    For Each k In constats.Keys
        If counts.Exists(constats(k).Classification) Then counts(constats(k).Classification) = counts(constats(k).Classification) + 1
    Next k
    ws.Cells(1, 1).Value = "Classification": ws.Cells(1, 2).Value = "Nombre"
    Dim row As Long: row = 2
    Dim c As Variant
    For Each c In counts.Keys
        ws.Cells(row, 1).Value = c: ws.Cells(row, 2).Value = counts(c): row = row + 1
    Next c
    Dim co As ChartObject: Set co = ws.ChartObjects.Add(0, 0, 400, 250)
    With co.Chart
        .ChartType = xlPie
        .SetSourceData Source:=ws.Range("A1:B5")
        .HasTitle = True: .ChartTitle.Text = "Constats par classification"
        .ChartTitle.Font.Size = 12: .ChartTitle.Font.Name = PoliceTitre()
        .HasLegend = True: .Legend.Position = xlLegendPositionRight
    End With
    Dim pngPath As String: pngPath = Environ$("TEMP") & "\iatf_graph1.png"
    co.Chart.Export Filename:=pngPath, FilterName:="PNG"
    co.Delete
    GenererGraph1 = pngPath
End Function

Private Function GenererGraph2(ByVal ws As Worksheet) As String
    ' Audits par statut
    Dim audits As Object: Set audits = g_CollMgr.LoadCollection("audits")
    Dim counts As Object: Set counts = CreateObject("Scripting.Dictionary")
    counts("Planifié") = 0: counts("En cours") = 0: counts("Clôturé") = 0
    Dim k As Variant
    For Each k In audits.Keys
        If counts.Exists(audits(k).Statut) Then counts(audits(k).Statut) = counts(audits(k).Statut) + 1
    Next k
    ws.Cells(10, 1).Value = "Statut": ws.Cells(10, 2).Value = "Nombre"
    Dim row As Long: row = 11
    Dim c As Variant
    For Each c In counts.Keys
        ws.Cells(row, 1).Value = c: ws.Cells(row, 2).Value = counts(c): row = row + 1
    Next c
    Dim co As ChartObject: Set co = ws.ChartObjects.Add(0, 260, 400, 250)
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ws.Range("A10:B13")
        .HasTitle = True: .ChartTitle.Text = "Audits par statut"
        .ChartTitle.Font.Size = 12: .ChartTitle.Font.Name = PoliceTitre()
        .HasLegend = False
    End With
    Dim pngPath As String: pngPath = Environ$("TEMP") & "\iatf_graph2.png"
    co.Chart.Export Filename:=pngPath, FilterName:="PNG"
    co.Delete
    GenererGraph2 = pngPath
End Function

Private Function GenererGraph3(ByVal ws As Worksheet) As String
    ' Activité 6 mois (audits + NC)
    ws.Cells(20, 1).Value = "Mois": ws.Cells(20, 2).Value = "Audits": ws.Cells(20, 3).Value = "NC"
    Dim m As Long, d As Date
    Dim audits As Object, ncs As Object
    Set audits = g_CollMgr.LoadCollection("audits")
    Set ncs = g_CollMgr.LoadCollection("nc")
    Dim row As Long: row = 21
    For m = 5 To 0 Step -1
        d = DateSerial(Year(Date), Month(Date) - m, 1)
        Dim nA As Long, nN As Long, k As Variant
        For Each k In audits.Keys
            If Year(audits(k).DateCreation) = Year(d) And Month(audits(k).DateCreation) = Month(d) Then nA = nA + 1
        Next k
        For Each k In ncs.Keys
            If Year(ncs(k).DateCreation) = Year(d) And Month(ncs(k).DateCreation) = Month(d) Then nN = nN + 1
        Next k
        ws.Cells(row, 1).Value = Format$(d, "yyyy-mm")
        ws.Cells(row, 2).Value = nA: ws.Cells(row, 3).Value = nN
        row = row + 1
    Next m
    Dim co As ChartObject: Set co = ws.ChartObjects.Add(0, 520, 400, 250)
    With co.Chart
        .ChartType = xlLine
        .SetSourceData Source:=ws.Range("A20:C26")
        .HasTitle = True: .ChartTitle.Text = "Activité sur 6 mois"
        .ChartTitle.Font.Size = 12: .ChartTitle.Font.Name = PoliceTitre()
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
    Dim pngPath As String: pngPath = Environ$("TEMP") & "\iatf_graph3.png"
    co.Chart.Export Filename:=pngPath, FilterName:="PNG"
    co.Delete
    GenererGraph3 = pngPath
End Function

Private Function GenererGraph4(ByVal ws As Worksheet) As String
    ' Top 10 RPN AMDEC
    Dim amdecs As Object: Set amdecs = g_CollMgr.LoadCollection("amdec")
    Dim top As Object: Set top = CreateObject("Scripting.Dictionary")
    Dim k As Variant, l As CLigneAmdec
    For Each k In amdecs.Keys
        For Each l In amdecs(k).Lignes
            If top.Count < 10 Or l.RPN() > Application.Min(top.Items) Then
                Dim cle As String
                cle = amdecs(k).Reference & " | " & Left$(l.ModeDefaillance, 30)
                top(cle) = l.RPN()
            End If
        Next l
    Next k
    ' Tri décroissant
    Dim keys As Variant, i As Long, j As Long, tmp As String
    keys = top.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If top(keys(j)) > top(keys(i)) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i
    ws.Cells(30, 1).Value = "Mode de défaillance": ws.Cells(30, 2).Value = "RPN"
    Dim row As Long: row = 31
    For i = LBound(keys) To UBound(keys)
        ws.Cells(row, 1).Value = Left$(keys(i), 25): ws.Cells(row, 2).Value = top(keys(i))
        row = row + 1
        If row > 40 Then Exit For
    Next i
    Dim co As ChartObject: Set co = ws.ChartObjects.Add(0, 780, 400, 250)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=ws.Range("A30:B" & (row - 1))
        .HasTitle = True: .ChartTitle.Text = "Top 10 RPN AMDEC"
        .ChartTitle.Font.Size = 12: .ChartTitle.Font.Name = PoliceTitre()
        .HasLegend = False
    End With
    Dim pngPath As String: pngPath = Environ$("TEMP") & "\iatf_graph4.png"
    co.Chart.Export Filename:=pngPath, FilterName:="PNG"
    co.Delete
    GenererGraph4 = pngPath
End Function