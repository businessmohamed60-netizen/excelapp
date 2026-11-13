Attribute VB_Name = "FKpi"
Option Explicit

Private m_Kpi As Collection

Private Sub UserForm_Initialize()
    Caption = "Tableau de bord qualité"
    cmdRefresh.Caption = "Actualiser"
    cmdGraphique.Caption = "Graphique"
    cmdExport.Caption = "Exporter"
    cmdClose.Caption = T("btn.close")
    With lstKpi
        .ColumnCount = 6
        .ColumnWidths = "200;70;60;70;40;60"
    End With
    InitAnnees
    RefreshKpi
End Sub

Private Sub InitAnnees()
    Dim y As Long, current As Long
    current = Year(Date)
    cboAnnee.Clear
    For y = current To current - 3 Step -1
        cboAnnee.AddItem CStr(y)
    Next y
    cboAnnee.ListIndex = 0
End Sub

Private Sub RefreshKpi()
    Set m_Kpi = CalculerTousKpi(CLng(cboAnnee.Text))
    lstKpi.Clear
    Dim i As Long, k As TKpi, okN As Long, alertN As Long, critN As Long
    For i = 1 To m_Kpi.Count
        k = m_Kpi(i)
        lstKpi.AddItem k.Nom
        lstKpi.List(lstKpi.ListCount - 1, 1) = Format$(k.Valeur, "0.0") & k.Unite
        lstKpi.List(lstKpi.ListCount - 1, 2) = Format$(k.Cible, "0.0") & k.Unite
        lstKpi.List(lstKpi.ListCount - 1, 3) = k.Statut
        lstKpi.List(lstKpi.ListCount - 1, 4) = k.Tendance
        lstKpi.List(lstKpi.ListCount - 1, 5) = Format$(k.DeltaMoisPrecedent, "+0.0;-0.0;0")
        Select Case k.Statut
            Case "OK": okN = okN + 1
            Case "Alerte": alertN = alertN + 1
            Case "Critique": critN = critN + 1
        End Select
    Next i
    lblSynthese.Caption = "Synthèse " & cboAnnee.Text & " : " & okN & " OK / " & _
        alertN & " alerte(s) / " & critN & " critique(s)"
End Sub

Private Sub cboAnnee_Change()
    If cboAnnee.ListIndex >= 0 Then RefreshKpi
End Sub

Private Sub cmdRefresh_Click()
    g_CollMgr.InvalidateAll
    RefreshKpi
End Sub

Private Sub cmdGraphique_Click()
    On Error GoTo ErrHandler
    ' Crée une feuille temporaire avec les KPI et un graphique
    Dim wb As Workbook, ws As Worksheet, i As Long, k As TKpi
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "KPI"
    ws.Cells(1, 1).Value = "KPI"
    ws.Cells(1, 2).Value = "Valeur"
    ws.Cells(1, 3).Value = "Cible"
    For i = 1 To m_Kpi.Count
        k = m_Kpi(i)
        ws.Cells(i + 1, 1).Value = k.Nom
        ws.Cells(i + 1, 2).Value = k.Valeur
        ws.Cells(i + 1, 3).Value = k.Cible
    Next i
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=350)
    ch.Chart.SetSourceData Source:=ws.Range("A1:C" & (m_Kpi.Count + 1))
    ch.Chart.ChartType = xlBarClustered
    ch.Chart.HasTitle = True
    ch.Chart.ChartTitle.Text = "KPI qualité " & cboAnnee.Text
    ws.Activate
    Exit Sub
ErrHandler:
    LogEx "FKpi", "cmdGraphique_Click", Err
    ShowError "Graphique", Err.Description
End Sub

Private Sub cmdExport_Click()
    On Error GoTo ErrHandler
    Dim path As String, wb As Workbook, ws As Worksheet, i As Long, k As TKpi
    path = AppDataPath() & "reports\KPI_" & cboAnnee.Text & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "KPI"
    ws.Range("A1:F1").Value = Array("KPI", "Valeur", "Cible", "Statut", "Tendance", "Delta")
    For i = 1 To m_Kpi.Count
        k = m_Kpi(i)
        ws.Cells(i + 1, 1).Value = k.Nom
        ws.Cells(i + 1, 2).Value = k.Valeur
        ws.Cells(i + 1, 3).Value = k.Cible
        ws.Cells(i + 1, 4).Value = k.Statut
        ws.Cells(i + 1, 5).Value = k.Tendance
        ws.Cells(i + 1, 6).Value = k.DeltaMoisPrecedent
    Next i
    Application.DisplayAlerts = False
    wb.SaveAs path, FileFormat:=xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    ShowInfo "Export", "Tableau de bord exporté :" & vbCrLf & path
    Exit Sub
ErrHandler:
    LogEx "FKpi", "cmdExport_Click", Err
    ShowError "Export", Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub