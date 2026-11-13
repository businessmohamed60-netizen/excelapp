Public Sub Rb_Notifications(ByVal control As IRibbonControl): FNotifications.Show vbModeless: End Sub
Public Sub Rb_Parametres(ByVal control As IRibbonControl): FParametres.Show vbModal: End Sub
Public Sub Rb_Backup(ByVal control As IRibbonControl): FBackup.Show vbModal: End Sub
Public Sub Rb_Journal(ByVal control As IRibbonControl)
    ' Ouvre une feuille temporaire avec le journal
    Dim wb As Workbook, ws As Worksheet, all As Object, k As Variant, row As Long
    Set wb = Workbooks.Add: Set ws = wb.Worksheets(1): ws.Name = "Journal"
    ws.Range("A1:G1").Value = Array("Date", "Utilisateur", "Poste", "Action", "Collection", "Référence", "Champs")
    Set all = g_CollMgr.LoadCollection("audittrail")
    row = 2
    For Each k In all.Keys
        ws.Cells(row, 1).Value = Format$(all(k).DateHeure, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(row, 2).Value = all(k).Utilisateur
        ws.Cells(row, 3).Value = all(k).Poste
        ws.Cells(row, 4).Value = all(k).Action
        ws.Cells(row, 5).Value = all(k).Collection
        ws.Cells(row, 6).Value = all(k).ObjetLibelle
        ws.Cells(row, 7).Value = all(k).ChampsModifies
        row = row + 1
    Next k
    ws.Columns("A:G").AutoFit
End Sub