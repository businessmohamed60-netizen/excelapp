Public Sub AjouterBlocSignature(ByVal ws As Worksheet, ByVal startRow As Long, _
                                 ByVal signataire As String, ByVal role As String)
    ws.Cells(startRow, 1).Value = "Signé par :"
    ws.Cells(startRow, 2).Value = signataire
    ws.Cells(startRow + 1, 1).Value = "Rôle :"
    ws.Cells(startRow + 1, 2).Value = role
    ws.Cells(startRow + 2, 1).Value = "Date :"
    ws.Cells(startRow + 2, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(startRow + 3, 1).Value = "Hash :"
    ws.Cells(startRow + 3, 2).Value = "(calculé après sauvegarde)"
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 3, 2))
        .Font.Size = 9
        .Font.Italic = True
    End With
End Sub