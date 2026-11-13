Public Sub Rb_Moyens(ByVal control As IRibbonControl): FMoyenMesure.Show vbModeless: End Sub
Public Sub Rb_Etalonnages(ByVal control As IRibbonControl)
    ' Ouvre FMoyenMesure filtré sur les étalonnages échus
    FMoyenMesure.Show vbModeless
End Sub
Public Sub Rb_Msa(ByVal control As IRibbonControl): FMsa.Show vbModeless: End Sub
Public Sub Rb_Capabilite(ByVal control As IRibbonControl): FCapabilite.Show vbModeless: End Sub
Public Sub Rb_Carte(ByVal control As IRibbonControl): FCarteControle.Show vbModeless: End Sub