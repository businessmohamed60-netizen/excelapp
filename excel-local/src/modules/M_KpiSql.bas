Attribute VB_Name = "M_KpiSql"
Option Explicit

Public Function CalculerTousKpiSql(ByVal annee As Long) As Collection
    Dim result As New Collection
    If Not DbOpen() Then
        Set CalculerTousKpiSql = CalculerTousKpi(annee)
        Exit Function
    End If

    Dim rs As Object

    ' 1. Taux de clôture audits
    Set rs = DbQuery("SELECT COUNT(*) AS total, " & _
        "SUM(CASE WHEN Statut='Clôturé' THEN 1 ELSE 0 END) AS clos " & _
        "FROM audits WHERE substr(DateDebut, 1, 4) = ?", Array(CStr(annee)))
    Dim total As Long, clos As Long
    If Not rs.EOF Then total = Nz(rs("total")): clos = Nz(rs("clos"))
    rs.Close
    result.Add EvaluerKpi("Taux de clôture audits", _
        IIf(total > 0, clos / total * 100, 0), 90, "Haut", "%")

    ' 2. NC ouvertes
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM nc WHERE Statut <> 'Clôturée'")
    result.Add EvaluerKpi("NC ouvertes", Nz(rs("n")), 5, "Bas", "")
    rs.Close

    ' 3. Taux de clôture NC
    Set rs = DbQuery("SELECT COUNT(*) AS total, " & _
        "SUM(CASE WHEN Statut='Clôturée' THEN 1 ELSE 0 END) AS clos FROM nc")
    If Not rs.EOF Then total = Nz(rs("total")): clos = Nz(rs("clos"))
    rs.Close
    result.Add EvaluerKpi("Taux de clôture NC", IIf(total > 0, clos / total * 100, 0), 85, "Haut", "%")

    ' 4. CAPA en retard
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM capa " & _
        "WHERE Statut <> 'Clôturée' AND DateEcheance < ?", Array(Format$(Date, "yyyy-mm-dd")))
    result.Add EvaluerKpi("CAPA en retard", Nz(rs("n")), 2, "Bas", "")
    rs.Close

    ' 5. 8D ouverts
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM huitd WHERE D8_Date IS NULL OR D8_Date = ''")
    result.Add EvaluerKpi("8D ouverts", Nz(rs("n")), 3, "Bas", "")
    rs.Close

    ' 6. Étalonnages échus
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM moyenmesure " & _
        "WHERE Statut='En service' AND ProchainEtalonnage < ?", Array(Format$(Date, "yyyy-mm-dd")))
    result.Add EvaluerKpi("Étalonnages échus", Nz(rs("n")), 0, "Bas", "")
    rs.Close

    ' 7. Actions AMDEC en retard
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM ligneamdec " & _
        "WHERE StatutAction <> 'Clôturée' AND DateEcheanceAction < ?", Array(Format$(Date, "yyyy-mm-dd")))
    result.Add EvaluerKpi("Actions AMDEC en retard", Nz(rs("n")), 0, "Bas", "")
    rs.Close

    ' 8. Lignes AMDEC critiques
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM ligneamdec " & _
        "WHERE Criticite='Critique' AND StatutAction <> 'Clôturée'")
    result.Add EvaluerKpi("Lignes AMDEC critiques ouvertes", Nz(rs("n")), 0, "Bas", "")
    rs.Close

    ' 9. Revues documentaires échues
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM document " & _
        "WHERE Statut='Approuvé' AND DateRevuePrevue < ?", Array(Format$(Date, "yyyy-mm-dd")))
    result.Add EvaluerKpi("Revues documentaires échues", Nz(rs("n")), 0, "Bas", "")
    rs.Close

    ' 10. Exigences non conformes
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM exigence " & _
        "WHERE Applicable=1 AND Conformite IN ('Non conforme','Partielle')")
    result.Add EvaluerKpi("Exigences non conformes", Nz(rs("n")), 0, "Bas", "")
    rs.Close

    ' 11. MSA validées
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM etudemsa WHERE Statut='Validée'")
    result.Add EvaluerKpi("Études MSA validées", Nz(rs("n")), 1, "Haut", "")
    rs.Close

    ' 12. Capabilités validées
    Set rs = DbQuery("SELECT COUNT(*) AS n FROM etudecapabilite WHERE Statut='Validée'")
    result.Add EvaluerKpi("Études capabilité validées", Nz(rs("n")), 1, "Haut", "")
    rs.Close

    Set CalculerTousKpiSql = result
End Function

Private Function Nz(ByVal v As Variant) As Long
    If IsNull(v) Or IsEmpty(v) Then Nz = 0 Else Nz = CLng(v)
End Function