Option Explicit

Public Sub RunUAT()
    ResetCounters
    UAT01: UAT03: UAT07: UAT08: UAT09
    PrintSummary "UAT"
    If GetFailedCount() = 0 Then
        MsgBox GetPassedCount() & " assertions UAT toutes vertes.", vbInformation, "UAT OK"
    Else
        MsgBox GetFailedCount() & " assertion(s) UAT échouée(s).", vbExclamation, "UAT KO"
    End If
End Sub

Private Sub UAT01()
    Dim c As CClient: Set c = New CClient: c.InitNew
    c.Code = "UAT01": c.RaisonSociale = "Client UAT 01"
    AssertTrue c.Validate(), "UAT01 client valide"
    g_CollMgr.Save c, "clients"
    g_CollMgr.InvalidateCache "clients"
    Dim d As Object: Set d = g_CollMgr.LoadCollection("clients")
    AssertTrue d.Exists(c.Id), "UAT01 client persisté"
    g_CollMgr.DeleteById c.Id, "clients"
End Sub

Private Sub UAT03()
    Dim nc As CNonConformite: Set nc = New CNonConformite: nc.InitNew
    nc.ClientId = NewGuid(): nc.Description = "NC UAT03"
    nc.Classification = "Majeur": nc.AssignNumero
    AssertTrue Len(nc.Numero) > 0, "UAT03 NC numérotée"
    Dim capa As CCapa: Set capa = New CCapa: capa.InitNew
    capa.NcId = nc.Id: capa.Description = "CAPA UAT03": capa.AssignNumero
    AssertTrue Len(capa.Numero) > 0, "UAT03 CAPA numérotée"
End Sub

Private Sub UAT07()
    Dim l As CLigneAmdec: Set l = New CLigneAmdec: l.InitNew
    l.S = 9: l.O = 4: l.D = 3
    AssertEquals 108&, l.RPN(), "UAT07 RPN"
    AssertEquals "H", l.ActionPriority(), "UAT07 AP H"
End Sub

Private Sub UAT08()
    Dim e As CSpcEtude: Set e = New CSpcEtude: e.InitNew
    e.TailleSousGroupe = 5: e.LimiteSpecInf = 9.8: e.LimiteSpecSup = 10.2
    Dim i As Long
    For i = 1 To 20: e.Mesures.Add 10# + ((i Mod 5) - 2) * 0.02: Next i
    AssertBetween e.Cp(), 0.5, 5, "UAT08 Cp plausible"
End Sub

Private Sub UAT09()
    Dim m As CMsaEtude: Set m = New CMsaEtude: m.InitNew
    m.NbPieces = 3: m.NbOperateurs = 3: m.NbEssais = 3
    Dim op As Long, p As Long, es As Long
    For op = 1 To 3: For p = 1 To 3: For es = 1 To 3
        m.SetMesure op, p, es, 10 + p * 0.1 + op * 0.02
    Next es: Next p: Next op
    m.Tolerance = 1
    Dim v As String: v = m.Verdict()
    AssertTrue (v = "Acceptable" Or v = "Conditionnel" Or v = "Rejeté"), "UAT09 verdict MSA"
End Sub