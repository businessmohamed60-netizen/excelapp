Option Explicit

Public Sub RunCoreToolsTests()
    ResetCounters
    Dim a As CProjetAPQP: Set a = New CProjetAPQP: a.InitNew
    AssertEquals 0&, a.Avancement, "APQP avancement 0"
    Dim p As CPpapDossier: Set p = New CPpapDossier: p.InitNew
    AssertEquals 3&, p.NiveauPpap, "PPAP niveau 3"
    Dim l As CLigneAmdec: Set l = New CLigneAmdec: l.InitNew
    l.S = 8: l.O = 5: l.D = 4
    AssertEquals 160&, l.RPN(), "RPN = 160"
    AssertEquals "H", l.ActionPriority(), "AP = H"
    Dim e As CSpcEtude: Set e = New CSpcEtude: e.InitNew
    e.TailleSousGroupe = 5
    Dim i As Long
    For i = 1 To 20: e.Mesures.Add 10# + ((i Mod 5) - 2) * 0.02: Next i
    AssertEquals 4&, e.NbSousGroupes(), "4 sous-groupes"
    Dim m As CMsaEtude: Set m = New CMsaEtude: m.InitNew
    m.NbPieces = 3: m.NbOperateurs = 3: m.NbEssais = 3
    Dim op As Long, p2 As Long, es As Long
    For op = 1 To 3: For p2 = 1 To 3: For es = 1 To 3
        m.SetMesure op, p2, es, 10 + p2 * 0.5 + op * 0.1 + es * 0.02
    Next es: Next p2: Next op
    m.Tolerance = 5: m.ClientId = NewGuid(): m.Caracteristique = "Longueur"
    AssertTrue m.Validate(), "MSA valide"
    PrintSummary "Core Tools"
End Sub