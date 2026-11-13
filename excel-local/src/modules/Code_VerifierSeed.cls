Option Explicit

Public Sub VerifierSeed()
    ResetCounters
    g_CollMgr.InvalidateAll
    Dim clients As Object, audits As Object, constats As Object
    Dim ncs As Object, capas As Object, huitds As Object
    Dim amdecs As Object, spcs As Object, msas As Object
    Dim apqps As Object, ppaps As Object
    Set clients = g_CollMgr.LoadCollection("clients")
    Set audits = g_CollMgr.LoadCollection("audits")
    Set constats = g_CollMgr.LoadCollection("constats")
    Set ncs = g_CollMgr.LoadCollection("nc")
    Set capas = g_CollMgr.LoadCollection("capa")
    Set huitds = g_CollMgr.LoadCollection("huitd")
    Set amdecs = g_CollMgr.LoadCollection("amdec")
    Set spcs = g_CollMgr.LoadCollection("spc")
    Set msas = g_CollMgr.LoadCollection("msa")
    Set apqps = g_CollMgr.LoadCollection("apqp")
    Set ppaps = g_CollMgr.LoadCollection("ppap")
    AssertEquals 3&, clients.Count, "3 clients"
    AssertEquals 10&, audits.Count, "10 audits"
    AssertEquals 40&, constats.Count, "40 constats"
    AssertEquals 20&, ncs.Count, "20 NC"
    AssertEquals 15&, capas.Count, "15 CAPA"
    AssertEquals 5&, huitds.Count, "5 8D"
    AssertEquals 2&, amdecs.Count, "2 AMDEC"
    AssertEquals 2&, spcs.Count, "2 SPC"
    AssertEquals 2&, msas.Count, "2 MSA"
    AssertEquals 3&, apqps.Count, "3 APQP"
    AssertEquals 3&, ppaps.Count, "3 PPAP"
    Dim k As Variant, orphelins As Long
    For Each k In audits.Keys: If Not clients.Exists(audits(k).ClientId) Then orphelins = orphelins + 1: Next k
    AssertEquals 0&, orphelins, "aucun audit orphelin"
    PrintSummary "Seed"
End Sub