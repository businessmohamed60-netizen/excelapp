Attribute VB_Name = "T_Amdec"
Option Explicit

Public Sub T_Amdec()
    ResetCounters

    ' En-tête
    Dim a As CAMDEC
    Set a = New CAMDEC
    a.InitNew
    AssertTrue Len(a.Id) > 0, "Id généré"
    AssertEquals "Brouillon", a.Statut, "statut par défaut"
    AssertEquals "Processus", a.TypeAmdec, "type par défaut"
    AssertEquals "1.0", a.Version, "version par défaut"
    AssertTrue a.DateRevue > Date, "revue future"

    AssertFalse a.Validate(), "validation KO sans processus/produit"
    a.Processus = "Usinage": a.Responsable = "J. Dupont"
    AssertTrue a.Validate(), "validation OK"

    ' Référence
    a.AssignReference
    AssertTrue Left$(a.Reference, 6) = "AMDEC-", "préfixe AMDEC-"

    ' Ligne AMDEC — RPN
    Dim l As CLigneAMDEC
    Set l = New CLigneAMDEC
    l.InitNew
    AssertEquals 125, l.Rpn, "RPN initial 5×5×5"
    l.Gravite = 9: l.Occurrence = 4: l.Detection = 3
    l.RecalculerRpn
    AssertEquals 108, l.Rpn, "RPN 9×4×3"
    AssertEquals "Critique", l.Criticite, "G=9 → Critique"

    l.Gravite = 3: l.Occurrence = 3: l.Detection = 3
    l.RecalculerRpn
    AssertEquals 27, l.Rpn, "RPN 3×3×3"
    AssertEquals "Faible", l.Criticite, "RPN 27 → Faible"

    l.Gravite = 5: l.Occurrence = 5: l.Detection = 5
    l.RecalculerRpn
    AssertEquals "Élevée", l.Criticite, "RPN 125 → Élevée"

    ' RPN après
    l.GraviteApres = 5: l.OccurrenceApres = 2: l.DetectionApres = 2
    l.RecalculerRpn
    AssertEquals 20, l.RpnApres, "RPN après 5×2×2"

    ' Validation
    AssertFalse l.Validate(), "validation KO sans AmdecId"
    l.AmdecId = NewGuid(): l.ModeDefaillance = "Dimension hors tolérance"
    AssertTrue l.Validate(), "validation OK"
    l.Gravite = 11
    AssertFalse l.Validate(), "validation KO si G > 10"

    ' EstEnRetard
    l.Gravite = 5
    l.DateEcheanceAction = Date - 3: l.StatutAction = "Ouverte"
    AssertTrue l.EstEnRetard(), "en retard si échéance passée"
    l.StatutAction = "Clôturée"
    AssertFalse l.EstEnRetard(), "pas en retard si clôturée"

    PrintSummary "CAMDEC / CLigneAMDEC"
End Sub