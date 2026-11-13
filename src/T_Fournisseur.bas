Attribute vb_Name = "T_Fournisseur"
Option Explicit

Public Sub T_Fournisseur()
    ResetCounters

    Dim f As CFournisseur
    Set f = New CFournisseur
    f.InitNew
    AssertTrue Len(f.Id) > 0, "Id généré"
    AssertEquals "Actif", f.Statut, "statut par défaut"
    AssertEquals "Composant", f.Categorie, "catégorie par défaut"
    AssertFalse f.Critique, "non critique par défaut"

    ' Validation
    AssertFalse f.Validate(), "validation KO sans raison sociale"
    f.RaisonSociale = "Valeo Systèmes"
    AssertTrue f.Validate(), "validation OK"

    ' Code
    f.AssignCode
    AssertTrue Left$(f.Code, 4) = "FOU-", "préfixe FOU-"

    ' Note globale
    f.NoteQualite = 90: f.NoteLivraison = 80: f.NoteReactivite = 70: f.NotePrix = 60
    f.RecalculerNoteGlobale
    AssertEquals 81#, f.NoteGlobale, "note globale pondérée (0.4*90 + 0.3*80 + 0.2*70 + 0.1*60 = 81)"

    ' Certification
    f.EstCertifie = True: f.DateExpirationCertification = Date - 10
    AssertTrue f.EstCertificationExpiree(), "certification expirée"
    f.DateExpirationCertification = Date + 100
    AssertFalse f.EstCertificationExpiree(), "certification valide"

    ' Évaluation échue
    f.ProchaineEvaluation = Date - 5
    AssertTrue f.EstEvaluationEchue(), "évaluation échue"

    PrintSummary "CFournisseur"
End Sub