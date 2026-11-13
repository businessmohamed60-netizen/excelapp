Attribute VB_Name = "M_Schema"
Option Explicit

' ============================================================
' M_Schema — Création du schéma relationnel
' ============================================================

Public Sub CreerSchema()
    On Error GoTo ErrHandler
    If Not DbOpen() Then Exit Sub

    Dim i As Long, tables As Variant, sql As String
    tables = Array( _
        "clients", "audits", "constats", "nc", "capa", "huitd", _
        "auditprogramme", "document", "exigence", _
        "amdec", "ligneamdec", "plancontrole", "ligneplan", "caracteristique", _
        "moyenmesure", "etalonnage", "etudemsa", "lignemsa", _
        "etudecapabilite", "cartecontrole", "lignecarte", _
        "audittrail")

    For i = LBound(tables) To UBound(tables)
        sql = SqlCreerTable(CStr(tables(i)))
        If Len(sql) > 0 Then DbExecute sql
    Next i

    ' Index
    DbExecute "CREATE INDEX IF NOT EXISTS idx_nc_client ON nc(ClientId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_nc_statut ON nc(Statut)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_capa_nc ON capa(NcId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_capa_statut ON capa(Statut)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_audit_client ON audits(ClientId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_audit_date ON audits(DateDebut)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_ligneamdec_amdec ON ligneamdec(AmdecId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_ligneplan_plan ON ligneplan(PlanId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_etal_moyen ON etalonnage(MoyenId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_lignemsa_etude ON lignemsa(EtudeId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_lignecarte_carte ON lignecarte(CarteId)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_trail_date ON audittrail(DateHeure)"
    DbExecute "CREATE INDEX IF NOT EXISTS idx_trail_objet ON audittrail(Collection, ObjetId)"

    Log LogInfo, "M_Schema", "Schéma créé (" & (UBound(tables) - LBound(tables) + 1) & " tables)"
    Exit Sub
ErrHandler:
    LogEx "M_Schema", "CreerSchema", Err
    ShowError "Schéma", Err.Description
End Sub

Private Function SqlCreerTable(ByVal table As String) As String
    ' Chaque table : Id TEXT PK + colonnes typées + DateCreation TEXT
    ' Les colonnes Date sont stockées en TEXT (ISO 8601) pour compatibilité cross-engine
    Select Case table
        Case "clients"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS clients (" & _
                "Id TEXT PRIMARY KEY, Code TEXT, RaisonSociale TEXT, " & _
                "Norme TEXT, Pays TEXT, Actif INTEGER, DateCreation TEXT)"
        Case "audits"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS audits (" & _
                "Id TEXT PRIMARY KEY, ClientId TEXT, Norme TEXT, Processus TEXT, " & _
                "TypeAudit TEXT, Auditeur TEXT, DateDebut TEXT, DateFin TEXT, " & _
                "Statut TEXT, Score REAL, Commentaire TEXT, DateCreation TEXT)"
        Case "constats"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS constats (" & _
                "Id TEXT PRIMARY KEY, AuditId TEXT, Chapitre TEXT, Classification TEXT, " & _
                "Description TEXT, Preuve TEXT, DateCreation TEXT)"
        Case "nc"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS nc (" & _
                "Id TEXT PRIMARY KEY, Numero TEXT UNIQUE, ClientId TEXT, AuditId TEXT, " & _
                "ConstatId TEXT, Source TEXT, Processus TEXT, Description TEXT, " & _
                "Classification TEXT, Gravite INTEGER, DateDetection TEXT, DetectePar TEXT, " & _
                "ContainmentImmediate TEXT, Statut TEXT, DateCloture TEXT, DateCreation TEXT)"
        Case "capa"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS capa (" & _
                "Id TEXT PRIMARY KEY, Numero TEXT UNIQUE, NcId TEXT, TypeAction TEXT, " & _
                "Description TEXT, Responsable TEXT, DateOuverture TEXT, DateEcheance TEXT, " & _
                "DateRealisation TEXT, Efficacite TEXT, Statut TEXT, DateCreation TEXT)"
        Case "huitd"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS huitd (" & _
                "Id TEXT PRIMARY KEY, Numero TEXT UNIQUE, NcId TEXT, Statut TEXT, " & _
                "D1_Champion TEXT, D1_Leader TEXT, D1_Membres TEXT, " & _
                "D2_Description TEXT, D2_Date TEXT, D2_5W2H TEXT, " & _
                "D3_Actions TEXT, D3_Date TEXT, D3_Verification TEXT, " & _
                "D4_Ishikawa TEXT, D4_5Pourquoi TEXT, D4_CauseRacine TEXT, " & _
                "D4_CauseEchappee TEXT, D4_Date TEXT, " & _
                "D5_Actions TEXT, D5_Responsable TEXT, D5_Date TEXT, " & _
                "D6_Actions TEXT, D6_Date TEXT, D6_Validation TEXT, " & _
                "D7_Actions TEXT, D7_Date TEXT, D8_Date TEXT, D8_Bilan TEXT, DateCreation TEXT)"
        Case "auditprogramme"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS auditprogramme (" & _
                "Id TEXT PRIMARY KEY, Annee INTEGER, ClientId TEXT, Processus TEXT, " & _
                "TypeAudit TEXT, Risque TEXT, FrequenceMois INTEGER, DatePrevue TEXT, " & _
                "DateRealisee TEXT, AuditId TEXT, Auditeur TEXT, Statut TEXT, " & _
                "Commentaire TEXT, DateCreation TEXT)"
        Case "document"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS document (" & _
                "Id TEXT PRIMARY KEY, Code TEXT, Titre TEXT, TypeDoc TEXT, Version TEXT, " & _
                "Statut TEXT, Redacteur TEXT, Approbateur TEXT, DateCreation TEXT, " & _
                "DateApprobation TEXT, DateRevuePrevue TEXT, DateRevueRealisee TEXT, " & _
                "Processus TEXT, CheminFichier TEXT, Commentaire TEXT)"
        Case "exigence"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS exigence (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT, Source TEXT, ClientId TEXT, " & _
                "Intitule TEXT, Chapitre TEXT, Applicable INTEGER, Conformite TEXT, " & _
                "Preuve TEXT, Responsable TEXT, DateEvaluation TEXT, DateRevue TEXT, " & _
                "Commentaire TEXT, DateCreation TEXT)"
        Case "amdec"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS amdec (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT UNIQUE, TypeAmdec TEXT, ClientId TEXT, " & _
                "Processus TEXT, Produit TEXT, Responsable TEXT, Equipe TEXT, " & _
                "DateCreation TEXT, DateRevue TEXT, Version TEXT, Statut TEXT, " & _
                "Approbateur TEXT, DateApprobation TEXT, Commentaire TEXT)"
        Case "ligneamdec"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS ligneamdec (" & _
                "Id TEXT PRIMARY KEY, AmdecId TEXT, NumeroLigne INTEGER, " & _
                "EtapeProcessus TEXT, Fonction TEXT, ModeDefaillance TEXT, Effet TEXT, " & _
                "Gravite INTEGER, Cause TEXT, Occurrence INTEGER, " & _
                "ControlesPreventifs TEXT, Detection INTEGER, ControlesDetectifs TEXT, " & _
                "Rpn INTEGER, Criticite TEXT, ActionRecommandee TEXT, " & _
                "ResponsableAction TEXT, DateEcheanceAction TEXT, " & _
                "DateRealisationAction TEXT, StatutAction TEXT, " & _
                "GraviteApres INTEGER, OccurrenceApres INTEGER, DetectionApres INTEGER, " & _
                "RpnApres INTEGER, DateCreation TEXT)"
        Case "plancontrole"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS plancontrole (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT UNIQUE, ClientId TEXT, " & _
                "Produit TEXT, Processus TEXT, Responsable TEXT, DateCreation TEXT, " & _
                "DateRevue TEXT, Version TEXT, Statut TEXT, Approbateur TEXT, " & _
                "DateApprobation TEXT, Commentaire TEXT)"
        Case "ligneplan"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS ligneplan (" & _
                "Id TEXT PRIMARY KEY, PlanId TEXT, NumeroLigne INTEGER, Etape TEXT, " & _
                "Caracteristique TEXT, CaracteristiqueId TEXT, TypeCaracteristique TEXT, " & _
                "Specification TEXT, MoyenControle TEXT, Frequence TEXT, " & _
                "TailleEchantillon INTEGER, CritereAcceptation TEXT, Responsable TEXT, " & _
                "Enregistrement TEXT, ReactionNonConformite TEXT, DateCreation TEXT)"
        Case "caracteristique"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS caracteristique (" & _
                "Id TEXT PRIMARY KEY, Code TEXT UNIQUE, Libelle TEXT, TypeCar TEXT, " & _
                "ClientId TEXT, Produit TEXT, Specification TEXT, Tolerance TEXT, " & _
                "SymboleClient TEXT, AmdecId TEXT, LigneAmdecId TEXT, " & _
                "PlanControleId TEXT, Statut TEXT, DateCreation TEXT, Commentaire TEXT)"
        Case "moyenmesure"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS moyenmesure (" & _
                "Id TEXT PRIMARY KEY, Code TEXT UNIQUE, Designation TEXT, TypeMoyen TEXT, " & _
                "Marque TEXT, Modele TEXT, NumeroSerie TEXT, Localisation TEXT, " & _
                "ClientId TEXT, EtendueMesure TEXT, Resolution TEXT, Incertitude TEXT, " & _
                "PeriodiciteMois INTEGER, DateMiseEnService TEXT, DernierEtalonnage TEXT, " & _
                "ProchainEtalonnage TEXT, Statut TEXT, Responsable TEXT, " & _
                "Commentaire TEXT, DateCreation TEXT)"
        Case "etalonnage"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS etalonnage (" & _
                "Id TEXT PRIMARY KEY, MoyenId TEXT, DateEtalonnage TEXT, " & _
                "TypeOperation TEXT, Organisme TEXT, CertificatNumero TEXT, " & _
                "Resultat TEXT, IncertitudeRelevee TEXT, ProchaineEcheance TEXT, " & _
                "Cout REAL, Commentaire TEXT, DateCreation TEXT)"
        Case "etudemsa"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS etudemsa (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT UNIQUE, MoyenId TEXT, " & _
                "Caracteristique TEXT, ClientId TEXT, NbOperateurs INTEGER, " & _
                "NbPieces INTEGER, NbEssais INTEGER, Tolerance REAL, USL REAL, LSL REAL, " & _
                "Moyenne REAL, EtendueMoyenne REAL, EcartTypeRepetabilite REAL, " & _
                "EcartTypeReproductibilite REAL, GrR REAL, PourcentGrR REAL, " & _
                "PourcentGrRTolerance REAL, Ndc INTEGER, Statut TEXT, DateEtude TEXT, " & _
                "Responsable TEXT, Commentaire TEXT, DateCreation TEXT)"
        Case "lignemsa"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS lignemsa (" & _
                "Id TEXT PRIMARY KEY, EtudeId TEXT, Operateur INTEGER, Piece INTEGER, " & _
                "Essai INTEGER, Valeur REAL, DateCreation TEXT)"
        Case "etudecapabilite"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS etudecapabilite (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT UNIQUE, Caracteristique TEXT, " & _
                "MoyenId TEXT, ClientId TEXT, USL REAL, LSL REAL, Cible REAL, " & _
                "NbValeurs INTEGER, Moyenne REAL, EcartTypeCourt REAL, " & _
                "EcartTypeLong REAL, Cp REAL, Cpk REAL, Pp REAL, Ppk REAL, " & _
                "Statut TEXT, DateEtude TEXT, Responsable TEXT, Commentaire TEXT, " & _
                "DateCreation TEXT)"
        Case "cartecontrole"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS cartecontrole (" & _
                "Id TEXT PRIMARY KEY, Reference TEXT UNIQUE, Caracteristique TEXT, " & _
                "MoyenId TEXT, TypeCarte TEXT, SousGroupeTaille INTEGER, " & _
                "USL REAL, LSL REAL, Cible REAL, LigneCentre REAL, " & _
                "LimiteSup REAL, LimiteInf REAL, LigneCentreR REAL, " & _
                "LimiteSupR REAL, LimiteInfR REAL, NbSousGroupes INTEGER, " & _
                "Statut TEXT, DateEtude TEXT, Responsable TEXT, " & _
                "Commentaire TEXT, DateCreation TEXT)"
        Case "lignecarte"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS lignecarte (" & _
                "Id TEXT PRIMARY KEY, CarteId TEXT, Numero INTEGER, " & _
                "Valeurs TEXT, DateCreation TEXT)"
        Case "audittrail"
            SqlCreerTable = "CREATE TABLE IF NOT EXISTS audittrail (" & _
                "Id TEXT PRIMARY KEY, DateHeure TEXT, Utilisateur TEXT, Poste TEXT, " & _
                "Action TEXT, Collection TEXT, ObjetId TEXT, ObjetLibelle TEXT, " & _
                "ChampsModifies TEXT, AncienneValeur TEXT, NouvelleValeur TEXT)"
        Case Else
            SqlCreerTable = ""
    End Select
End Function