IATFConsultantAssistant/
│
├── build/                                    ← Livrable compilé (.xlam)
│   └── IATFConsultantAssistant.xlam
│
├── src/                                      ← Sources VBA (à importer dans Excel)
│   │
│   ├── ThisWorkbook.cls
│   ├── M_Config.bas
│   ├── M_Global.bas                          ← g_Config, g_CollMgr, g_SessionHashes
│   ├── M_Utils.bas                           ← NewGuid, FormatIsoDate, EscapeXml, Log...
│   ├── M_Repository.bas                      ← helpers CRepository
│   ├── M_Storage.bas                         ← Phase 7 : DataPath, HashFichier
│   ├── M_Db.bas                              ← Phase 8 : SQLite
│   ├── M_Schema.bas                          ← Phase 8 : schéma SQL
│   ├── M_Migration.bas                       ← Phase 8 : XML → SQL
│   ├── M_Kpi.bas                             ← Phase 4 : KPI XML
│   ├── M_KpiSql.bas                          ← Phase 8 : KPI SQL
│   ├── M_Report.bas                          ← Phases 2,3,5,6,9 : rapports
│   ├── M_Amdec.bas                           ← Phase 5
│   ├── M_Stats.bas                           ← Phase 6
│   ├── M_Api.bas                             ← Phase 8
│   ├── M_PowerBi.bas                         ← Phase 8
│   ├── M_AuditTrail.bas                      ← Phase 7
│   ├── M_Notify.bas                          ← Phase 7
│   ├── M_Signature.bas                       ← Phase 7
│   ├── M_Ribbon.bas                          ← handlers ruban
│   ├── M_Bootstrap.bas                       ← Bootstrap, Shutdown
│   ├── RunAllTests.bas                       ← runner 33 suites
│   │
│   ├── CRepository.cls
│   ├── CCollectionManager.cls
│   │
│   ├── CClient.cls
│   ├── CAudit.cls
│   ├── CConstat.cls
│   ├── CNonConformite.cls
│   ├── CCapa.cls
│   ├── C8D.cls
│   ├── CAuditProgramme.cls
│   ├── CDocument.cls
│   ├── CExigence.cls
│   ├── CAMDEC.cls
│   ├── CLigneAMDEC.cls
│   ├── CPlanControle.cls
│   ├── CLignePlan.cls
│   ├── CCaracteristique.cls
│   ├── CMoyenMesure.cls
│   ├── CEtalonnage.cls
│   ├── CEtudeMsa.cls
│   ├── CLigneMsa.cls
│   ├── CEtudeCapabilite.cls
│   ├── CCarteControle.cls
│   ├── CLigneCarte.cls
│   ├── CAuditTrail.cls
│   ├── CFournisseur.cls
│   ├── CEvaluationFournisseur.cls
│   ├── CAuditFournisseur.cls
│   ├── CEquipement.cls
│   ├── CMaintenancePreventive.cls
│   ├── CInterventionMaintenance.cls
│   ├── CRejetInterne.cls
│   ├── CAnalyseRejet.cls
│   ├── CPlanSurveillance.cls
│   ├── CLigneSurveillance.cls
│   │
│   ├── FMain.frm / .frx
│   ├── FClient.frm / .frx
│   ├── FAudit.frm / .frx
│   ├── FConstat.frm / .frx
│   ├── FNonConformite.frm / .frx
│   ├── FCapa.frm / .frx
│   ├── F8D.frm / .frx
│   ├── FAuditProgramme.frm / .frx
│   ├── FDocument.frm / .frx
│   ├── FExigence.frm / .frx
│   ├── FKpi.frm / .frx
│   ├── FAmdec.frm / .frx
│   ├── FPlanControle.frm / .frx
│   ├── FCaracteristique.frm / .frx
│   ├── FMoyenMesure.frm / .frx
│   ├── FMsa.frm / .frx
│   ├── FCapabilite.frm / .frx
│   ├── FCarteControle.frm / .frx
│   ├── FParametres.frm / .frx
│   ├── FNotifications.frm / .frx
│   ├── FBackup.frm / .frx
│   ├── FMigration.frm / .frx
│   ├── FApi.frm / .frx
│   ├── FFournisseur.frm / .frx
│   ├── FEvalFournisseur.frm / .frx
│   ├── FAuditFournisseur.frm / .frx
│   ├── FEquipement.frm / .frx
│   ├── FMaintenance.frm / .frx
│   ├── FRejetInterne.frm / .frx
│   ├── FPlanSurveillance.frm / .frx
│   │
│   ├── T_Utils.bas
│   ├── T_Client.bas
│   ├── T_Audit.bas
│   ├── T_Constat.bas
│   ├── T_Repository.bas
│   ├── T_NonConformite.bas
│   ├── T_Capa.bas
│   ├── T_8D.bas
│   ├── T_AuditProgramme.bas
│   ├── T_Document.bas
│   ├── T_Exigence.bas
│   ├── T_Kpi.bas
│   ├── T_Amdec.bas
│   ├── T_PlanControle.bas
│   ├── T_Caracteristique.bas
│   ├── T_Stats.bas
│   ├── T_MoyenMesure.bas
│   ├── T_Msa.bas
│   ├── T_Capabilite.bas
│   ├── T_Carte.bas
│   ├── T_Storage.bas
│   ├── T_AuditTrail.bas
│   ├── T_Notify.bas
│   ├── T_Signature.bas
│   ├── T_Db.bas
│   ├── T_DbConcurrency.bas
│   ├── T_Migration.bas
│   ├── T_KpiSql.bas
│   ├── T_Api.bas
│   ├── T_Fournisseur.bas
│   ├── T_Equipement.bas
│   ├── T_Rejet.bas
│   └── T_Surveillance.bas
│
├── templates/                                ← Modèles .xlsx (plages nommées)
│   ├── Rapport_Audit_Processus.xlsx          ← Phase 2
│   ├── Rapport_8D.xlsx                       ← Phase 3
│   ├── Revue_Direction.xlsx                  ← Phase 4
│   ├── Matrice_Conformite.xlsx               ← Phase 4
│   ├── AMDEC_Processus.xlsx                  ← Phase 5
│   ├── Plan_Controle.xlsx                    ← Phase 5
│   ├── Fiche_Moyen_Mesure.xlsx               ← Phase 6
│   ├── Rapport_MSA.xlsx                      ← Phase 6
│   ├── Rapport_Capabilite.xlsx               ← Phase 6
│   ├── Carte_Controle.xlsx                   ← Phase 6
│   ├── Rapport_Fournisseur.xlsx              ← Phase 9
│   ├── Rapport_TPM.xlsx                      ← Phase 9
│   ├── Rapport_Rejets.xlsx                   ← Phase 9
│   └── Rapport_Surveillance.xlsx             ← Phase 9
│
├── ribbon/
│   └── Ribbon.xml                            ← onglets IATF, Pilotage, Risques, Métrologie, Admin, Données, Modules
│
├── scripts/
│   ├── install.ps1
│   ├── uninstall.ps1
│   ├── build.ps1
│   └── verifier.ps1                          ← NOUVEAU : vérification pré-installation
│
├── docs/
│   ├── README.md
│   ├── INSTALLATION.md
│   ├── CHANGELOG.md
│   └── Matrice_Tests.md
│
└── dist/                                     ← Généré par build.ps1
    ├── IATFConsultantAssistant.xlam
    ├── install.ps1
    ├── uninstall.ps1
    ├── verifier.ps1
    └── templates/