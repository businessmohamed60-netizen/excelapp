# Manuel d'Intégration Excel — IATF Consultant Assistant

## Vue d'ensemble

**IATF Consultant Assistant** est un add-in Excel (VBA) pour piloter la conformité à la norme **IATF 16949** (système de management de la qualité automobile). Il s'intègre directement dans Excel sous forme d'onglets de ruban supplémentaires et fournit :

- **Tableau de bord** avec KPIs de conformité
- **Gestion des clients et audits** (processus, fournisseurs)
- **Core Tools** : APQP, PPAP, AMDEC/FMEA, SPC, MSA
- **Métrologie** : moyens de mesure, étalonnages, capabilité, cartes de contrôle
- **Plans de contrôle** et caractéristiques CC/SC
- **Non-conformités** et actions correctives (8D, CAPA)
- **Gestion documentaire** avec matrice de conformité
- **Gestion des fournisseurs** et évaluations
- **Équipements et maintenance** (TPM)
- **Rejets internes** et plans de surveillance
- **Notifications** automatiques et **journal d'audit**
- **Base de données** SQLite (mode SQL) ou XML (mode fichier)
- **API REST** et intégration **Power BI**

---

## Architecture du projet

```
IATFConsultantAssistant/
│
├── src/
│   ├── classes/              26 classes metier (C*.cls)
│   ├── classes/fragments/    Extraits de code (gestion evenements)
│   ├── modules/              22 modules VBA (M_*.bas)
│   ├── modules/              62 handlers/Code (Code_*.cls)
│   ├── modules/fragments/    Fragments de code supplementaires
│   ├── forms/                10 formulaires UserForm (.frm)
│   ├── forms/fragments/      22 definitions de formulaires (reference)
│   └── tests/                28 suites de tests unitaires (T_*.bas)
│
├── ribbon/
│   ├── Ribbon.xml            Ruban unifie (7 onglets) ← A UTILISER
│   ├── Ruban_XML_*.xml       7 fragments individuels (archive)
│
├── scripts/
│   ├── build.ps1             Assemblage du .xlam
│   ├── install.ps1           Installation du add-in
│   ├── uninstall.ps1         Desinstallation
│   └── verifier.ps1          Verification pre-installation (SQLite)
│
├── docs/                     Documentation
├── build/                    Livrable .xlam (a generer)
└── templates/                Modeles .xlsx (a creer)
```

---

## Prérequis

### Logiciel requis

1. **Microsoft Excel 2016 ou plus récent** (Office 365 recommandé)
2. **Office RibbonX Editor** (gratuit) — pour injecter le ruban XML
   - Téléchargement : https://github.com/fernandreu/office-ribbonx-editor/releases

### Optionnel (recommandé)

3. **Driver SQLite3 ODBC** — pour le mode base de données SQL
   - Téléchargement : https://www.ch-werner.de/sqliteodbc/
   - Version 64-bit si Excel est 64-bit, 32-bit sinon
   - Sans ce driver, l'application fonctionne en mode XML (fichiers)

### Références VBA requises

Dans l'éditeur VBA (Alt+F11) > Outils > Références, cochez :

| Référence | Description |
|-----------|-------------|
| Microsoft Scripting Runtime | Dictionnaires, gestion de fichiers |
| Microsoft XML, v6.0 | Lecture/écriture XML (config, données) |
| Microsoft Forms 2.0 Object Library | UserForms (ajouté automatiquement) |

---

## Étape 1 : Assemblage du add-in

### 1.1 — Préparer le classeur

1. Ouvrez **Excel**
2. Créez un **nouveau classeur vide**
3. Appuyez sur **Alt+F11** pour ouvrir l'éditeur VBA

### 1.2 — Ajouter les références

4. Menu **Outils > Références**
5. Cochez :
   - `Microsoft Scripting Runtime`
   - `Microsoft XML, v6.0`
   - `Microsoft Forms 2.0 Object Library` (si absent, insérez d'abord un UserForm)

### 1.3 — Importer les modules (.bas)

6. Dans l'explorateur de projets (à gauche), clic droit sur `VBAProject`
7. **Importer un fichier** > sélectionnez tous les fichiers `.bas` du dossier `src/modules/`
8. Répété pour chaque fichier (ou importez-les un par un)

**22 modules à importer :**

| Module | Rôle |
|--------|------|
| M_Bootstrap | Démarrage/arrêt du add-in |
| M_Utils | Fonctions utilitaires (GUID, dates, log) |
| M_Storage | Gestion du stockage (chemin données) |
| M_Db | Base de données SQLite |
| M_Schema | Schéma SQL (22 tables) |
| M_Migration | Migration XML vers SQL |
| M_Kpi / M_KpiSql | Calcul des indicateurs |
| M_KpiFormatter | Formatage des KPIs |
| M_Report | Génération de rapports |
| M_Amdec | Logique AMDEC/FMEA |
| M_Stats | Statistiques et analyses |
| M_Api | API REST |
| M_PowerBi | Intégration Power BI |
| M_AuditTrail | Journal d'audit |
| M_Notify | Notifications |
| M_Signature | Signatures électroniques |
| M_I18n | Internationalisation (FR/EN) |
| M_Theme | Thème visuel |
| M_ChartExporter | Export de graphiques |
| M_Heatmap | Cartes de chaleur |
| M_DocRappels | Rappels documentaires |

### 1.4 — Importer les classes (.cls)

9. Clic droit sur `VBAProject` > **Importer un fichier**
10. Importez tous les `.cls` depuis `src/classes/` (26 classes métier)
11. Importez aussi les `.cls` depuis `src/modules/` (62 fichiers `Code_*.cls`)

**26 classes métier principales :**

| Classe | Objet IATF 16949 |
|--------|------------------|
| CAMDEC | Analyse AMDEC/FMEA |
| CLigneAMDEC | Ligne AMDEC (mode de défaillance) |
| CAuditProgramme | Programme annuel d'audits |
| CAuditFournisseur | Audit fournisseur |
| CDocument | Document (gestion documentaire) |
| CVersionDoc | Version de document |
| CDocManager | Gestionnaire documentaire |
| CExigence | Exigence normative |
| CFournisseur | Fournisseur |
| CEvaluationFournisseur | Évaluation fournisseur |
| CCaracteristique | Caractéristique CC/SC |
| CPlanSurveillance | Plan de surveillance |
| CLigneSurveillance | Ligne de surveillance |
| CLignePlan | Ligne de plan de contrôle |
| CMoyenMesure | Moyen de mesure |
| CEtalonnage | Étalonnage |
| CEtudeMsa | Étude MSA (R&R) |
| CLigneMsa | Ligne MSA (mesure) |
| CEtudeCapabilite | Étude de capabilité (Cp/Cpk) |
| CCarteControle | Carte de contrôle SPC |
| CEquipement | Équipement (TPM) |
| CMaintenancePreventive | Maintenance préventive |
| CInterventionMaintenance | Intervention de maintenance |
| CRejetInterne | Rejet interne |
| CAnalyseRejet | Analyse de rejet |
| CAuditTrail | Journal d'audit |

### 1.5 — Importer les formulaires (.frm)

12. Importez tous les `.frm` depuis `src/forms/`

**10 formulaires :**

| Formulaire | Fonction |
|------------|----------|
| FAmdec | Gestion AMDEC |
| FMsa | Étude MSA (R&R) |
| FKpi | Tableau de bord KPIs |
| FAuditProgramme | Programme d'audits |
| FDocument | Gestion documentaire |
| FNotifications | Centre de notifications |
| FParametres | Paramètres |
| FBackup | Sauvegarde |
| FMigration | Migration XML vers SQL |
| FApi | Configuration API REST |

**Note :** 22 définitions de formulaires supplémentaires sont disponibles dans `src/forms/fragments/` comme référence pour créer les UserForms manquants (FMain, FClient, FAudit, FNC, F8D, FCapa, etc.).

### 1.6 — Importer les tests (optionnel)

13. Importez les `.bas` depuis `src/tests/` (28 suites de tests)

### 1.7 — Compiler

14. Menu **Débogage > Compiler VBAProject**
15. La compilation doit passer **sans erreur**

### 1.8 — Tester

16. Ouvrez la fenêtre d'exécution (**Ctrl+G**)
17. Tapez `RunAllTests` et appuyez sur Entrée
18. Toutes les assertions doivent être **vertes**

---

## Étape 2 : Injection du ruban XML

### 2.1 — Enregistrer en .xlsm

1. Fermez l'éditeur VBA
2. Dans Excel : **Fichier > Enregistrer sous**
3. Type : **Classeur Excel prenant en charge les macros (*.xlsm)**
4. Nommez-le `IATFConsultantAssistant.xlsm`

### 2.2 — Injecter le ruban avec Office RibbonX Editor

5. Téléchargez et ouvrez **Office RibbonX Editor**
6. **File > Open** > sélectionnez le fichier `.xlsm`
7. Clic droit sur le fichier dans l'arbre > **Insert Office 2010+ Custom UI Part**
8. Un onglet `customUI.xml` apparaît
9. Ouvrez le fichier `ribbon/Ribbon.xml` de ce projet
10. Copiez tout son contenu et collez-le dans l'éditeur RibbonX
11. Clic droit > **Validate** — doit passer sans erreur
12. **File > Save** et fermez RibbonX Editor

### 2.3 — Vérifier le ruban

13. Ouvrez le `.xlsm` dans Excel
14. Vous devez voir **7 nouveaux onglets** dans le ruban :

| Onglet | Contenu |
|--------|---------|
| IATF Assistant | Tableau de bord, Clients, Audits, Core Tools, Outils |
| Pilotage | KPIs, Programme annuel, Documents, Revue de direction |
| Risques | AMDEC, Plans de contrôle, Caractéristiques CC/SC |
| Métrologie | Moyens de mesure, Étalonnages, MSA, Capabilité, SPC |
| Modules IATF | Fournisseurs, TPM, Rejets, Surveillance |
| Administration | Notifications, Paramètres, Sauvegarde, Journal |
| Données | Migration SQL, API REST, Power BI |

---

## Étape 3 : Sauvegarder en .xlam (Add-in)

1. Dans Excel avec le `.xlsm` ouvert
2. **Fichier > Enregistrer sous**
3. Type : **Macro Add-In Excel (*.xlam)**
4. Nom : `IATFConsultantAssistant.xlam`
5. Enregistrez dans le dossier `build/`

Le fichier `.xlam` est maintenant prêt à être déployé.

---

## Étape 4 : Installation du add-in

### Méthode automatique (PowerShell)

```powershell
# Dans PowerShell, depuis le dossier scripts/
.\install.ps1
```

### Méthode manuelle

1. Copiez `IATFConsultantAssistant.xlam` vers :
   `%APPDATA%\Microsoft\AddIns\`
2. Ouvrez Excel
3. **Fichier > Options > Compléments**
4. En bas, **Gérer : Compléments Excel** > cliquez **Atteindre**
5. Cliquez **Parcourir** > sélectionnez le `.xlam`
6. Cochez **IATFConsultantAssistant** dans la liste
7. Cliquez **OK**

L'onglet **IATF Assistant** apparaît dans le ruban Excel.

---

## Étape 5 : Configuration initiale

### 5.1 — Premier démarrage

Au premier lancement, l'application crée automatiquement :
- `%APPDATA%\IATFConsultantAssistant\` (dossier de données)
- `%APPDATA%\IATFConsultantAssistant\config.xml` (configuration)
- `%APPDATA%\IATFConsultantAssistant\templates\` (modèles)
- `%APPDATA%\IATFConsultantAssistant\backup\` (sauvegardes)

### 5.2 — Paramètres

Onglet **Administration > Paramètres** :
- **Langue** : Français ou Anglais
- **Mode de stockage** : XML (fichiers) ou SQL (SQLite)
- **Notifications automatiques** : activation et fréquence
- **Licence** : saisie de la clé de licence

### 5.3 — Driver SQLite (optionnel)

Pour activer le mode base de données SQL :
1. Installez le driver SQLite3 ODBC (https://www.ch-werner.de/sqliteodbc/)
2. Choisissez la version correspondant à votre Excel (32 ou 64-bit)
3. Dans Paramètres, passez le mode de stockage à **SQL**
4. Utilisez l'onglet **Données > Migration XML vers SQL** pour migrer les données existantes

---

## Structure de la base de données

L'application gère **22 tables** principales :

| Table | Description | Champ clé |
|-------|-------------|-----------|
| clients | Clients | Id, Code, RaisonSociale |
| audits | Audits processus | Id, ClientId, Norme, Statut |
| constats | Constats d'audit | Id, AuditId, Classification |
| nc | Non-conformités | Id, Numero, ClientId, Gravite |
| capa | Actions correctives | Id, NcId, TypeAction, Statut |
| huitd | Résolutions 8D | Id, NcId, D1-D8 |
| auditprogramme | Programme annuel | Id, Annee, ClientId |
| document | Documents | Id, Code, Titre, Version |
| exigence | Exigences normatives | Id, Reference, Conformite |
| amdec | Analyses AMDEC | Id, Reference, TypeAmdec |
| ligneamdec | Lignes AMDEC | Id, AmdecId, Rpn |
| plancontrole | Plans de contrôle | Id, Reference, ClientId |
| ligneplan | Lignes de plan | Id, PlanId, Caracteristique |
| caracteristique | Caractéristiques CC/SC | Id, Code, TypeCar |
| moyenmesure | Moyens de mesure | Id, Code, Designation |
| etalonnage | Étalonnages | Id, MoyenId, DateEtalonnage |
| etudemsa | Études MSA | Id, Reference, GrR |
| lignemsa | Mesures MSA | Id, EtudeId, Valeur |
| etudecapabilite | Études capabilité | Id, Reference, Cp, Cpk |
| cartecontrole | Cartes de contrôle | Id, Reference, TypeCarte |
| lignecarte | Données cartes | Id, CarteId, Valeurs |
| audittrail | Journal d'audit | Id, DateHeure, Action |

---

## Tests et validation

### Suite de tests (28 modules)

L'application inclut 28 modules de tests unitaires couvrant tous les domaines :

| Domaine | Modules de test |
|---------|----------------|
| Core Tools | T_Amdec, T_Capa, T_8D, T_Caracteristique, T_PlanControle |
| Métrologie | T_Msa, T_Capabilite, T_Carte, T_MoyenMesure |
| Fournisseurs | T_Fournisseur, T_Equipement, T_Rejet, T_Surveillance |
| Pilotage | T_Kpi, T_KpiSql, T_Stats, T_AuditProgramme |
| Infrastructure | T_Db, T_DbConcurrency, T_Storage, T_Migration |
| Sécurité | T_AuditTrail, T_Notify, T_Signature, T_Api |
| Documents | T_Document, T_Exigence, T_NC |

### Exécuter les tests

1. Ouvrez la fenêtre d'exécution (**Ctrl+G**)
2. Tapez `RunAllTests` et appuyez sur Entrée
3. Vérifiez que toutes les assertions passent (vertes)

---

## Dépannage

### Le ruban n'apparaît pas

- Vérifiez que le fichier `.xlam` est activé dans **Options > Compléments**
- Fermez et rouvrez Excel complètement
- Vérifiez que le XML a été validé dans RibbonX Editor

### Erreur de compilation

- Vérifiez que les **3 références** sont cochées (Scripting Runtime, MSXML v6.0, Forms 2.0)
- Vérifiez que tous les `.cls` et `.bas` ont été importés
- Menu **Débogage > Compiler VBAProject** pour voir l'erreur exacte

### Le mode SQL ne fonctionne pas

- Installez le driver **SQLite3 ODBC** (https://www.ch-werner.de/sqliteodbc/)
- Vérifiez la version 32/64-bit correspond à Excel
- L'application bascule automatiquement en mode XML en cas d'échec

### Les données sont stockées où ?

- **Windows** : `%APPDATA%\IATFConsultantAssistant\`
- **Mode XML** : fichiers `.xml` dans le dossier de données
- **Mode SQL** : base `iatf.db` (SQLite) dans le dossier de données
- **Sauvegardes** : `%APPDATA%\IATFConsultantAssistant\backup\`

---

## Désinstallation

### Automatique

```powershell
.\uninstall.ps1
```

### Manuelle

1. Excel > **Fichier > Options > Compléments** > désactivez le add-in
2. Supprimez `%APPDATA%\Microsoft\AddIns\IATFConsultantAssistant.xlam`
3. (Optionnel) Supprimez `%APPDATA%\IATFConsultantAssistant\` (données)

---

## Correspondance avec la norme IATF 16949

| Exigence IATF 16949 | Module de l'application |
|---------------------|------------------------|
| 4.4.1 Processus | AMDEC, Plans de contrôle |
| 5.1.1 Leadership | Revue de direction, Tableau de bord |
| 6.1 Actions face aux risques | AMDEC/FMEA |
| 6.2 Objectifs qualité | KPIs, Indicateurs |
| 7.1.3 Infrastructure | Équipements, Maintenance (TPM) |
| 7.1.5.1 Métrologie | Moyens de mesure, Étalonnages |
| 7.1.5.2 Étalonnage | Étalonnages, Certificats |
| 7.1.5.3 MSA | Études MSA (R&R) |
| 7.2 Compétences | Audits, Formations |
| 7.5 Informations documentées | Gestion documentaire |
| 8.2 Conception (APQP) | APQP, PPAP |
| 8.5.1 Contrôle production | Plans de contrôle, SPC |
| 8.5.1.1 Contrôle | Plans de surveillance |
| 8.6 Libération | Caractéristiques CC/SC |
| 8.7 Non-conformité | Non-conformités, 8D, CAPA |
| 9.2 Audit interne | Programme d'audits, Audits |
| 9.3 Revue de direction | Revue de direction |
| 10.2 Amélioration | CAPA, 8D |

---

## Support

Pour toute question sur l'intégration :
1. Consultez `docs/Arborescence Projet.md` pour la structure détaillée
2. Consultez `docs/Procedure Assemblage Final.md` pour la procédure résumée
3. Exécutez `scripts/build.ps1` pour un guide pas-à-pas interactif
