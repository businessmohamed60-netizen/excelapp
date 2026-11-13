# IATF Consultant Assistant

Add-in Excel (VBA) pour piloter la conformite a la norme **IATF 16949** avec un systeme de licences web.

## Structure du projet

Le projet est separe en **deux dossiers independants** :

```
IATFConsultantAssistant/
├── excel-local/          A installer sur chaque poste Excel (add-in VBA)
└── serveur-web/          A installer sur votre serveur (gestion des licences)
```

---

## 1. Dossier `excel-local/` — Add-in Excel

Tout ce qui doit etre installe sur le poste de travail de chaque utilisateur.

- **26 classes metier** : AMDEC, Fournisseurs, Metrologie, Audits, Documents...
- **22 modules VBA** + 62 handlers + module de licence web
- **10 formulaires** UserForm
- **28 suites de tests** unitaires
- **Ruban Excel** unifie (7 onglets)
- **Scripts PowerShell** d'installation

### Installation rapide
1. Lisez `excel-local/docs/MANUEL_INTEGRATION_EXCEL.md`
2. Executez `excel-local/scripts/build.ps1` dans PowerShell
3. Activez le add-in dans Excel > Options > Complements

Voir `excel-local/README.md` pour les details.

---

## 2. Dossier `serveur-web/` — Gestionnaire de licences

Application web a installer sur votre serveur pour gerer les licences des utilisateurs.

- **Interface web** d'administration (React)
- **API REST** pour la verification des licences depuis Excel
- **Base de donnees** Supabase (licences + activations)
- **Statistiques** en temps reel

### Installation rapide
1. `cd serveur-web && npm install`
2. Configurez le fichier `.env` (deja pre-rempli)
3. `npm run dev` pour lancer l'interface web
4. `npm run build` pour la production

Voir `serveur-web/README.md` pour les details.

---

## Flux de licence

```
[Admin] --> Interface web --> Genere une licence
                                      |
                                      v
                            [Cle envoyee au client]
                                      |
                                      v
[Client] --> Excel add-in --> Saisit la cle
                                      |
                                      v
                     M_LicenceWeb.bas --> API REST --> Supabase
                                      |                   |
                                      v                   v
                              Licence validee     Activation enregistree
```

## Types de licences

| Type | Description |
|------|-------------|
| STANDARD | Acces aux fonctionnalites de base |
| PREMIUM | Acces complet (API REST, Power BI, SQL) |
| Trial | Periode d'essai 30 jours (automatique) |

## Correspondance IATF 16949

L'add-in couvre l'ensemble des exigences de la norme : AMDEC/FMEA, plans de controle,
metrologie (MSA, capabilite, cartes de controle), audits, non-conformites, 8D, CAPA,
gestion documentaire, fournisseurs, TPM, KPIs, revue de direction.
