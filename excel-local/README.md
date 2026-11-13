# IATF Consultant Assistant — Add-in Excel (Local)

Ce dossier contient **tout ce qui doit etre installe sur le poste local** (Excel) pour faire fonctionner l'add-in IATF 16949.

## Structure

```
excel-local/
├── src/
│   ├── classes/         26 classes metier (C*.cls)
│   ├── modules/         22 modules VBA + 62 handlers (Code_*.cls) + M_LicenceWeb.bas
│   ├── forms/           10 formulaires (.frm) + fragments de reference
│   └── tests/           28 suites de tests (T_*.bas)
├── ribbon/              Ribbon.xml (7 onglets unifies)
├── scripts/             build.ps1, install.ps1, uninstall.ps1, verifier.ps1
├── docs/                Manuel d'integration + documentation
├── build/               Dossier de sortie pour le .xlam
└── templates/           Modeles .xlsx (a creer)
```

## Installation

1. Lisez `docs/MANUEL_INTEGRATION_EXCEL.md` pour la procedure complete
2. Ou executez `scripts/build.ps1` dans PowerShell pour un guide pas-a-pas

## Licence Web

Le module `M_LicenceWeb.bas` permet de verifier les licences aupres du serveur web (dossier `serveur-web/`).
Avant utilisation, modifiez dans ce fichier :
- `LICENCE_API_URL` : l'URL de votre serveur de licences
- `LICENCE_API_KEY` : la cle API (anon key) de votre projet

Voir `docs/MANUEL_INTEGRATION_EXCEL.md` > section "Licence Web" pour les details.
