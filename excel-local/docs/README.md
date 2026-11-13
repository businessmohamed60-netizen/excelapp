# IATFConsultantAssistant

Arborescence préparée à partir des fichiers fournis.

- `src/` : modules VBA, classes, formulaires et suites de tests identifiés par leur nom.
- `ribbon/` : fragments XML du ruban, conservés séparément.
- `scripts/` : scripts PowerShell présents dans l'archive.
- `docs/` : documentation source.
- `_a_trier/` : fragments, variantes et fichiers non classifiables sans risque.

**Important :** aucun fichier `.xlam` compilé ni modèle `.xlsx` n'était présent dans l'archive source. Les dossiers `build/`, `templates/` et `dist/templates/` sont donc laissés prêts pour les livrables. Les formulaires `.frm` sont des exports texte ; les fichiers binaires `.frx` associés restent à générer/exporter depuis l'éditeur VBA si nécessaire.
