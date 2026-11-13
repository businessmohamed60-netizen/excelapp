Procédure finale d'assemblage
Créer un classeur .xlsm vide, ouvrir l'éditeur VBA (Alt+F11).

Ajouter les 3 références : Scripting Runtime, MSXML v6.0, Forms 2.0.

Créer les 20 modules .bas listés ci-dessus (Insertion > Module, renommer via F4, coller le code).

Créer les 15 classes .cls (Insertion > Module de classe, renommer, coller).

Créer les 18 UserForms (Insertion > UserForm, renommer, dessiner les contrôles listés dans les en-têtes de chaque bloc, coller le code).

Créer les templates .xlsx dans %APPDATA%\IATFConsultantAssistant\templates\ (cellules nommées, voir phases 2, 3, 4).

Injecter le Ribbon XML via RibbonX Editor.

Compiler (Débogage > Compiler VBAProject) — doit passer sans erreur.

Exécuter RunAllTests (Ctrl+G) — toutes les assertions vertes.

Enregistrer en .xlam et déployer via install.ps1.