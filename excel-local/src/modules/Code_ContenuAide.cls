Option Explicit

Private Function ContenuAide(ByVal cle As String) As String
    Select Case cle
        Case "FAudit"
            ContenuAide = "Formulaire Audits" & vbCrLf & _
                "— Sélectionnez un client puis renseignez le processus, les dates et l'auditeur." & vbCrLf & _
                "— Cliquez sur 'Constats...' pour ajouter les constats de l'audit." & vbCrLf & _
                "— 'Générer rapport' produit un fichier .xlsx dans le dossier reports/." & vbCrLf & _
                "— Normes référencées : ISO 9001 §9.2, IATF 16949 §9.2.2.3."
        Case "FNonConformite"
            ContenuAide = "Formulaire Non-conformités" & vbCrLf & _
                "— Numérotation automatique au format NC-AAAA-NNN." & vbCrLf & _
                "— Une NC peut être liée à un audit et un constat." & vbCrLf & _
                "— Une NC ne peut être clôturée que si toutes ses CAPA sont clôturées."
        Case "FCapa"
            ContenuAide = "Formulaire CAPA" & vbCrLf & _
                "— Action corrective ou préventive." & vbCrLf & _
                "— Statuts : Ouverte → En cours → Réalisée → Vérifiée → Clôturée."
        Case "F8D"
            ContenuAide = "Formulaire 8D" & vbCrLf & _
                "— 8 étapes : D1 équipe, D2 problème, D3 actions immédiates," & vbCrLf & _
                "  D4 causes racines, D5 actions correctives, D6 mise en œuvre," & vbCrLf & _
                "  D7 prévention, D8 clôture." & vbCrLf & _
                "— Un 8D ne peut être ouvert que sur une NC Majeure ou Critique."
        Case "FApqa"
            ContenuAide = "Formulaire APQP" & vbCrLf & _
                "— 5 phases AIAG : Planification → Conception produit →" & vbCrLf & _
                "  Conception processus → Validation → Production." & vbCrLf & _
                "— 20 % d'avancement par phase validée."
        Case "FPpap"
            ContenuAide = "Formulaire PPAP" & vbCrLf & _
                "— 18 éléments selon AIAG PPAP 4e éd." & vbCrLf & _
                "— Niveaux : 1 (PSW seul) à 5 (audit site)."
        Case "FAmdec"
            ContenuAide = "Formulaire AMDEC" & vbCrLf & _
                "— S, O, D cotées de 1 à 10. RPN = S × O × D." & vbCrLf & _
                "— AP : H (action requise) / M / L (matrice AIAG-VDA 2019)."
        Case "FSpc"
            ContenuAide = "Formulaire SPC" & vbCrLf & _
                "— Mesures ligne par ligne. n = 2 à 10 (défaut 5)." & vbCrLf & _
                "— Indices : Cp, Cpk, Pp, Ppk. Règles Nelson R1 à R4."
        Case "FMsa"
            ContenuAide = "Formulaire MSA" & vbCrLf & _
                "— GRR : 10 pièces × 3 opérateurs × 3 essais (AIAG)." & vbCrLf & _
                "— %GRR < 10 % acceptable ; 10–30 % conditionnel ; > 30 % rejeté."
        Case "FMain"
            ContenuAide = "Tableau de bord" & vbCrLf & _
                "— 8 KPI synthétiques, alertes prioritaires, activités récentes."
        Case Else
            ContenuAide = "Aide non disponible pour ce formulaire."
    End Select
End Function

Public Sub AfficherAide(ByVal cleFormulaire As String)
    FAide.Contenu = ContenuAide(cleFormulaire)
    FAide.Show vbModeless
End Sub