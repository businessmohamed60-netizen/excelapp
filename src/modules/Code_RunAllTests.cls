Public Sub RunAllTests()
    T_Repository
    T_Client
    T_Audit
    T_Constat
    RunUtilsTests
    T_NonConformite
    T_Capa
    T_8D
    T_AuditProgramme
    T_Document
    T_Exigence
    T_Kpi
    T_Amdec
    T_PlanControle
    T_Caracteristique
    MsgBox "Tests terminés — voir Immediate Window.", vbInformation
End Sub