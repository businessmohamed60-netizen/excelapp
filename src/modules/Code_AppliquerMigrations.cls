Option Explicit

Private Const CURRENT_SCHEMA As String = "1.1.0"

Public Sub AppliquerMigrations()
    On Error GoTo ErrHandler
    Dim schemaActuel As String
    schemaActuel = g_Config.Version
    If schemaActuel = CURRENT_SCHEMA Then Exit Sub
    Log LogInfo, "M_Migration", "Migration " & schemaActuel & " → " & CURRENT_SCHEMA
    If schemaActuel = "0.1.0" Or schemaActuel = "1.0.0" Then Migrer_V1_1
    g_Config.Version = CURRENT_SCHEMA
    g_Config.Save
    Log LogInfo, "M_Migration", "Migration terminée"
    Exit Sub
ErrHandler:
    LogEx "M_Migration", "AppliquerMigrations", Err
End Sub

Private Sub Migrer_V1_1()
    On Error Resume Next
    Dim fname As String, dom As Object, item As Object, node As Object
    fname = DataPath() & "nc.xml"
    If Dir(fname) = "" Then Exit Sub
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    If Not dom.Load(fname) Then Exit Sub
    For Each item In dom.documentElement.ChildNodes
        If item.selectSingleNode("Priorite") Is Nothing Then
            Set node = dom.createElement("Priorite")
            node.Text = "Moyenne"
            item.appendChild node
        End If
    Next item
    dom.save fname
    Log LogInfo, "M_Migration", "Ajout Priorite aux NC"
End Sub