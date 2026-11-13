Option Explicit

' ============================================================
' CDocManager — CRUD + workflow + versions pour les documents
' ============================================================

Private m_Repo As CRepository

Private Sub Class_Initialize()
    Set m_Repo = New CRepository
End Sub

' ---- CRUD documents ----
Public Sub SauvegarderDocument(ByVal doc As CDocument)
    If Not doc.Validate() Then
        Err.Raise vbObjectError + 8001, "CDocManager", "Document invalide : code, titre et propriétaire requis."
    End If
    If Len(doc.Code) = 0 Then doc.AssignCode()
    doc.CalculerProchaineRevue
    g_CollMgr.Save doc, "documents"
    Log LogInfo, "CDocManager", "Document enregistré : " & doc.Code
End Sub

Public Function ChargerDocument(ByVal id As String) As CDocument
    Dim d As Object
    Set d = g_CollMgr.LoadCollection("documents")
    If d.Exists(id) Then Set ChargerDocument = d(id)
End Function

Public Sub SupprimerDocument(ByVal id As String)
    g_CollMgr.DeleteById id, "documents"
    g_CollMgr.InvalidateCache "documents"
    Log LogInfo, "CDocManager", "Document supprimé : " & id
End Sub

' ---- Workflow ----
Public Sub ChangerStatut(ByVal doc As CDocument, ByVal nouveauStatut As String)
    If Not doc.PeutPasserA(nouveauStatut) Then
        Err.Raise vbObjectError + 8002, "CDocManager", _
            "Transition invalide : " & doc.Statut & " → " & nouveauStatut
    End If
    doc.Statut = nouveauStatut
    If nouveauStatut = "Publié" Then
        doc.DatePublication = Date
        If doc.DateDerniereRevue = 0 Then doc.DateDerniereRevue = Date
    ElseIf nouveauStatut = "Archivé" Then
        doc.DateArchivage = Date
    End If
    doc.CalculerProchaineRevue
    g_CollMgr.Save doc, "documents"
    g_CollMgr.InvalidateCache "documents"
    Log LogInfo, "CDocManager", doc.Code & " → " & nouveauStatut
End Sub

' ---- Nouvelle version ----
Public Function CreerNouvelleVersion(ByVal doc As CDocument, ByVal commentaire As String, _
                                     ByVal cheminFichier As String) As CVersionDoc
    Dim v As CVersionDoc: Set v = New CVersionDoc
    v.InitNew
    v.DocumentId = doc.Id
    v.Commentaire = commentaire
    v.CheminFichier = cheminFichier
    If Len(cheminFichier) > 0 And Dir(cheminFichier) <> "" Then
        v.HashFichier = SHA256FichierHex(cheminFichier)
    End If
    ' Calculer le numéro de version
    v.NumeroVersion = CalculerNumeroVersionSuivant(doc)
    g_CollMgr.Save v, "docversions"
    ' Mettre à jour le document
    doc.VersionCouranteId = v.Id
    If doc.Statut = "Publié" Then
        doc.Statut = "En revue"  ' une nouvelle version remet en revue
    End If
    g_CollMgr.Save doc, "documents"
    Log LogInfo, "CDocManager", "Nouvelle version " & v.NumeroVersion & " pour " & doc.Code
    Set CreerNouvelleVersion = v
End Function

Private Function CalculerNumeroVersionSuivant(ByVal doc As CDocument) As String
    If Len(doc.VersionCouranteId) = 0 Then
        CalculerNumeroVersionSuivant = "1.0"
        Exit Function
    End If
    Dim versions As Object, k As Variant
    Set versions = g_CollMgr.LoadCollection("docversions")
    If versions.Exists(doc.VersionCouranteId) Then
        Dim v As CVersionDoc: Set v = versions(doc.VersionCouranteId)
        Dim parts() As String: parts = Split(v.NumeroVersion, ".")
        If UBound(parts) >= 1 Then
            CalculerNumeroVersionSuivant = parts(0) & "." & CStr(CLng(parts(1)) + 1)
            Exit Function
        End If
    End If
    CalculerNumeroVersionSuivant = "1.0"
End Function

' ---- Approuver / Rejeter une version ----
Public Sub ApprouverVersion(ByVal versionId As String, ByVal approbateur As String)
    Dim v As CVersionDoc
    Dim versions As Object: Set versions = g_CollMgr.LoadCollection("docversions")
    If Not versions.Exists(versionId) Then Exit Sub
    Set v = versions(versionId)
    v.StatutApprobation = "Approuvée"
    v.Approbateur = approbateur
    v.DateApprobation = Date
    g_CollMgr.Save v, "docversions"
    Log LogInfo, "CDocManager", "Version approuvée : " & versionId
End Sub

Public Sub RejeterVersion(ByVal versionId As String, ByVal motif As String)
    Dim v As CVersionDoc
    Dim versions As Object: Set versions = g_CollMgr.LoadCollection("docversions")
    If Not versions.Exists(versionId) Then Exit Sub
    Set v = versions(versionId)
    v.StatutApprobation = "Rejetée"
    v.MotifRejet = motif
    g_CollMgr.Save v, "docversions"
    Log LogInfo, "CDocManager", "Version rejetée : " & versionId
End Sub

' ---- Versions d'un document (triées par numéro décroissant) ----
Public Function VersionsDeDocument(ByVal docId As String) As Collection
    Dim versions As Object, k As Variant, arr As New Collection
    Set versions = g_CollMgr.LoadCollection("docversions")
    Dim tmp As Object: Set tmp = CreateObject("Scripting.Dictionary")
    For Each k In versions.Keys
        If versions(k).DocumentId = docId Then tmp.Add k, versions(k)
    Next k
    ' Tri décroissant sur numéro de version (comme chaîne, suffisant)
    Dim keys As Variant, i As Long, j As Long, t As String
    keys = tmp.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If tmp(keys(j)).NumeroVersion > tmp(keys(i)).NumeroVersion Then
                t = keys(i): keys(i) = keys(j): keys(j) = t
            End If
        Next j
    Next i
    For i = LBound(keys) To UBound(keys): arr.Add tmp(keys(i)): Next i
    Set VersionsDeDocument = arr
End Function

' ---- Recherche par code ----
Public Function TrouverParCode(ByVal code As String) As CDocument
    Set TrouverParCode = g_CollMgr.FindByKey("documents", "Code", code)
End Function

' ---- Vérification d'intégrité des fichiers ----
Public Function VerifierIntegriteVersion(ByVal v As CVersionDoc) As Boolean
    If Len(v.CheminFichier) = 0 Then VerifierIntegriteVersion = True: Exit Function
    If Dir(v.CheminFichier) = "" Then VerifierIntegriteVersion = False: Exit Function
    Dim hashActuel As String
    hashActuel = SHA256FichierHex(v.CheminFichier)
    VerifierIntegriteVersion = (hashActuel = v.HashFichier)
End Function