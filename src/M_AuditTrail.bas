Attribute VB_Name = "M_AuditTrail"
Option Explicit

' ============================================================
' M_AuditTrail — Journalisation des modifications
' ============================================================

Public Sub Journaliser(ByVal action As String, ByVal coll As String, _
                       ByVal objetId As String, ByVal libelle As String, _
                       Optional ByVal champs As String = "", _
                       Optional ByVal ancienne As String = "", _
                       Optional ByVal nouvelle As String = "")
    On Error GoTo ErrHandler
    Dim e As CAuditTrail
    Set e = New CAuditTrail
    e.InitNew action, coll, objetId, libelle
    e.ChampsModifies = champs
    e.AncienneValeur = ancienne
    e.NouvelleValeur = nouvelle
    g_CollMgr.Save e, "audittrail"
    g_CollMgr.InvalidateCache "audittrail"
    Exit Sub
ErrHandler:
    LogEx "M_AuditTrail", "Journaliser", Err
End Sub

' Charge les entrées filtrées par objet
Public Function JournalObjet(ByVal coll As String, ByVal objetId As String) As Collection
    Dim all As Object, k As Variant, result As New Collection
    Set all = g_CollMgr.LoadCollection("audittrail")
    For Each k In all.Keys
        If all(k).Collection = coll And all(k).ObjetId = objetId Then
            result.Add all(k)
        End If
    Next k
    Set JournalObjet = result
End Function

' Charge les entrées d'une période
Public Function JournalPeriode(ByVal debut As Date, ByVal fin As Date) As Collection
    Dim all As Object, k As Variant, result As New Collection
    Set all = g_CollMgr.LoadCollection("audittrail")
    For Each k In all.Keys
        If all(k).DateHeure >= debut And all(k).DateHeure <= fin Then
            result.Add all(k)
        End If
    Next k
    Set JournalPeriode = result
End Function

' Purge les entrées antérieures à N jours (paramétrable)
Public Sub PurgerJournal(ByVal joursConservation As Long)
    Dim all As Object, k As Variant, seuil As Date
    seuil = Date - joursConservation
    Set all = g_CollMgr.LoadCollection("audittrail")
    For Each k In all.Keys
        If all(k).DateHeure < seuil Then
            g_CollMgr.DeleteById k, "audittrail"
        End If
    Next k
    g_CollMgr.InvalidateCache "audittrail"
End Sub