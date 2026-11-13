Attribute VB_Name = "M_I18n"
Option Explicit

' ============================================================
' M_I18n — Table de chaînes centralisée (FR / EN)
' ============================================================

Private m_lang As String
Private m_table As Object

Private Sub InitTable()
    Set m_table = CreateObject("Scripting.Dictionary")

    ' ---- FRANÇAIS ----
    AddL "FR", "app.title",       "IATF Consultant Assistant"
    AddL "FR", "menu.clients",    "Clients"
    AddL "FR", "menu.audits",     "Audits"
    AddL "FR", "menu.nc",         "Non-conformités"
    AddL "FR", "menu.capa",       "CAPA"
    AddL "FR", "menu.amdec",      "AMDEC"
    AddL "FR", "menu.spc",        "SPC"
    AddL "FR", "menu.msa",        "MSA"
    AddL "FR", "menu.ppap",       "PPAP"
    AddL "FR", "menu.apqp",       "APQP"
    AddL "FR", "menu.settings",   "Paramètres"
    AddL "FR", "menu.about",      "À propos"
    AddL "FR", "btn.new",         "Nouveau"
    AddL "FR", "btn.save",        "Enregistrer"
    AddL "FR", "btn.delete",      "Supprimer"
    AddL "FR", "btn.close",       "Fermer"
    AddL "FR", "msg.confirm_delete", "Confirmer la suppression ?"
    AddL "FR", "err.required",    "Ce champ est requis."

    ' ---- ENGLISH ----
    AddL "EN", "app.title",       "IATF Consultant Assistant"
    AddL "EN", "menu.clients",    "Clients"
    AddL "EN", "menu.audits",     "Audits"
    AddL "EN", "menu.nc",         "Non-conformities"
    AddL "EN", "menu.capa",       "CAPA"
    AddL "EN", "menu.amdec",      "FMEA"
    AddL "EN", "menu.spc",        "SPC"
    AddL "EN", "menu.msa",        "MSA"
    AddL "EN", "menu.ppap",       "PPAP"
    AddL "EN", "menu.apqp",       "APQP"
    AddL "EN", "menu.settings",   "Settings"
    AddL "EN", "menu.about",      "About"
    AddL "EN", "btn.new",         "New"
    AddL "EN", "btn.save",        "Save"
    AddL "EN", "btn.delete",      "Delete"
    AddL "EN", "btn.close",       "Close"
    AddL "EN", "msg.confirm_delete", "Confirm deletion?"
    AddL "EN", "err.required",    "This field is required."

    m_lang = "FR"
End Sub

Private Sub AddL(ByVal lang As String, ByVal key As String, ByVal val As String)
    m_table.Add lang & "." & key, val
End Sub

Public Function T(ByVal key As String) As String
    If m_table Is Nothing Then InitTable
    Dim k As String
    k = m_lang & "." & key
    If m_table.Exists(k) Then
        T = m_table(k)
    Else
        T = "[" & key & "]"
    End If
End Function

Public Sub SetLanguage(ByVal lang As String)
    m_lang = UCase$(lang)
End Sub