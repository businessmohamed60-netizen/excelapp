Attribute VB_Name = "M_LicenceWeb"
Option Explicit

' ============================================================
' M_LicenceWeb — Verification de licence via serveur web
' Communique avec l'API REST de gestion de licences
' ============================================================

Private Const LICENCE_API_URL As String = "https://VOTRE-PROJET.supabase.co/functions/v1/license-api"
Private Const LICENCE_API_KEY As String = "VOTRE_ANON_KEY"

' Verifie une cle de licence aupres du serveur web
' Retourne True si la licence est valide
Public Function VerifierLicenceWeb(ByVal cleLicence As String, ByVal machineId As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    Dim url As String
    url = LICENCE_API_URL & "/verify?key=" & UCase$(Trim$(cleLicence)) & "&machineId=" & machineId
    
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & LICENCE_API_KEY
    http.setRequestHeader "X-Client-Info", "IATF-Excel-AddIn"
    http.send
    
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText
        ' Si la reponse contient "valid":true, la licence est valide
        VerifierLicenceWeb = (InStr(response, """valid"":true") > 0)
    Else
        VerifierLicenceWeb = False
    End If
    
    Exit Function
ErrHandler:
    LogEx "M_LicenceWeb", "VerifierLicenceWeb", Err
    VerifierLicenceWeb = False
End Function

' Active une licence sur cette machine
Public Function ActiverLicenceWeb(ByVal cleLicence As String, ByVal machineId As String, Optional ByVal machineInfo As String = "") As Boolean
    On Error GoTo ErrHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    Dim body As String
    body = "{""license_key"":""" & UCase$(Trim$(cleLicence)) & """,""machine_id"":""" & machineId & """,""machine_info"":""" & machineInfo & """}"
    
    http.Open "POST", LICENCE_API_URL & "/activate", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & LICENCE_API_KEY
    http.setRequestHeader "X-Client-Info", "IATF-Excel-AddIn"
    http.send body
    
    If http.Status = 200 Then
        ActiverLicenceWeb = (InStr(http.responseText, """success"":true") > 0)
    Else
        ActiverLicenceWeb = False
    End If
    
    Exit Function
ErrHandler:
    LogEx "M_LicenceWeb", "ActiverLicenceWeb", Err
    ActiverLicenceWeb = False
End Function

' Recupere les informations de licence depuis le serveur
Public Function InfosLicenceWeb(ByVal cleLicence As String) As String
    On Error GoTo ErrHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    http.Open "GET", LICENCE_API_URL & "/verify?key=" & UCase$(Trim$(cleLicence)), False
    http.setRequestHeader "Authorization", "Bearer " & LICENCE_API_KEY
    http.setRequestHeader "X-Client-Info", "IATF-Excel-AddIn"
    http.send
    
    If http.Status = 200 Then
        InfosLicenceWeb = http.responseText
    Else
        InfosLicenceWeb = ""
    End If
    
    Exit Function
ErrHandler:
    LogEx "M_LicenceWeb", "InfosLicenceWeb", Err
    InfosLicenceWeb = ""
End Function

' Verifie et active la licence en une seule operation
' Utilise MachineFingerprint() comme identifiant machine
Public Function VerifierEtActiverLicence(ByVal cleLicence As String) As Boolean
    Dim mid As String
    mid = MachineFingerprint()
    
    ' D'abord verifier
    If VerifierLicenceWeb(cleLicence, mid) Then
        VerifierEtActiverLicence = True
        Exit Function
    End If
    
    ' Sinon tenter l'activation
    Dim info As String
    info = Environ$("COMPUTERNAME") & "|" & Environ$("USERNAME")
    VerifierEtActiverLicence = ActiverLicenceWeb(cleLicence, mid, info)
End Function
