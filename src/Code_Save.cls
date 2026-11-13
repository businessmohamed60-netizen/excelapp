Public Sub Save(ByVal obj As Object, ByVal collection As String)
    Dim path As String
    path = DataPath() & collection & ".xml"

    ' Détection de conflit (si le fichier a été modifié par un autre poste)
    Dim hashAttendu As String
    hashAttendu = HashAttendu(collection)
    If Len(hashAttendu) > 0 Then
        If VerifierConflit(path, hashAttendu) Then
            Err.Raise vbObjectError + 4001, "CRepository", _
                "Conflit détecté sur " & collection & ".xml : le fichier a été modifié " & _
                "par un autre utilisateur. Rechargez avant de sauvegarder."
        End If
    End If

    ' ... écriture XML existante ...
    ' Après écriture :
    EnregistrerHash collection, path
End Sub