Public Function EtapeCourante() As String
    If IsDate(m_D8_Date) And m_D8_Date > 0 Then
        EtapeCourante = "D8"
    ElseIf IsDate(m_D7_Date) And m_D7_Date > 0 Then
        EtapeCourante = "D7"
    ElseIf IsDate(m_D6_Date) And m_D6_Date > 0 Then
        EtapeCourante = "D6"
    ElseIf IsDate(m_D5_Date) And m_D5_Date > 0 Then
        EtapeCourante = "D5"
    ElseIf IsDate(m_D4_Date) And m_D4_Date > 0 Then
        EtapeCourante = "D4"
    ElseIf IsDate(m_D3_Date) And m_D3_Date > 0 Then
        EtapeCourante = "D3"
    ElseIf IsDate(m_D2_Date) And m_D2_Date > 0 Then
        EtapeCourante = "D2"
    Else
        EtapeCourante = "D1"
    End If
End Function

Public Function ValidateChronologie() As Boolean
    Dim dates(1 To 7) As Date
    Dim flags(1 To 7) As Boolean
    Dim i As Long, last As Date
    flags(1) = IsDate(m_D2_Date) And m_D2_Date > 0: If flags(1) Then dates(1) = m_D2_Date
    flags(2) = IsDate(m_D3_Date) And m_D3_Date > 0: If flags(2) Then dates(2) = m_D3_Date
    flags(3) = IsDate(m_D4_Date) And m_D4_Date > 0: If flags(3) Then dates(3) = m_D4_Date
    flags(4) = IsDate(m_D5_Date) And m_D5_Date > 0: If flags(4) Then dates(4) = m_D5_Date
    flags(5) = IsDate(m_D6_Date) And m_D6_Date > 0: If flags(5) Then dates(5) = m_D6_Date
    flags(6) = IsDate(m_D7_Date) And m_D7_Date > 0: If flags(6) Then dates(6) = m_D7_Date
    flags(7) = IsDate(m_D8_Date) And m_D8_Date > 0: If flags(7) Then dates(7) = m_D8_Date

    last = 0
    For i = 1 To 7
        If flags(i) Then
            If last > 0 And dates(i) < last Then Exit Function
            last = dates(i)
        End If
    Next i
    ValidateChronologie = True
End Function