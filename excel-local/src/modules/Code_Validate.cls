Public Function Validate() As Boolean
    If Len(m_NcId) = 0 Then Validate = False: Exit Function
    If Len(Trim$(m_Description)) = 0 Then Validate = False: Exit Function
    If m_DateEcheance < m_DateOuverture Then Validate = False: Exit Function
    Validate = True
End Function