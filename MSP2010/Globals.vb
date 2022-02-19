
Public Enum SYS_ERRORS
    ERR_ADDMODE_G = 51141
    MP_REMOVE_1 = 51596
    MP_REMOVE_2 = 51597
    MP_REMOVE_3 = 51598
    MP_REMOVE_4 = 51599
    MP_ITEM_1 = 51600
    MP_ITEM_2 = 51601
    MP_ITEM_3 = 51602
    MP_ITEM_4 = 51603
    MP_ADD_1 = 51604
    MP_ADD_2 = 51605
    MP_ADD_3 = 51606
    MP_SET_KEY = 51607
End Enum

Public Module Globals

    Friend Sub g_ErrorReport(ByVal ErrNumber As Integer, ByVal ErrDescription As String, ByVal ErrSource As String)
        '//MessageBox.Show(System.Convert.ToString(ErrNumber) & ": " & ErrDescription & " (" & ErrSource & ")", "AGVBN Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
    End Sub

    Friend Function g_StrIsNumeric(ByVal Expression As String) As Boolean
        Dim dDummy As Double
        Return Double.TryParse(Expression, dDummy)
    End Function

    Public Function g_Format(ByVal Expression As Integer, ByVal sFormat As String) As String
        Return Expression.ToString(sFormat)
    End Function

End Module
