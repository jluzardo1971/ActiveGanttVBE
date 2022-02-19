Option Explicit On

Public Class TextEditEventArgs
    Inherits System.EventArgs

    Public ObjectType As E_TEXTOBJECTTYPE
    Public ObjectIndex As Integer
    Public ParentObjectIndex As Integer
    Public Text As String

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        ObjectType = Nothing
        ObjectIndex = 0
        ParentObjectIndex = 0
        Text = ""
    End Sub

End Class
