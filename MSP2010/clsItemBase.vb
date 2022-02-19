Option Explicit On 

Public MustInherit Class clsItemBase

    Friend mp_sKey As String
    Friend mp_lIndex As Integer

    Friend Sub New()
        mp_sKey = ""
        mp_lIndex = 0
    End Sub

    Public Property Index() As Integer
        Get
            Return mp_lIndex
        End Get
        Set(ByVal Value As Integer)
            mp_lIndex = Value
        End Set
    End Property

End Class
