Public Class DeflaterPending
    Inherits PendingBuffer

    Public Sub New()
        MyBase.New(DeflaterConstants.PENDING_BUF_SIZE)
    End Sub

End Class