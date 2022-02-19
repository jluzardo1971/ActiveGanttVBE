Public Class PredecessorExceptionEventArgs
    Inherits System.EventArgs

    Public PredecessorIndex As Integer
    Public PredecessorType As E_CONSTRAINTTYPE

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        PredecessorIndex = Nothing
        PredecessorType = Nothing
    End Sub
End Class
