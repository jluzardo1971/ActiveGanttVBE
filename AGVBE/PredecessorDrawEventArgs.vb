Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class PredecessorDrawEventArgs
    Inherits System.EventArgs

    Public CustomDraw As Boolean
    'Public Graphics As Graphics
    Public PredecessorIndex As Integer
    Public TaskIndex As Integer
    Public PredecessorType As E_CONSTRAINTTYPE

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        CustomDraw = Nothing
        'Graphics = Nothing
        PredecessorIndex = Nothing
        TaskIndex = Nothing
        PredecessorType = Nothing
    End Sub
End Class
