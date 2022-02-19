Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class NodeEventArgs
    Inherits System.EventArgs

    Public Index As Integer

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        Index = Nothing
    End Sub
End Class
