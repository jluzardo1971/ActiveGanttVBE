Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class ErrorEventArgs
    Inherits System.EventArgs

    Public Number As Integer
    Public Description As String
    Public Source As String

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        Number = Nothing
        Description = Nothing
        Source = Nothing
    End Sub
End Class
