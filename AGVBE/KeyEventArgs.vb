Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class KeyEventArgs
    Inherits System.EventArgs

    Public KeyCode As System.Windows.Input.Key
    Public Cancel As Boolean
    Public CharacterCode As Char

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        KeyCode = Nothing
        Cancel = Nothing
    End Sub
End Class
