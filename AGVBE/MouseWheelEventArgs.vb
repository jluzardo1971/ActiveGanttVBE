Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class MouseWheelEventArgs
    Inherits System.EventArgs


    Public X As Integer
    Public Y As Integer
    Public Button As E_MOUSEBUTTONS
    Public Delta As Integer

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        X = Nothing
        Y = Nothing
        Button = Nothing
        Delta = Nothing
    End Sub

End Class
