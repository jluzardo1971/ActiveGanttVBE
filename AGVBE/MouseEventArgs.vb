Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class MouseEventArgs
    Inherits System.EventArgs

    Public X As Integer
    Public Y As Integer
    Public EventTarget As E_EVENTTARGET
    Public Operation As E_OPERATION
    Public Button As E_MOUSEBUTTONS
    Public Cancel As Boolean

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        X = Nothing
        Y = Nothing
        EventTarget = Nothing
        Operation = Nothing
        Button = Nothing
        Cancel = Nothing
    End Sub
End Class
