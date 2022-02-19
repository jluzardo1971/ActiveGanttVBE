Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class ScrollEventArgs
    Inherits System.EventArgs

    Public ScrollBarType As E_SCROLLBAR
    Public Offset As Integer

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        ScrollBarType = Nothing
        Offset = Nothing
    End Sub
End Class
