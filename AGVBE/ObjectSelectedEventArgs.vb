Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection

Public Class ObjectSelectedEventArgs
    Inherits System.EventArgs

    Public EventTarget As E_EVENTTARGET
    Public ObjectIndex As Integer
    Public ParentObjectIndex As Integer

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        EventTarget = Nothing
        ObjectIndex = Nothing
        ParentObjectIndex = Nothing
    End Sub
End Class
