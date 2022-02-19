Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class ObjectAddedEventArgs
    Inherits System.EventArgs

    Public TaskIndex As Integer
    Public PredecessorObjectIndex As Integer
    Public PredecessorTaskIndex As Integer
    Public PredecessorType As E_CONSTRAINTTYPE
    Public TaskKey As String
    Public PredecessorTaskKey As String
    Public EventTarget As E_EVENTTARGET

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        TaskIndex = Nothing
        PredecessorObjectIndex = Nothing
        PredecessorTaskIndex = Nothing
        PredecessorType = Nothing
        TaskKey = Nothing
        PredecessorTaskKey = Nothing
        EventTarget = Nothing
    End Sub
End Class
