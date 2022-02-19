Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class ObjectStateChangedEventArgs
    Inherits System.EventArgs

    Public EventTarget As E_EVENTTARGET
    Public Index As Integer
    Public Cancel As Boolean
    Public DestinationIndex As Integer
    Public InitialRowIndex As Integer
    Public FinalRowIndex As Integer
    Public InitialColumnIndex As Integer
    Public FinalColumnIndex As Integer
    Public StartDate As AGVBE.DateTime
    Public EndDate As AGVBE.DateTime

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        EventTarget = 0
        Index = 0
        StartDate = New AGVBE.DateTime()
        EndDate = New AGVBE.DateTime()
        Cancel = False
    End Sub
End Class
