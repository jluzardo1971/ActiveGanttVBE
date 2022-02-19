Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class ToolTipEventArgs
    Inherits System.EventArgs

    Public InitialRowIndex As Integer
    Public FinalRowIndex As Integer
    Public TaskIndex As Integer
    Public MilestoneIndex As Integer
    Public PercentageIndex As Integer
    Public RowIndex As Integer
    Public CellIndex As Integer
    Public ColumnIndex As Integer
    Public InitialStartDate As AGVBE.DateTime
    Public InitialEndDate As AGVBE.DateTime
    Public StartDate As AGVBE.DateTime
    Public EndDate As AGVBE.DateTime
    Public XStart As Integer
    Public XEnd As Integer
    Public Operation As E_OPERATION
    Public EventTarget As E_EVENTTARGET
    Public TaskPosition As String
    Public PredecessorPosition As String
    Public X As Integer
    Public Y As Integer
    Public CustomDraw As Boolean
    '<Description("A pointer to the Graphics object of the control.")> _
    Public Graphics As Canvas
    Public ToolTipType As E_TOOLTIPTYPE

    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        InitialRowIndex = Nothing
        FinalRowIndex = Nothing
        RowIndex = Nothing
        TaskIndex = Nothing
        MilestoneIndex = Nothing
        PercentageIndex = Nothing
        CellIndex = Nothing
        ColumnIndex = Nothing
        StartDate = New AGVBE.DateTime()
        EndDate = New AGVBE.DateTime()
        InitialStartDate = New AGVBE.DateTime()
        InitialEndDate = New AGVBE.DateTime()
        XStart = Nothing
        XEnd = Nothing
        X = Nothing
        Y = Nothing
        Operation = Nothing
        EventTarget = Nothing
        ToolTipType = Nothing
    End Sub
End Class
