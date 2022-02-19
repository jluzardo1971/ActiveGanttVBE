Option Explicit On 

Imports System
Imports System.ComponentModel
Imports System.Reflection


Public Class CustomTierDrawEventArgs
    Inherits System.EventArgs

    Public Text As String
    Public CustomDraw As Boolean
    Public StyleIndex As String
    Public TierPosition As E_TIERPOSITION
    Public StartDate As AGVBE.DateTime
    Public EndDate As AGVBE.DateTime
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
    Public LeftTrim As Integer
    Public RightTrim As Integer
    'Public Graphics As DrawingContext
    Public Interval As E_INTERVAL
    Public Factor As Integer


    Friend Sub New()
        Clear()
    End Sub

    Friend Sub Clear()
        Text = ""
        CustomDraw = False
        StyleIndex = ""
        TierPosition = E_TIERPOSITION.SP_UPPER
        StartDate = New AGVBE.DateTime()
        EndDate = New AGVBE.DateTime()
        Left = 0
        Top = 0
        Right = 0
        Bottom = 0
        LeftTrim = 0
        RightTrim = 0
        'Graphics = Nothing
        Interval = E_INTERVAL.IL_SECOND
        Factor = 0
    End Sub
End Class
