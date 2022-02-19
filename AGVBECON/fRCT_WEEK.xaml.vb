﻿Imports AGVBE

Partial Public Class fRCT_WEEK
    Inherits Page

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Page_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim oView As clsView
        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_MINUTE, 20, E_TIERTYPE.ST_MONTH, E_TIERTYPE.ST_DAYOFWEEK, E_TIERTYPE.ST_DAYOFWEEK, "View1")
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TickMarkArea.Visible = False

        ActiveGanttVBECtl1.TierFormatScope = E_TIERFORMATSCOPE.TFS_CONTROL
        ActiveGanttVBECtl1.TierFormat.DayOfWeekIntervalFormat = "ddd d"

        ActiveGanttVBECtl1.CurrentView = "View1"

        Dim i As Integer
        For i = 1 To 50
            ActiveGanttVBECtl1.Rows.Add("K" & i.ToString())
        Next

        Dim oTimeBlock As clsTimeBlock
        Dim dtDate As AGVBE.DateTime
        dtDate = New AGVBE.DateTime(2000, 1, 1, 0, 0, 0) 'RCT_WEEK's BaseDate will ignore Year, Month and Day values

        oTimeBlock = ActiveGanttVBECtl1.TimeBlocks.Add("TimeBlock1")
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_WEEK
        oTimeBlock.BaseWeekDay = E_WEEKDAY.WD_SATURDAY
        oTimeBlock.BaseDate = dtDate
        oTimeBlock.DurationInterval = E_INTERVAL.IL_HOUR
        oTimeBlock.DurationFactor = 48
    End Sub

    Private Sub cmdBack_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub
End Class
