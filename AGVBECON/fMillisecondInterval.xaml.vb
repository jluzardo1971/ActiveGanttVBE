Imports AGVBE

Partial Public Class fMillisecondInterval
    Inherits Page

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Page_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim oView As clsView
        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_MILLISECOND, 5, E_TIERTYPE.ST_MINUTE, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_SECOND, "MSI")
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.TierFormat.MinuteIntervalFormat = "MMM dd, yyyy hh:mm tt"
        oView.TimeLine.Position(New AGVBE.DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day, System.DateTime.Now.Hour, System.DateTime.Now.Minute, 58))

        ActiveGanttVBECtl1.CurrentView = "MSI"

        Dim i As Integer
        ActiveGanttVBECtl1.Columns.Add("", "C1", 125, "")
        For i = 1 To 10
            Dim oRow As clsRow
            oRow = ActiveGanttVBECtl1.Rows.Add("K" & i.ToString, "K" & i.ToString(), True, True, "")
        Next
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectMove(sender As System.Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectMove
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBECtl1.Tasks.Item(e.Index.ToString())
            sText = ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectSize(sender As System.Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectSize
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBECtl1.Tasks.Item(e.Index.ToString())
            sText = ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_ObjectAdded(sender As System.Object, e As AGVBE.ObjectAddedEventArgs) Handles ActiveGanttVBECtl1.ObjectAdded
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            Dim oTask As clsTask
            Dim sText As String
            oTask = ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString())
            sText = ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_MILLISECOND, oTask.StartDate, oTask.EndDate).ToString()
            oTask.Text = sText & "ms"
        End If
    End Sub

    Private Sub cmdBack_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub

End Class
