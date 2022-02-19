Imports AGVBE

Partial Public Class fFastLoading
    Inherits Page

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub cmdBack_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub

    Private Sub Page_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim i As Integer
        ActiveGanttVBECtl1.Columns.Add("Tasks", "", 200, "")
        ActiveGanttVBECtl1.TreeviewColumnIndex = 1
        ActiveGanttVBECtl1.Rows.BeginLoad(False)
        ActiveGanttVBECtl1.Tasks.BeginLoad(False)
        Dim lCurrentDepth As Integer = 0
        For i = 0 To 5000
            Dim oRow As clsRow
            Dim oTask As clsTask
            oRow = ActiveGanttVBECtl1.Rows.Load("K" & i.ToString)
            oTask = ActiveGanttVBECtl1.Tasks.Load("K" & i.ToString(), "K" & i.ToString)
            oRow.Text = "Task K" & i.ToString()
            oRow.MergeCells = True
            oRow.Node.Depth = lCurrentDepth
            oTask.Text = "K" & i.ToString()
            oTask.StartDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(0, 5), AGVBE.DateTime.Now)
            oTask.EndDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(2, 7), oTask.StartDate)
            lCurrentDepth = lCurrentDepth + GetRnd(-1, 2)
            If lCurrentDepth < 0 Then
                lCurrentDepth = 0
            End If
        Next
        ActiveGanttVBECtl1.Tasks.EndLoad()
        ActiveGanttVBECtl1.Rows.EndLoad()
        ActiveGanttVBECtl1.Rows.BeginLoad(True)
        ActiveGanttVBECtl1.Tasks.BeginLoad(True)
        For i = 5001 To 10000
            Dim oRow As clsRow
            Dim oTask As clsTask
            oRow = ActiveGanttVBECtl1.Rows.Load("KL" & i.ToString)
            oTask = ActiveGanttVBECtl1.Tasks.Load("KL" & i.ToString, "KL" & i.ToString)
            oRow.Text = "Task KL" & i.ToString()
            oRow.MergeCells = True
            oTask.Text = "KL" & i.ToString()
            oTask.StartDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(0, 5), AGVBE.DateTime.Now)
            oTask.EndDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, GetRnd(2, 7), oTask.StartDate)
        Next
        ActiveGanttVBECtl1.Tasks.EndLoad()
        ActiveGanttVBECtl1.Rows.EndLoad()
        ActiveGanttVBECtl1.Redraw()
    End Sub
End Class
