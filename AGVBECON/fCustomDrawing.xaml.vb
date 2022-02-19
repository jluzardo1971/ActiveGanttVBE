Imports AGVBE

Partial Public Class fCustomDrawing
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
        ActiveGanttVBECtl1.Columns.Add("Column 1", "", 125, "")
        ActiveGanttVBECtl1.Columns.Add("Column 2", "", 100, "")
        For i = 1 To 10
            ActiveGanttVBECtl1.Rows.Add("K" & i.ToString(), "Row " & i.ToString() & " (Key: " & "K" & i.ToString() & ")", True, True, "")
        Next

        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(New AGVBE.DateTime(2011, 11, 21, 0, 0, 0))
        ActiveGanttVBECtl1.Tasks.Add("Task 1", "K1", New AGVBE.DateTime(2011, 11, 21, 0, 0, 0), New AGVBE.DateTime(2011, 11, 21, 3, 0, 0), "", "", "")
        ActiveGanttVBECtl1.Tasks.Add("Task 2", "K2", New AGVBE.DateTime(2011, 11, 21, 1, 0, 0), New AGVBE.DateTime(2011, 11, 21, 4, 0, 0), "", "", "")
        ActiveGanttVBECtl1.Tasks.Add("Task 3", "K3", New AGVBE.DateTime(2011, 11, 21, 2, 0, 0), New AGVBE.DateTime(2011, 11, 21, 5, 0, 0), "", "", "")

        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub ActiveGanttVBECtl1_Draw(sender As System.Object, e As AGVBE.DrawEventArgs) Handles ActiveGanttVBECtl1.Draw
        If e.EventTarget = E_EVENTTARGET.EVT_TASK Then
            If ActiveGanttVBECtl1.SelectedTaskIndex = e.ObjectIndex Then
                e.CustomDraw = True
                Dim oTask As clsTask
                oTask = ActiveGanttVBECtl1.Tasks.Item(e.ObjectIndex.ToString())
                Dim oFont As New Font("Arial", 10, E_FONTSIZEUNITS.FSU_POINTS, FontWeights.Normal)
                Dim oTextFlags As New clsTextFlags(ActiveGanttVBECtl1)
                oTextFlags.HorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER
                oTextFlags.VerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER
                Dim oImage As New Image()
                Dim oURI As System.Uri = New System.Uri("AGVBECON;component/Images/sampleimage.jpg", UriKind.Relative)
                Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
                Dim oSRI As System.Windows.Resources.StreamResourceInfo = Application.GetResourceStream(oURI)
                oBitmap.SetSource(oSRI.Stream)
                oImage.Width = 24
                oImage.Height = 24
                oImage.Source = oBitmap
                ActiveGanttVBECtl1.Drawing.PaintImage(oImage, oTask.Left + 40, oTask.Top + 10, oTask.Left + 64, oTask.Top + 34)
                ActiveGanttVBECtl1.Drawing.DrawLine(oTask.Left, System.Convert.ToInt32(((oTask.Bottom - oTask.Top) / 2) + oTask.Top), oTask.Right, System.Convert.ToInt32(((oTask.Bottom - oTask.Top) / 2) + oTask.Top), Colors.Green, GRE_LINEDRAWSTYLE.LDS_SOLID, 1)
                ActiveGanttVBECtl1.Drawing.DrawRectangle(oTask.Left, oTask.Top, oTask.Left + 10, oTask.Top + 10, Colors.Red, GRE_LINEDRAWSTYLE.LDS_SOLID, 1)
                ActiveGanttVBECtl1.Drawing.DrawBorder(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, Colors.Blue, GRE_LINEDRAWSTYLE.LDS_SOLID, 2)
                ActiveGanttVBECtl1.Drawing.DrawAlignedText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, oTask.Text & " Is Selected", GRE_HORIZONTALALIGNMENT.HAL_RIGHT, GRE_VERTICALALIGNMENT.VAL_BOTTOM, Colors.Blue, oFont)
                ActiveGanttVBECtl1.Drawing.DrawText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, "Draw Text", oTextFlags, Colors.Red, oFont)
            End If
        End If
    End Sub
End Class
