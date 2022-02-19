Partial Public Class MainPage
    Inherits UserControl



    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub MainPage_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ActiveGanttVBECtl1.Columns.Add("Tasks", "", 125, "DS_COLUMN")
        Dim i As Integer
        For i = 1 To 100
            ActiveGanttVBECtl1.Rows.Add("K" & i.ToString())
        Next

    End Sub
End Class
