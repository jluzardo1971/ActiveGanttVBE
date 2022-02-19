Imports AGVBE

Partial Public Class fSortRows
    Inherits Page

    Private mp_bDescending As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Page_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Dim i As Integer
        ActiveGanttVBECtl1.Columns.Add("", "C1", 125, "")
        For i = 1 To 10
            Dim si As String
            si = i.ToString
            While si.Length < 2
                si = "0" & si
            End While
            ActiveGanttVBECtl1.Rows.Add("K" & si, "K" & si, True, True, "")
        Next
    End Sub

    Private Sub cmdSortRows_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdSortRows.Click
        mp_bDescending = Not mp_bDescending
        ActiveGanttVBECtl1.Rows.SortRows("Text", mp_bDescending, E_SORTTYPE.ES_STRING, -1, -1)
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdBack_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub


End Class
