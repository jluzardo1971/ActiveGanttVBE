Partial Public Class fLoad
    Inherits ChildWindow

    Friend mp_oFileList As List(Of CON_File)
    Friend sFileName As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub fLoad_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        drpFile.ItemsSource = mp_oFileList
        drpFile.DisplayMemberPath = "sDescription"
        drpFile.SelectedValuePath = "sFileName"
    End Sub

    Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdOK.Click
        sFileName = System.Convert.ToString(drpFile.SelectedValue)
        If sFileName.Length = 0 Then
            Return
        End If
        Me.DialogResult = True
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles cmdCancel.Click
        Me.DialogResult = False
    End Sub


End Class
