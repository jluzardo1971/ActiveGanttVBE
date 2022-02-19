Partial Public Class fSave
    Inherits ChildWindow

    Friend mp_oFileList As List(Of CON_File)
    Friend sFileName As String = ""
    Friend sSuggestedFileName As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub ChildWindow_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If sSuggestedFileName.Length > 0 Then
            Dim j As Integer = 0
            While True
                Dim i As Integer = 0
                For i = 0 To mp_oFileList.Count - 1
                    If j = 0 Then
                        If (sSuggestedFileName.ToLower() & ".xml") = mp_oFileList(i).sFileName.ToLower() Then
                            j = j + 1
                            Exit For
                        End If
                    Else
                        If (sSuggestedFileName.ToLower() & j.ToString() & ".xml") = mp_oFileList(i).sFileName.ToLower() Then
                            j = j + 1
                            Exit For
                        End If
                    End If
                Next
                Exit While
            End While
            If j = 0 Then
                txtFileName.Text = sSuggestedFileName & ".xml"
            Else
                txtFileName.Text = sSuggestedFileName & j.ToString() & ".xml"
            End If
        End If
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
        Me.DialogResult = False
    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdOK.Click
        Dim i As Integer
        If txtFileName.Text.Length = 0 Then
            Return
        End If
        If txtFileName.Text.ToLower().EndsWith(".xml") = False Then
            txtFileName.Text = txtFileName.Text + ".xml"
        End If
        For i = 0 To mp_oFileList.Count - 1
            If txtFileName.Text.ToLower() = mp_oFileList(i).sFileName.ToLower() Then
                Return
            End If
        Next
        sFileName = txtFileName.Text
        Me.DialogResult = True
    End Sub


End Class
