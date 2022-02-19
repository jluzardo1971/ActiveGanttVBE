Partial Public Class fYesNoMsgBox
    Inherits ChildWindow

    Private mp_sType As String

    Public Property Prompt() As String
        Get
            Return txtPrompt.Text
        End Get
        Set(ByVal value As String)
            txtPrompt.Text = value
        End Set
    End Property

    Public Property Type() As String
        Get
            Return mp_sType
        End Get
        Set(ByVal value As String)
            mp_sType = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
        'mp_sType = sType
    End Sub

    Private Sub YesButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles YesButton.Click
        Me.DialogResult = True
    End Sub

    Private Sub NoButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles NoButton.Click
        Me.DialogResult = False
    End Sub

End Class
