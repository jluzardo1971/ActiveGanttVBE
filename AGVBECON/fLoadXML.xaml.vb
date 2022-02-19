Imports AGVBECON.Web
Imports System.ServiceModel.DomainServices.Client

Partial Public Class fLoadXML
    Inherits Page

    Private oServiceContext As ActiveGanttXMLContext = New ActiveGanttXMLContext()

    Friend WithEvents invkGetXML As InvokeOperation(Of String)
    Friend WithEvents invkGetFileList As InvokeOperation(Of String)
    Friend WithEvents invkSetXML As InvokeOperation

    Private mp_oFileList As List(Of CON_File)
    Private mp_fLoad As fLoad

#Region "Constructors"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "Page Loaded"

    Private Sub fLoadXML_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        invkGetFileList = oServiceContext.GetFileList()
        mp_fLoad = New fLoad()
        AddHandler mp_fLoad.Closed, AddressOf mp_fLoad_Closed
    End Sub

    Private Sub invkGetFileList_Completed(sender As Object, e As System.EventArgs) Handles invkGetFileList.Completed
        Dim sFileList As String = invkGetFileList.Value
        If sFileList.Length > 0 Then
            Dim aFileList As String()
            Dim i As Integer
            aFileList = sFileList.Split(System.Convert.ToChar("|"))
            mp_oFileList = New List(Of CON_File)
            For i = 0 To aFileList.Length - 1
                Dim oFile As New CON_File()
                oFile.sDescription = aFileList(i)
                oFile.sFileName = aFileList(i)
                mp_oFileList.Add(oFile)
            Next
        End If
    End Sub

#End Region

#Region "Toolbar Buttons"

#Region "cmdLoadXML"

    Private Sub cmdLoadXML_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        mp_fLoad.Title = "Load XML File"
        mp_fLoad.mp_oFileList = mp_oFileList
        mp_fLoad.Show()
    End Sub

    Private Sub mp_fLoad_Closed()
        If mp_fLoad.DialogResult = True Then
            If mp_fLoad.sFileName.Length > 0 Then
                invkGetXML = oServiceContext.GetXML(mp_fLoad.sFileName)
            End If
        End If
    End Sub

    Private Sub invkGetXML_Completed(sender As Object, e As System.EventArgs) Handles invkGetXML.Completed
        Dim sXML As String = invkGetXML.Value
        ActiveGanttVBECtl1.SetXML(sXML)
        ActiveGanttVBECtl1.Redraw()
    End Sub

#End Region

#Region "cmdSaveXML"

    Private Sub cmdSaveXML_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        Dim sXML As String = ""
        sXML = ActiveGanttVBECtl1.GetXML()
        If sXML.Length > 0 Then
            Dim oFile As CON_File = New CON_File()
            oFile.sDescription = "TestXML1.xml"
            oFile.sFileName = "TestXML1.xml"
            mp_oFileList.Add(oFile)
            invkSetXML = oServiceContext.SetXML(sXML, "Test1.xml")
            cmdSaveXML.IsEnabled = False
        End If
    End Sub

    Private Sub invkSetXML_Completed(sender As Object, e As System.EventArgs) Handles invkSetXML.Completed
        cmdSaveXML.IsEnabled = True
    End Sub

#End Region

    Private Sub cmdBack_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub

#End Region

End Class


