Imports System.Windows.Media.Imaging

Partial Public Class fMain
    Inherits Page

    Dim mp_oParentNode As TreeViewItem

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub fMain_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        AddTitleNode("AGEX", "ActiveGantt Examples:", 4, 5)
        AddNode("AGEX", "GanttCharts", "Gantt Charts:", 4, 5)

        AddNode("GanttCharts", "WBS", "Work Breakdown Structure (WBS) Project Management Examples:", 4, 5)
        AddNode("WBS", "WBSProject", "No data source (32bit and 64bit compatible)", 2, 2)
        'AddNode("WBS", "WBSProjectXML", "XML data source (32bit and 64bit compatible)", 2, 2)
        'AddNode("WBS", "WBSProjectAccess", "Microsoft Access data source (32bit compatible only)", 2, 2)

        AddNode("GanttCharts", "MSPI", "Microsoft Project Integration Examples (32bit and 64bit compatible):", 4, 5)
        AddNode("MSPI", "Project2003", "Demonstrates how ActiveGantt integrates with MS Project 2003 (using XML Files and the MSP2003 Integration Library)", 2, 2)
        AddNode("MSPI", "Project2007", "Demonstrates how ActiveGantt integrates with MS Project 2007 (using XML Files and the MSP2007 Integration Library)", 2, 2)
        AddNode("MSPI", "Project2010", "Demonstrates how ActiveGantt integrates with MS Project 2010 (using XML Files and the MSP2010 Integration Library)", 2, 2)

        AddNode("AGEX", "Schedules", "Schedules and Rosters:", 4, 5)

        AddNode("Schedules", "VRFC", "Vehicle Rental/Fleet Control Roster Examples:", 4, 5)
        AddNode("VRFC", "CarRental", "No data source (32bit and 64bit compatible)", 2, 2)
        'AddNode("VRFC", "CarRentalXML", "XML data source (32bit and 64bit compatible)", 2, 2)
        'AddNode("VRFC", "CarRentalAccess", "Microsoft Access data source (32bit compatible only)", 2, 2)

        AddNode("AGEX", "OTHER", "Other examples:", 4, 5)
        AddNode("OTHER", "FastLoad", "Fast Loading of Row and Task objects", 2, 2)
        AddNode("OTHER", "CustomDrawing", "Custom Drawing", 2, 2)
        AddNode("OTHER", "SortRows", "Sort Rows", 2, 2)
        AddNode("OTHER", "MillisecondInterval", "5 Millisecond Interval View", 2, 2)

        AddNode("OTHER", "TimeBlocks", "TimeBlocks and Duration Tasks:", 4, 5)
        AddNode("TimeBlocks", "RCT_DAY", "Daily Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_WEEK", "Weekly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_MONTH", "Monthly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "RCT_YEAR", "Yearly Recurrent TimeBlocks", 2, 2)
        AddNode("TimeBlocks", "DurationTasks", "Duration Tasks (can skip over non-working TimeBlock intervals)", 2, 2)

        AddTitleNode("HLP", "Help", 7, 7)
        AddNode("HLP", "GS-VBE", "How to create a simple Silverlight 4 application using the ActiveGanttVBE component", 3, 3)
        AddNode("HLP", "OnlineDocumentation", "ActiveGanttVBE Online Documentation", 6, 6)
        AddNode("HLP", "BugReport", "Submit a Bug Report", 3, 3)
        AddNode("HLP", "Request", "Request Further Explanations, Code Samples and Submit Technical Queries", 6, 6)

        AddTitleNode("SCS", "The Source Code Store LLC - Website (http://www.sourcecodestore.com/)", 3, 3)
        AddNode("SCS", "OnlineStore", "Online Store - Purchase ActiveGantt Online", 3, 3)
        AddNode("SCS", "ContactUs", "Contact Us (use this form for non technical queries only)", 3, 3)

        Dim oNode As TreeViewItem
        oNode = FindNode("OTHER")
        oNode.IsExpanded = False

    End Sub

    Private Sub AddTitleNode(ByVal sKey As String, ByVal sText As String, ByVal ImageIndex As Integer, ByVal SelectedImageIndex As Integer)
        Dim oNode As New TreeViewItem()
        oNode.Name = sKey
        oNode.Header = GetStackPanel(sText, ImageIndex)
        oNode.Tag = sKey
        oNode.IsExpanded = True
        TreeView1.Items.Add(oNode)
        mp_oParentNode = oNode
    End Sub

    Private Sub AddNode(ByVal sParentKey As String, ByVal sKey As String, ByVal sText As String, ByVal ImageIndex As Integer, ByVal SelectedImageIndex As Integer)
        Dim oNode As New TreeViewItem()
        Dim oParentNode As TreeViewItem
        oNode.Name = sKey
        oNode.Header = GetStackPanel(sText, ImageIndex)
        oNode.Tag = sKey
        oNode.IsExpanded = True
        oParentNode = FindNode(sParentKey)
        oParentNode.Items.Add(oNode)
    End Sub

    Private Function FindNode(ByVal sName As String) As TreeViewItem
        Dim i As Integer
        Dim oReturnTreeViewItem As TreeViewItem = Nothing
        For i = 0 To TreeView1.Items.Count - 1
            Dim oTreeViewItem As TreeViewItem = DirectCast(TreeView1.Items(i), TreeViewItem)
            oReturnTreeViewItem = FindNode_Intermediate(oTreeViewItem, sName)
            If Not (oReturnTreeViewItem Is Nothing) Then
                Return oReturnTreeViewItem
            End If
            oReturnTreeViewItem = FindNode_Final(oTreeViewItem, sName)
            If Not oReturnTreeViewItem Is Nothing Then
                Return oReturnTreeViewItem
            End If
        Next
        Return oReturnTreeViewItem
    End Function

    Private Function FindNode_Intermediate(ByRef oTreeViewItem As TreeViewItem, ByVal sName As String) As TreeViewItem
        Dim i As Integer
        Dim oReturnTreeViewItem As TreeViewItem = Nothing
        For i = 0 To oTreeViewItem.Items.Count - 1
            Dim oChildTreeViewItem As TreeViewItem = DirectCast(oTreeViewItem.Items(i), TreeViewItem)
            oReturnTreeViewItem = FindNode_Intermediate(oChildTreeViewItem, sName)
            If Not oReturnTreeViewItem Is Nothing Then
                Return oReturnTreeViewItem
            End If
        Next
        oReturnTreeViewItem = FindNode_Final(oTreeViewItem, sName)
        Return oReturnTreeViewItem
    End Function

    Private Function FindNode_Final(ByRef oTreeViewItem As TreeViewItem, ByVal sName As String) As TreeViewItem
        If oTreeViewItem.Name = sName Then
            Return oTreeViewItem
        End If
        Return Nothing
    End Function

    Private Function GetStackPanel(ByVal sText As String, ByVal ImageIndex As Integer) As StackPanel
        Dim oStackPanel As New StackPanel
        Dim oImage As New Image
        Dim oTextBlock As New TextBlock
        oImage.Source = GetImage(ImageIndex)
        oTextBlock.Text = " " & sText
        oStackPanel.Orientation = Orientation.Horizontal
        oStackPanel.Children.Add(oImage)
        oStackPanel.Children.Add(oTextBlock)
        Return oStackPanel
    End Function

    Private Function GetImage(ByVal lImageIndex As Integer) As BitmapImage
        Dim sImage As String = ""
        Select Case lImageIndex
            Case 4
                sImage = "folderopen.png"
            Case 2
                sImage = "ActiveGantt.png"
            Case 3
                sImage = "internet.png"
            Case 6
                sImage = "onlinedocumentation.png"
            Case 7
                sImage = "localCHMdocumentation.png"
        End Select
        Dim oURI As System.Uri = New System.Uri("AGVBECON;component/Images/" & sImage, UriKind.Relative)
        Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
        Dim oSRI As System.Windows.Resources.StreamResourceInfo = Application.GetResourceStream(oURI)
        oBitmap.SetSource(oSRI.Stream)
        Return oBitmap
    End Function


    Private Sub TreeView1_MouseLeftButtonUp(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles TreeView1.MouseLeftButtonUp
        Dim oSelectedTreeViewItem As TreeViewItem = DirectCast(TreeView1.SelectedItem, TreeViewItem)
        Dim sSelectedTag As String = DirectCast(oSelectedTreeViewItem.Tag, String)
        Dim oApp As App = DirectCast(Application.Current, App)
        Select Case sSelectedTag
            Case "WBSProject"
                Dim oForm As New fWBSProject()
                oForm.DataSource = E_DATASOURCETYPE.DST_NONE
                Me.Content = oForm
            Case "LoadXML"
                Dim oForm As New fLoadXML()
                Me.Content = oForm
            Case "WBSProjectXML"
                'Dim oForm As New fWBSProject()
                'oForm.DataSource = E_DATASOURCETYPE.DST_XML
                'Me.Content = oForm
            Case "WBSProjectAccess"
                'Dim oForm As New fWBSProject()
                'oForm.DataSource = E_DATASOURCETYPE.DST_ACCESS
                'Me.Content = oForm
            Case "Project2003"
                Dim oForm As New fMSProject11()
                Me.Content = oForm
            Case "Project2007"
                Dim oForm As New fMSProject12()
                Me.Content = oForm
            Case "Project2010"
                Dim oForm As New fMSProject14()
                Me.Content = oForm
            Case "CarRental"
                Dim oForm As New fCarRental()
                oForm.DataSource = E_DATASOURCETYPE.DST_NONE
                Me.Content = oForm
            Case "CarRentalXML"
                'Dim oForm As New fCarRental()
                'oForm.DataSource = E_DATASOURCETYPE.DST_XML
                'Me.Content = oForm
            Case "CarRentalAccess"
                'Dim oForm As New fCarRental()
                'oForm.DataSource = E_DATASOURCETYPE.DST_ACCESS
                'Me.Content = oForm
            Case "FastLoad"
                Dim oForm As New fFastLoading()
                Me.Content = oForm
            Case "CustomDrawing"
                Dim oForm As New fCustomDrawing()
                Me.Content = oForm
            Case "SortRows"
                Dim oForm As New fSortRows()
                Me.Content = oForm
            Case "MillisecondInterval"
                Dim oForm As New fMillisecondInterval()
                Me.Content = oForm
            Case "DurationTasks"
                Dim oForm As New fDurationTasks()
                Me.Content = oForm
            Case "RCT_DAY"
                Dim oForm As New fRCT_DAY()
                Me.Content = oForm
            Case "RCT_WEEK"
                Dim oForm As New fRCT_WEEK()
                Me.Content = oForm
            Case "RCT_MONTH"
                Dim oForm As New fRCT_MONTH()
                Me.Content = oForm
            Case "RCT_YEAR"
                Dim oForm As New fRCT_YEAR()
                Me.Content = oForm
            Case "SCS"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/"), "_blank")
            Case "GS-VBE"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Article.aspx?ID=18#Create"), "_blank")
            Case "OnlineStore"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/OnlineStore/"), "_blank")
            Case "OnlineDocumentation"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Documentation/DOCFrameset.aspx?PN=AG&PL=VBE"), "_blank")
            Case "BugReport"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Support/Report.aspx?T=1"), "_blank")
            Case "Request"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Support/Report.aspx?T=2"), "_blank")
            Case "ContactUs"
                System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/contactus.aspx"), "_blank")
        End Select
    End Sub
End Class
