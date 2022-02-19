﻿Imports AGVBECON.Web
Imports System.ServiceModel.DomainServices.Client
Imports AGVBE

Partial Public Class fMSProject11
    Inherits Page

    Private oServiceContext As MSP2003Context = New MSP2003Context()

    Friend WithEvents invkGetXML As InvokeOperation(Of String)
    Friend WithEvents invkGetFileList As InvokeOperation(Of String)

    Private mp_oFileList As List(Of CON_File)
    Private mp_fLoad As fLoad

    Private Const mp_sFontName As String = "Tahoma"
    Private oMP11 As MSP2003.MP11

#Region "Constructors"

    Public Sub New()
        InitializeComponent()
    End Sub

#End Region

#Region "Page Loaded"

    Private Sub fMSProject11_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        invkGetFileList = oServiceContext.GetFileList()
        mp_fLoad = New fLoad()
        AddHandler mp_fLoad.Closed, AddressOf mp_fLoad_Closed

        Me.Title = "The Source Code Store - ActiveGantt Scheduler Control Version " & ActiveGanttVBECtl1.Version & " - Microsoft Project 2003 integration using XML Files and the MSP2003 Integration Library"
        InitializeAG()
        ActiveGanttVBECtl1.Redraw()
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

#Region "ActiveGantt Event Handlers"

    Private Sub ActiveGanttVBECtl1_CustomTierDraw(sender As Object, e As AGVBE.CustomTierDrawEventArgs) Handles ActiveGanttVBECtl1.CustomTierDraw
        If e.TierPosition = E_TIERPOSITION.SP_UPPER Then
            e.StyleIndex = "TimeLineTiers"
            If System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) <= 4 Then
                e.Text = e.StartDate.Year() & " Q" & e.StartDate.Quarter()
            Else
                e.Text = e.StartDate.ToString("MMMM, yyyy")
            End If
        ElseIf e.TierPosition = E_TIERPOSITION.SP_LOWER Then
            e.StyleIndex = "TimeLineTiers"
            If System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) <= 4 Then
                e.Text = e.StartDate.ToString("MMM")
            Else
                e.Text = e.StartDate.ToString("ddd")
            End If
        End If
    End Sub

#End Region

#Region "Functions"

    Private Sub InitializeAG()

        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing

        oStyle = ActiveGanttVBECtl1.Styles.Add("TimeLineTiers")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED
        oStyle.BorderColor = Colors.DarkGray
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE

        oStyle = ActiveGanttVBECtl1.Styles.Add("TaskStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Blue
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH
        oStyle.HatchBackColor = Colors.White
        oStyle.HatchForeColor = Colors.Blue
        oStyle.HatchStyle = GRE_HATCHSTYLE.HS_PERCENT50
        oStyle.PredecessorStyle.LineColor = Colors.Black
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.MilestoneStyle.FillColor = Colors.Blue
        oStyle.MilestoneStyle.BorderColor = Colors.Blue
        oStyle.PredecessorStyle.XOffset = 4
        oStyle.PredecessorStyle.YOffset = 4

        oStyle = ActiveGanttVBECtl1.Styles.Add("SummaryStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextPlacement = E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 10
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndFillColor = Colors.Green
        oStyle.TaskStyle.EndBorderColor = Colors.Green
        oStyle.TaskStyle.StartFillColor = Colors.Green
        oStyle.TaskStyle.StartBorderColor = Colors.Green
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBECtl1.Styles.Add("CellStyleKeyColumn")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 128, 128, 128)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        ActiveGanttVBECtl1.AllowRowMove = True
        ActiveGanttVBECtl1.AllowRowSize = True
        ActiveGanttVBECtl1.AddMode = E_ADDMODE.AT_BOTH
        ActiveGanttVBECtl1.Splitter.Position = 285
        ActiveGanttVBECtl1.Treeview.Images = True
        ActiveGanttVBECtl1.Treeview.CheckBoxes = True
        ActiveGanttVBECtl1.Treeview.FullColumnSelect = True
        ActiveGanttVBECtl1.Treeview.TreeLines = False
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.TimerInterval = 50

        Dim oColumn As clsColumn

        oColumn = ActiveGanttVBECtl1.Columns.Add("ID", "", 30, "")
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBECtl1.Columns.Add("Task Name", "", 255, "")
        oColumn.AllowTextEdit = True

        ActiveGanttVBECtl1.TreeviewColumnIndex = 2
        ActiveGanttVBECtl1.Splitter.Position = 285

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 24, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBE.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = False

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 12, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBE.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 12
        oView.TimeLine.TimeLineScrollBar.LargeChange = 240
        oView.TimeLine.TimeLineScrollBar.Max = 2000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 6, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBE.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 6
        oView.TimeLine.TimeLineScrollBar.LargeChange = 480
        oView.TimeLine.TimeLineScrollBar.Max = 4000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 3, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBE.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 3
        oView.TimeLine.TimeLineScrollBar.LargeChange = 960
        oView.TimeLine.TimeLineScrollBar.Max = 8000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 1, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_DAY
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = AGVBE.DateTime.Now
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 48
        oView.TimeLine.TimeLineScrollBar.LargeChange = 2880
        oView.TimeLine.TimeLineScrollBar.Max = 24000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.ClientArea.DetectConflicts = False

        ActiveGanttVBECtl1.CurrentView = "5"

    End Sub

    Private Sub AGSetStartDate(ByVal dtStart As AGVBE.DateTime)
        Dim i As Integer
        For i = 1 To ActiveGanttVBECtl1.Views.Count
            ActiveGanttVBECtl1.Views.Item(i).TimeLine.TimeLineScrollBar.StartDate = dtStart
        Next
    End Sub

    Private Sub MP11_To_AG()
        Dim oAGTask As clsTask
        Dim oAGRow As clsRow
        Dim oMPTask As MSP2003.Task
        Dim dtStartDate As AGVBE.DateTime = AGVBE.DateTime.Now
        Dim i As Integer
        Dim j As Integer
        '// Load Project Tasks
        For i = 1 To oMP11.oTasks.Count
            oMPTask = oMP11.oTasks.Item(i)
            oAGRow = ActiveGanttVBECtl1.Rows.Add("K" & oMPTask.lUID.ToString())
            oAGRow.Cells.Item("1").Text = oMPTask.lUID.ToString()
            oAGRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
            oAGRow.Height = 20
            oAGTask = ActiveGanttVBECtl1.Tasks.Add("", "K" & oMPTask.lUID.ToString(), FromDate(oMPTask.dtStart), FromDate(oMPTask.dtFinish))
            oAGTask.Key = "K" & oMPTask.lUID.ToString()
            oAGTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
            oAGTask.AllowTextEdit = True
            If FromDate(oMPTask.dtStart) < dtStartDate Then
                dtStartDate = FromDate(oMPTask.dtStart)
            End If
            If oAGTask.StartDate = oAGTask.EndDate Then
                oAGTask.Text = oAGTask.StartDate.ToString("M/d")
            End If
            oAGRow.Node.Depth = oMPTask.lOutlineLevel
            oAGRow.Node.Text = oMPTask.sName
            oAGRow.Node.AllowTextEdit = True
            If oMPTask.sNotes.Length > 0 Then
                oAGRow.Node.Image = GetImage("Note.png")
                oAGRow.Node.ImageVisible = True
            End If
        Next
        ActiveGanttVBECtl1.Rows.UpdateTree()
        '// Indent & set Predecessors
        For i = 1 To oMP11.oTasks.Count
            oMPTask = oMP11.oTasks.Item(i)
            oAGRow = ActiveGanttVBECtl1.Rows.Item(i)
            oAGTask = ActiveGanttVBECtl1.Tasks.Item(i)
            If oAGRow.Node.Children > 0 Then
                oAGTask.StyleIndex = "SummaryStyle"
            Else
                oAGTask.StyleIndex = "TaskStyle"
            End If
            For j = 1 To oMPTask.oPredecessorLink_C.Count
                Dim oMPPredecessor As MSP2003.TaskPredecessorLink
                oMPPredecessor = oMPTask.oPredecessorLink_C.Item(j)
                ActiveGanttVBECtl1.Predecessors.Add("K" & oMPTask.lUID.ToString(), "K" & oMPPredecessor.lPredecessorUID.ToString(), GetAGPredecessorType(oMPPredecessor.yType), "", "TaskStyle")
            Next
        Next
        'Assignments
        For i = 1 To oMP11.oAssignments.Count
            Dim oAssignment As MSP2003.Assignment
            oAssignment = oMP11.oAssignments.Item(i)
            oAGTask = ActiveGanttVBECtl1.Tasks.Item("K" & oAssignment.lTaskUID)
            If oAGTask.StartDate <> oAGTask.EndDate Then
                If oAssignment.lResourceUID > 0 Then
                    If oAGTask.Text.Length = 0 Then
                        oAGTask.Text = oMP11.oResources.Item("K" & oAssignment.lResourceUID).sName
                    Else
                        oAGTask.Text = oAGTask.Text & ", " & oMP11.oResources.Item("K" & oAssignment.lResourceUID).sName
                    End If
                End If
            End If
        Next
        dtStartDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_DAY, -3, dtStartDate)
        AGSetStartDate(dtStartDate)
    End Sub

    Private Function GetAGPredecessorType(ByVal MPPredecessorType As MSP2003.E_TYPE_3) As AGVBE.E_CONSTRAINTTYPE
        Select Case MPPredecessorType
            Case MSP2003.E_TYPE_3.T_3_FF
                Return AGVBE.E_CONSTRAINTTYPE.PCT_END_TO_END
            Case MSP2003.E_TYPE_3.T_3_FS
                Return AGVBE.E_CONSTRAINTTYPE.PCT_END_TO_START
            Case MSP2003.E_TYPE_3.T_3_SF
                Return AGVBE.E_CONSTRAINTTYPE.PCT_START_TO_END
            Case MSP2003.E_TYPE_3.T_3_SS
                Return AGVBE.E_CONSTRAINTTYPE.PCT_START_TO_START
        End Select
        Return AGVBE.E_CONSTRAINTTYPE.PCT_END_TO_START
    End Function

    Private Function GetImage(ByVal sImage As String) As Image
        Dim oReturnImage = New Image
        Dim oURI As System.Uri = New System.Uri("AGVBECON;component/Images/MSP/" & sImage, UriKind.Relative)
        Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
        Dim oSRI As System.Windows.Resources.StreamResourceInfo = Application.GetResourceStream(oURI)
        oBitmap.SetSource(oSRI.Stream)
        oReturnImage.Height = 16
        oReturnImage.Width = 16
        oReturnImage.Source = oBitmap
        Return oReturnImage
    End Function

#End Region

#Region "Toolbar Buttons"

#Region "cmdLoad"

    Private Sub cmdLoadXML_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        mp_fLoad.Title = "Load MS-Project 2003 XML file"
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
        sXML = g_RemoveXMLNameSpaces(sXML)
        Me.Cursor = Cursors.Wait
        ActiveGanttVBECtl1.Clear()
        oMP11 = New MSP2003.MP11()
        oMP11.SetXML(sXML)
        Me.Cursor = Cursors.Wait
        InitializeAG()
        MP11_To_AG()
        ActiveGanttVBECtl1.Redraw()
        ActiveGanttVBECtl1.VerticalScrollBar.LargeChange = ActiveGanttVBECtl1.CurrentViewObject.ClientArea.LastVisibleRow - ActiveGanttVBECtl1.CurrentViewObject.ClientArea.FirstVisibleRow
        ActiveGanttVBECtl1.Redraw()
        Me.Cursor = Cursors.Arrow
    End Sub

#End Region

    Private Sub cmdBack_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub

    Private Sub cmdZoomin_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdZoomin.Click
        If (System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) < ActiveGanttVBECtl1.Views.Count) Then
            ActiveGanttVBECtl1.CurrentView = System.Convert.ToString(System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) + 1)
            ActiveGanttVBECtl1.Redraw()
        End If
    End Sub

    Private Sub cmdZoomout_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdZoomout.Click
        If (System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) > 1) Then
            ActiveGanttVBECtl1.CurrentView = System.Convert.ToString(System.Convert.ToInt32(ActiveGanttVBECtl1.CurrentView) - 1)
            ActiveGanttVBECtl1.Redraw()
        End If
    End Sub

#End Region

End Class
