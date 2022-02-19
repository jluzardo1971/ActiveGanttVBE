Imports AGVBE
Imports System.Windows.Controls.Primitives
Imports System.ServiceModel.DomainServices.Client
Imports AGVBECON.Web

Partial Public Class fCarRental
    Inherits Page

    Public Enum HPE_ADDMODE
        AM_RESERVATION = 0
        AM_RENTAL = 1
        AM_MAINTENANCE = 2
    End Enum


    Private mp_yAddMode As HPE_ADDMODE = HPE_ADDMODE.AM_RENTAL
    Private mp_sAddModeStyleIndex As String
    Private mp_lZoom As Integer
    Private mp_sEditRowKey As String
    Private mp_sEditTaskKey As String

    Private Const mp_sFontName As String = "Tahoma"
    Friend mp_yDataSourceType As E_DATASOURCETYPE

    Friend mp_o_AG_CR_Rows As List(Of AG_CR_Row)
    Friend mp_o_AG_CR_Rentals As List(Of AG_CR_Rental)

    Friend mp_o_AG_CR_Car_Types As List(Of AG_CR_Car_Type)
    Friend mp_o_AG_CR_US_States As List(Of AG_CR_US_State)
    Friend mp_o_AG_CR_ACRISS_Codes As List(Of AG_CR_ACRISS_Code)
    Friend mp_o_AG_CR_ACRISS_Codes_1 As List(Of AG_CR_ACRISS_Code)
    Friend mp_o_AG_CR_ACRISS_Codes_2 As List(Of AG_CR_ACRISS_Code)
    Friend mp_o_AG_CR_ACRISS_Codes_3 As List(Of AG_CR_ACRISS_Code)
    Friend mp_o_AG_CR_ACRISS_Codes_4 As List(Of AG_CR_ACRISS_Code)
    Friend mp_o_AG_CR_Taxes_Surcharges_Options As List(Of AG_CR_Tax_Surcharge_Option)

    Private mp_fCarRentalBranch As fCarRentalBranch
    Private mp_fCarRentalReservation As fCarRentalReservation
    Private mp_fCarRentalVehicle As fCarRentalVehicle

    Private mp_fYesNoMsgBox As fYesNoMsgBox


    Private mp_oCursorPosition As Point

    Private mp_oTaskPopUp As PopUp
    Private mp_oTaskStackPanel As StackPanel
    Private mp_mnuEditTask As TextBlock
    Private mp_mnuDeleteTask As TextBlock
    Private mp_mnuConvertToRental As TextBlock

    Private mp_oRowPopUp As Popup
    Private mp_oRowStackPanel As StackPanel
    Private mp_mnuEditRow As TextBlock
    Private mp_mnuDeleteRow As TextBlock

    Private WithEvents invkGetFileList As InvokeOperation(Of String)
    Private WithEvents invkSetXML As InvokeOperation
    Private oServiceContext As New ActiveGanttXMLContext()
    Private mp_oFileList As List(Of CON_File)
    Private WithEvents mp_fSave As fSave

#Region "Constructors"

    Public Sub New()
        InitializeComponent()
        mp_o_AG_CR_Rows = New List(Of AG_CR_Row)
        mp_o_AG_CR_Rentals = New List(Of AG_CR_Rental)
        mp_o_AG_CR_Car_Types = New List(Of AG_CR_Car_Type)
        mp_o_AG_CR_US_States = New List(Of AG_CR_US_State)
        mp_o_AG_CR_ACRISS_Codes = New List(Of AG_CR_ACRISS_Code)
        mp_o_AG_CR_ACRISS_Codes_1 = New List(Of AG_CR_ACRISS_Code)
        mp_o_AG_CR_ACRISS_Codes_2 = New List(Of AG_CR_ACRISS_Code)
        mp_o_AG_CR_ACRISS_Codes_3 = New List(Of AG_CR_ACRISS_Code)
        mp_o_AG_CR_ACRISS_Codes_4 = New List(Of AG_CR_ACRISS_Code)
        mp_o_AG_CR_Taxes_Surcharges_Options = New List(Of AG_CR_Tax_Surcharge_Option)
        mp_fCarRentalBranch = New fCarRentalBranch(Me)
        mp_fCarRentalReservation = New fCarRentalReservation(Me)
        mp_fCarRentalVehicle = New fCarRentalVehicle(Me)
        mp_fYesNoMsgBox = New fYesNoMsgBox()
        AddHandler mp_fYesNoMsgBox.Closed, AddressOf YesNoMsgBox_Closed
    End Sub

#End Region

#Region "Page Loaded"

    Private Sub fCarRental_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        InitRowContextMenu()
        InitTaskContextMenu()

        invkGetFileList = oServiceContext.GetFileList()


        If mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            'g_VerifyWriteAccess("CR_XML")
            'XML_Load_Car_Types()
            'XML_Load_US_States()
            'XML_Load_ACRISS_Codes()
            'XML_Load_Taxes_Surcharges_Options()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_Load_Car_Types()
            NoDataSource_Load_US_States()
            NoDataSource_Load_ACRISS_Codes(mp_o_AG_CR_ACRISS_Codes, 0)
            NoDataSource_Load_ACRISS_Codes(mp_o_AG_CR_ACRISS_Codes_1, 1)
            NoDataSource_Load_ACRISS_Codes(mp_o_AG_CR_ACRISS_Codes_2, 2)
            NoDataSource_Load_ACRISS_Codes(mp_o_AG_CR_ACRISS_Codes_3, 3)
            NoDataSource_Load_ACRISS_Codes(mp_o_AG_CR_ACRISS_Codes_4, 4)
            NoDataSource_Load_Taxes_Surcharges_Options()
        End If


        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing
        Dim oTimeBlock As clsTimeBlock = Nothing

        oStyle = ActiveGanttVBECtl1.Styles.Add("ScrollBar")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 158, 168)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ArrowButtons")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 158, 168)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ThumbButton")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBECtl1.Styles.Add("SplitterStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 220, 220, 220)

        oStyle = ActiveGanttVBECtl1.Styles.Add("Columns")
        oStyle.Font = New Font(mp_sFontName, 8, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Top = False
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM

        oStyle = ActiveGanttVBECtl1.Styles.Add("TimeLine")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.CustomBorderStyle.Left = True
        oStyle.CustomBorderStyle.Top = True
        oStyle.CustomBorderStyle.Right = False
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM

        oStyle = ActiveGanttVBECtl1.Styles.Add("TimeLineVA")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 148, 164, 189)
        oStyle.EndGradientColor = Color.FromArgb(255, 178, 199, 228)
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.DrawTextInVisibleArea = True

        oStyle = ActiveGanttVBECtl1.Styles.Add("Branch")
        oStyle.Font = New Font(mp_sFontName, 9, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.TextYMargin = 5
        oStyle.ForeColor = Colors.White
        oStyle.BorderColor = Colors.Black
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.ImageAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.ImageAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
        oStyle.ImageXMargin = 5
        oStyle.ImageYMargin = 5
        oStyle.UseMask = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("BranchCA")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)
        oStyle.ForeColor = Colors.White

        oStyle = ActiveGanttVBECtl1.Styles.Add("Weekend")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 133, 143, 154)
        oStyle.EndGradientColor = Color.FromArgb(255, 172, 183, 194)

        oStyle = ActiveGanttVBECtl1.Styles.Add("Reservation")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 109, 122, 136)
        oStyle.EndGradientColor = Color.FromArgb(255, 179, 199, 229)

        oStyle = ActiveGanttVBECtl1.Styles.Add("Rental")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 162, 78, 50)
        oStyle.EndGradientColor = Color.FromArgb(255, 215, 92, 54)

        oStyle = ActiveGanttVBECtl1.Styles.Add("Maintenance")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.ForeColor = Colors.White
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        oStyle.TextXMargin = 5
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 255, 77, 1)
        oStyle.EndGradientColor = Color.FromArgb(255, 244, 172, 43)

        ActiveGanttVBECtl1.ControlTag = "CarRental"
        ActiveGanttVBECtl1.Columns.Add("", "", 45, "Columns")
        ActiveGanttVBECtl1.Columns.Add("", "", 95, "Columns")
        ActiveGanttVBECtl1.Columns.Add("", "", 250, "Columns")

        ActiveGanttVBECtl1.Splitter.Position = 340
        ActiveGanttVBECtl1.Splitter.Type = E_SPLITTERTYPE.SA_STYLE
        ActiveGanttVBECtl1.Splitter.Width = 6
        ActiveGanttVBECtl1.Splitter.StyleIndex = "SplitterStyle"

        ActiveGanttVBECtl1.ScrollBarSeparator.StyleIndex = "ScrollBar"

        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButton"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButton"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButton"

        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButton"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButton"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButton"

        oTimeBlock = ActiveGanttVBECtl1.TimeBlocks.Add("")
        oTimeBlock.TimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING
        oTimeBlock.RecurringType = E_RECURRINGTYPE.RCT_WEEK
        oTimeBlock.BaseWeekDay = E_WEEKDAY.WD_FRIDAY
        oTimeBlock.BaseDate = New AGVBE.DateTime(2013, 1, 1, 0, 0, 0)
        oTimeBlock.DurationFactor = 48
        oTimeBlock.DurationInterval = E_INTERVAL.IL_HOUR
        oTimeBlock.StyleIndex = "Weekend"

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_MINUTE, 30, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.MiddleTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Interval = E_INTERVAL.IL_DAY
        oView.TimeLine.TierArea.MiddleTier.Factor = 1
        oView.TimeLine.TierArea.MiddleTier.Visible = True
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TierArea.LowerTier.Factor = 12
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TickMarkArea.StyleIndex = "TimeLine"
        oView.TimeLine.Style.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oView.TimeLine.Style.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oView.TimeLine.Style.BackColor = Colors.Black
        oView.ClientArea.Grid.VerticalLines = True
        oView.ClientArea.Grid.SnapToGrid = True
        ActiveGanttVBECtl1.CurrentView = oView.Index.ToString()
        Zoom = 5

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'Access_LoadRowsAndTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            'XML_LoadRowsAndTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_LoadRowsAndTasks()
        End If
        ActiveGanttVBECtl1.Rows.UpdateTree()

        Mode = HPE_ADDMODE.AM_RESERVATION

        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(New DateTime(2009, 6, 9))
    End Sub

    Private Sub invkGetFileList_Completed(sender As Object, e As System.EventArgs) Handles invkGetFileList.Completed
        Dim sFileList As String = invkGetFileList.Value
        If sFileList.Length > 0 Then
            Dim aFileList As String() = Nothing
            Dim i As Integer = 0
            aFileList = sFileList.Split(System.Convert.ToChar("|"))
            mp_oFileList = New List(Of CON_File)()
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
        If e.Interval = E_INTERVAL.IL_HOUR And e.Factor = 12 Then
            e.Text = e.StartDate.ToString("tt").ToUpper()
            e.StyleIndex = "TimeLine"
        End If
        If e.Interval = E_INTERVAL.IL_MONTH And e.Factor = 1 Then
            e.Text = e.StartDate.ToString("MMMM yyyy")
            e.StyleIndex = "TimeLineVA"
        End If
        If e.Interval = E_INTERVAL.IL_DAY And e.Factor = 1 Then
            e.Text = e.StartDate.ToString("ddd d")
            e.StyleIndex = "TimeLine"
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_ObjectAdded(sender As Object, e As AGVBE.ObjectAddedEventArgs) Handles ActiveGanttVBECtl1.ObjectAdded
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask = Nothing
                Dim lTaskID As Integer = 0
                oTask = ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString())
                oTask.StyleIndex = mp_sAddModeStyleIndex
                oTask.Tag = mp_yAddMode.ToString()
                If Mode = HPE_ADDMODE.AM_RESERVATION Then
                    mp_fCarRentalReservation.Mode = PRG_DIALOGMODE.DM_ADD
                    mp_fCarRentalReservation.mp_sTaskID = oTask.Key.Replace("K", "")
                    mp_fCarRentalReservation.Show()
                ElseIf Mode = HPE_ADDMODE.AM_RENTAL Then
                    mp_fCarRentalReservation.Mode = PRG_DIALOGMODE.DM_ADD
                    mp_fCarRentalReservation.mp_sTaskID = oTask.Key.Replace("K", "")
                    mp_fCarRentalReservation.Show()
                ElseIf Mode = HPE_ADDMODE.AM_MAINTENANCE Then
                    If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                        '//TODO
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                        '//TODO
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                        Dim oRental As AG_CR_Rental
                        For Each oRental In mp_o_AG_CR_Rentals
                            If oRental.lTaskID > lTaskID Then
                                lTaskID = oRental.lTaskID
                            End If
                        Next
                        lTaskID = lTaskID + 1
                        oRental = New AG_CR_Rental
                        oRental.lRowID = System.Convert.ToInt32(oTask.RowKey.Replace("K", ""))
                        oRental.yMode = 2
                        oRental.dtPickUp = oTask.StartDate.DateTimePart
                        oRental.dtReturn = oTask.EndDate.DateTimePart
                        oRental.bGPS = False
                        oRental.bFSO = False
                        oRental.bLDW = False
                        oRental.bPAI = False
                        oRental.bPEP = False
                        oRental.bALI = False
                        oRental.lTaskID = lTaskID
                        mp_o_AG_CR_Rentals.Add(oRental)
                        oTask.Key = "K" & lTaskID.ToString()
                        oTask.Text = "Scheduled Maintenance"
                        oTask.Tag = System.Convert.ToString(System.Convert.ToInt32(HPE_ADDMODE.AM_MAINTENANCE))
                    End If
                End If
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectMove(sender As Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectMove
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask = Nothing
                oTask = ActiveGanttVBECtl1.Tasks.Item(e.Index.ToString())
                CalculateRate(oTask)
            Case E_EVENTTARGET.EVT_ROW
                Dim i As Integer = 0
                Dim oRow As clsRow = Nothing
                For i = 1 To ActiveGanttVBECtl1.Rows.Count
                    oRow = ActiveGanttVBECtl1.Rows.Item(i.ToString())
                    If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                        '//TODO
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                        '//TODO
                    ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                        Dim oCRRow As AG_CR_Row = Nothing
                        For Each oCRRow In mp_o_AG_CR_Rows
                            If oCRRow.lRowID = System.Convert.ToInt32(oRow.Key.Replace("K", "")) Then
                                Exit For
                            End If
                        Next
                        oCRRow.lOrder = i
                    End If
                Next
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectSize(sender As Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectSize
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                Dim oTask As clsTask = Nothing
                oTask = ActiveGanttVBECtl1.Tasks.Item(e.Index.ToString())
                CalculateRate(oTask)
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_ControlMouseDown(sender As Object, e As AGVBE.MouseEventArgs) Handles ActiveGanttVBECtl1.ControlMouseDown
        Dim lIndex As Integer
        mp_oTaskPopUp.IsOpen = False
        mp_oRowPopUp.IsOpen = False
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_SELECTEDROW, E_EVENTTARGET.EVT_ROW
                lIndex = ActiveGanttVBECtl1.MathLib.GetRowIndexByPosition(e.Y)
                If e.Button = E_MOUSEBUTTONS.BTN_LEFT Then
                    Dim oRow As clsRow
                    oRow = ActiveGanttVBECtl1.Rows.Item(ActiveGanttVBECtl1.MathLib.GetRowIndexByPosition(e.Y).ToString())
                    If e.X > ActiveGanttVBECtl1.Splitter.Position - 20 And e.X < ActiveGanttVBECtl1.Splitter.Position - 5 And e.Y < oRow.Bottom - 5 And e.Y > oRow.Bottom - 20 Then
                        oRow.Node.Expanded = Not oRow.Node.Expanded
                        ActiveGanttVBECtl1.Redraw()
                        e.Cancel = True
                    End If
                ElseIf e.Button = E_MOUSEBUTTONS.BTN_RIGHT Then
                    e.Cancel = True
                    mp_sEditRowKey = ActiveGanttVBECtl1.Rows.Item(lIndex.ToString()).Key
                    mp_oRowPopUp.SetValue(Canvas.LeftProperty, mp_oCursorPosition.X)
                    mp_oRowPopUp.SetValue(Canvas.TopProperty, mp_oCursorPosition.Y)
                    mp_oRowPopUp.IsOpen = True
                End If
            Case E_EVENTTARGET.EVT_SELECTEDTASK, E_EVENTTARGET.EVT_TASK
                If e.Button = E_MOUSEBUTTONS.BTN_RIGHT Then
                    Dim sTag As String
                    e.Cancel = True
                    lIndex = ActiveGanttVBECtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y)
                    mp_sEditTaskKey = ActiveGanttVBECtl1.Tasks.Item(lIndex.ToString()).Key
                    sTag = ActiveGanttVBECtl1.Tasks.Item(lIndex.ToString()).Tag
                    If sTag = "0" Then
                        mp_mnuConvertToRental.Visibility = Windows.Visibility.Visible
                    Else
                        mp_mnuConvertToRental.Visibility = Windows.Visibility.Collapsed
                    End If
                    If sTag = "2" Then
                        mp_mnuEditTask.Visibility = Windows.Visibility.Collapsed
                    Else
                        mp_mnuEditTask.Visibility = Windows.Visibility.Visible
                    End If
                    mp_oTaskPopUp.SetValue(Canvas.LeftProperty, mp_oCursorPosition.X)
                    mp_oTaskPopUp.SetValue(Canvas.TopProperty, mp_oCursorPosition.Y)
                    mp_oTaskPopUp.IsOpen = True
                End If
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_MouseMove(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles ActiveGanttVBECtl1.MouseMove
        mp_oCursorPosition.X = e.GetPosition(LayoutRoot).X
        mp_oCursorPosition.Y = e.GetPosition(LayoutRoot).Y
    End Sub

    Private Sub ActiveGanttVBECtl1_ControlKeyDown(sender As Object, e As AGVBE.KeyEventArgs) Handles ActiveGanttVBECtl1.ControlKeyDown
        If e.KeyCode = Key.F2 Then
            Mode = HPE_ADDMODE.AM_RENTAL
        End If
        If e.KeyCode = Key.F9 Then
            Mode = HPE_ADDMODE.AM_MAINTENANCE
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_ControlKeyUp(sender As Object, e As AGVBE.KeyEventArgs) Handles ActiveGanttVBECtl1.ControlKeyUp
        Mode = HPE_ADDMODE.AM_RESERVATION
    End Sub

    Private Sub ActiveGanttVBECtl1_ControlMouseWheel(sender As Object, e As AGVBE.MouseWheelEventArgs) Handles ActiveGanttVBECtl1.ControlMouseWheel
        If (e.Delta = 0) Or (ActiveGanttVBECtl1.VerticalScrollBar.Visible = False) Then
            Return
        End If
        Dim lDelta As Integer = System.Convert.ToInt32(-(e.Delta / 100))
        Dim lInitialValue As Integer = ActiveGanttVBECtl1.VerticalScrollBar.Value
        If (ActiveGanttVBECtl1.VerticalScrollBar.Value + lDelta < 1) Then
            ActiveGanttVBECtl1.VerticalScrollBar.Value = 1
        ElseIf (((ActiveGanttVBECtl1.VerticalScrollBar.Value + lDelta) > ActiveGanttVBECtl1.VerticalScrollBar.Max)) Then
            ActiveGanttVBECtl1.VerticalScrollBar.Value = ActiveGanttVBECtl1.VerticalScrollBar.Max
        Else
            ActiveGanttVBECtl1.VerticalScrollBar.Value = ActiveGanttVBECtl1.VerticalScrollBar.Value + lDelta
        End If
        ActiveGanttVBECtl1.Redraw()
    End Sub

#End Region

#Region "Page Properties"

    Friend Property DataSource As E_DATASOURCETYPE
        Get
            Return mp_yDataSourceType
        End Get
        Set(ByVal Value As E_DATASOURCETYPE)
            mp_yDataSourceType = Value
        End Set
    End Property

    Public Property Mode() As HPE_ADDMODE
        Get
            Return mp_yAddMode
        End Get
        Set(ByVal Value As HPE_ADDMODE)
            mp_yAddMode = Value
            Select Case mp_yAddMode
                Case HPE_ADDMODE.AM_RESERVATION
                    lblMode.Content = "Add Reservation Mode"
                    lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 153, 170, 194))
                    mp_sAddModeStyleIndex = "Reservation"
                Case HPE_ADDMODE.AM_RENTAL
                    lblMode.Content = "Add Rental Mode"
                    lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 162, 78, 50))
                    mp_sAddModeStyleIndex = "Rental"
                Case HPE_ADDMODE.AM_MAINTENANCE
                    lblMode.Content = "Add Maintenance Mode"
                    lblMode.Background = New SolidColorBrush(Color.FromArgb(255, 255, 77, 1))
                    mp_sAddModeStyleIndex = "Maintenance"
            End Select
        End Set
    End Property

    Private Property Zoom() As Integer
        Get
            Return mp_lZoom
        End Get
        Set(ByVal Value As Integer)
            If Value > 5 Or Value < 1 Then
                Return
            End If
            mp_lZoom = Value
            Dim oView As clsView = Nothing
            oView = ActiveGanttVBECtl1.CurrentViewObject
            Select Case mp_lZoom
                Case 5
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 30
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 12
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 4
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 15
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 6
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 3
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 10
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 3
                    oView.TimeLine.TickMarkArea.Visible = False
                Case 2
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 5
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_HOUR
                    oView.ClientArea.Grid.Factor = 2
                    oView.TimeLine.TickMarkArea.Visible = True
                    oView.TimeLine.TickMarkArea.Height = 30
                    oView.TimeLine.TickMarkArea.TickMarks.Clear()
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 6, E_TICKMARKTYPES.TLT_BIG, True, "hh:mmtt")
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 1, E_TICKMARKTYPES.TLT_SMALL, False, "h")
                Case 1
                    oView.Interval = E_INTERVAL.IL_MINUTE
                    oView.Factor = 1
                    oView.ClientArea.Grid.Interval = E_INTERVAL.IL_MINUTE
                    oView.ClientArea.Grid.Factor = 15
                    oView.TimeLine.TickMarkArea.Visible = True
                    oView.TimeLine.TickMarkArea.Height = 30
                    oView.TimeLine.TickMarkArea.TickMarks.Clear()
                    oView.TimeLine.TickMarkArea.TickMarks.Add(E_INTERVAL.IL_HOUR, 1, E_TICKMARKTYPES.TLT_BIG, True, "hh:mmtt")
            End Select
            ActiveGanttVBECtl1.Redraw()
        End Set
    End Property

#End Region

#Region "Functions"

    Friend Function GetDescription(ByVal lCarTypeID As Integer) As String
        Dim sReturn As String = ""
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oCarType As AG_CR_Car_Type
            For Each oCarType In mp_o_AG_CR_Car_Types
                If oCarType.lCarTypeID = lCarTypeID Then
                    sReturn = oCarType.sDescription
                End If
            Next
        End If
        Return sReturn
    End Function

    Private Sub CalculateRate(ByRef oTask As clsTask)
        Dim dFactor As Double = 0
        Dim sRowTag As String()
        Dim lRate As Double = 0
        Dim dSubTotal As Double = 0
        Dim dOptions As Double = 0
        Dim dSurcharge As Double = 0
        Dim dTax As Double = 0
        Dim dALI As Double = 0
        Dim dCRF As Double = 0
        Dim dERF As Double = 0
        Dim dGPS As Double = 0
        Dim dLDW As Double = 0
        Dim dPAI As Double = 0
        Dim dPEP As Double = 0
        Dim dRCFC As Double = 0
        Dim dVLF As Double = 0
        Dim dWTB As Double = 0
        Dim bGPS As Boolean = False
        Dim bLDW As Boolean = False
        Dim bPAI As Boolean = False
        Dim bPEP As Boolean = False
        Dim bALI As Boolean = False
        Dim sName As String = ""
        Dim sPhone As String = ""

        Dim sEstimatedTotal As String = ""
        Dim dEstimatedTotal As Double = 0


        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oRental As AG_CR_Rental = Nothing
            For Each oRental In mp_o_AG_CR_Rentals
                If oRental.lTaskID = System.Convert.ToInt32(oTask.Key.Replace("K", "")) Then
                    Exit For
                End If
            Next
            sName = oRental.sName
            sPhone = oRental.sPhone

            bGPS = oRental.bGPS
            dGPS = oRental.dGPS
            bLDW = oRental.bLDW
            dLDW = oRental.dLDW
            bPAI = oRental.bPAI
            dPAI = oRental.dPAI
            bPEP = oRental.bPEP
            dPEP = oRental.dPEP
            bALI = oRental.bALI
            dALI = oRental.dALI

            dERF = oRental.dERF
            dWTB = oRental.dWTB
            dRCFC = oRental.dRCFC
            dVLF = oRental.dVLF
            dCRF = oRental.dCRF
        End If

        dFactor = System.Convert.ToDouble(ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_HOUR, oTask.StartDate, oTask.EndDate) / 24)

        If bGPS = True Then
            dGPS = dGPS * dFactor
        Else
            dGPS = 0
        End If
        If bLDW = True Then
            dLDW = dLDW * dFactor
        Else
            dLDW = 0
        End If
        If bPAI = True Then
            dPAI = dPAI * dFactor
        Else
            dPAI = 0
        End If
        If bPEP = True Then
            dPEP = dPEP * dFactor
        Else
            dPEP = 0
        End If
        If bALI = True Then
            dALI = dALI * dFactor
        Else
            dALI = 0
        End If
        sRowTag = oTask.Row.Tag.Split("|"c)
        lRate = CType(sRowTag(1), System.Double)
        dERF = dERF * dFactor
        dWTB = dWTB * dFactor
        dRCFC = dRCFC * dFactor
        dVLF = dVLF * dFactor
        dCRF = dCRF * lRate * dFactor
        dSurcharge = dERF + dWTB + dRCFC + dVLF + dCRF
        dOptions = CType(dGPS + dLDW + dPAI + dPEP + dALI, System.Double)
        dSubTotal = CType(dSurcharge + (lRate * dFactor), System.Double)
        Dim sState As String = ""
        dTax = dSubTotal * GetStateTax(oTask, sState)
        dEstimatedTotal = dSubTotal + dTax + dOptions
        sEstimatedTotal = dEstimatedTotal.ToString("0.00")
        If oTask.Tag = "0" Or oTask.Tag = "1" Then
            oTask.Text = sName & vbCrLf & "Phone: " & sPhone & vbCrLf & "Estimated Total: " & sEstimatedTotal & " USD"
        Else
            dEstimatedTotal = 0
            lRate = 0
        End If

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oRental As AG_CR_Rental = Nothing
            For Each oRental In mp_o_AG_CR_Rentals
                If oRental.lTaskID = System.Convert.ToInt32(oTask.Key.Replace("K", "")) Then
                    Exit For
                End If
            Next
            oRental.dtPickUp = oTask.StartDate.DateTimePart
            oRental.dtReturn = oTask.EndDate.DateTimePart
            oRental.dRate = lRate
            oRental.dEstimatedTotal = dEstimatedTotal
        End If
    End Sub

    Friend Function GetStateTax(ByRef oTask As clsTask, ByRef sState As String) As Double
        Dim oNode As clsNode = Nothing
        Dim dTax As Double = 0
        oNode = oTask.Row.Node.Parent()
        If oNode Is Nothing Then
            Return 0.1
        Else
            If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                '//TODO
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                '//TODO
            ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                Dim oCRRow As AG_CR_Row
                For Each oCRRow In mp_o_AG_CR_Rows
                    If oCRRow.lRowID = System.Convert.ToInt32(oNode.Row.Key.Replace("K", "")) Then
                        sState = oCRRow.sState
                        Exit For
                    End If
                Next
                Dim oState As AG_CR_US_State
                For Each oState In mp_o_AG_CR_US_States
                    If oState.sState = sState Then
                        dTax = System.Convert.ToDouble(oState.dCarRentalTax)
                        Exit For
                    End If
                Next
            End If
        End If
        Return dTax
    End Function

    Friend Function GetImage(ByVal sImage As String) As Image
        Dim oReturnImage = New Image
        Dim oURI As System.Uri = Nothing
        If App.Current.Host.Source.ToString().Contains("file:///") Then
            Dim sSource = App.Current.Host.Source.ToString()
            sSource = sSource.Substring(0, sSource.IndexOf("AGVBECON")) & "AGVBECON.Web/" & sImage.Replace("\", "/")
            oURI = New System.Uri(sSource)
        Else
            oURI = New System.Uri(App.Current.Host.Source, "../" & sImage)
        End If
        Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
        AddHandler oBitmap.ImageOpened, AddressOf mp_oBitmapOpened
        oBitmap.UriSource = oURI
        If sImage.Contains("/Small/") = True Then
            oReturnImage.Width = 95
            oReturnImage.Height = 40
        ElseIf sImage.EndsWith("minus.jpg") Or sImage.EndsWith("plus.jpg") Then
            oReturnImage.Width = 14
            oReturnImage.Height = 14
        End If
        oReturnImage.Source = oBitmap
        Return oReturnImage
    End Function

    Private Sub mp_oBitmapOpened()
        ActiveGanttVBECtl1.Redraw()
    End Sub

#End Region

#Region "Toolbar Buttons"

#Region "cmdSave"

    Private Sub cmdSaveXML_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        mp_fSave = New fSave()
        mp_fSave.sSuggestedFileName = "AGVBE_CR"
        mp_fSave.Title = "Save XML File"
        mp_fSave.mp_oFileList = mp_oFileList
        mp_fSave.Show()
    End Sub

    Private Sub mp_fSave_Closed(sender As Object, e As System.EventArgs) Handles mp_fSave.Closed
        If mp_fSave.DialogResult = True Then
            Dim sXML As String = ""
            sXML = ActiveGanttVBECtl1.GetXML()
            If sXML.Length > 0 Then
                invkSetXML = oServiceContext.SetXML(sXML, mp_fSave.sFileName)
                cmdSaveXML.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub invkSetXML_Completed(sender As Object, e As System.EventArgs) Handles invkSetXML.Completed
        invkGetFileList = oServiceContext.GetFileList()
        cmdSaveXML.IsEnabled = True
    End Sub

#End Region

    Private Sub cmdLoadXML_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdLoadXML.Click
        Dim oForm As New fLoadXML()
        Me.Content = oForm
    End Sub

    Private Sub cmdBack_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        Dim oForm As New fMain()
        Me.Content = oForm
    End Sub

    Private Sub cmdZoomin_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdZoomin.Click
        Zoom = Zoom - 1
    End Sub

    Private Sub cmdZoomout_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdZoomout.Click
        Zoom = Zoom + 1
    End Sub

    Private Sub cmdAddVehicle_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdAddVehicle.Click
        mp_fCarRentalVehicle.Mode = PRG_DIALOGMODE.DM_ADD
        mp_fCarRentalVehicle.mp_sRowID = ""
        mp_fCarRentalVehicle.Show()
    End Sub

    Private Sub cmdAddBranch_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdAddBranch.Click
        mp_fCarRentalBranch.Mode = PRG_DIALOGMODE.DM_ADD
        mp_fCarRentalBranch.mp_sRowID = ""
        mp_fCarRentalBranch.Show()
    End Sub

    Private Sub cmdHelp_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdHelp.Click
        System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Article.aspx?ID=16"), "_blank")
    End Sub

    Private Sub cmdBack2_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack2.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, -10 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdBack1_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack1.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, -5 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdBack0_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBack0.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, -1 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdFwd0_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdFwd0.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, 1 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdFwd1_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdFwd1.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, 5 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdFwd2_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdFwd2.Click
        ActiveGanttVBECtl1.CurrentViewObject.TimeLine.Position(ActiveGanttVBECtl1.MathLib.DateTimeAdd(ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Interval, 10 * ActiveGanttVBECtl1.CurrentViewObject.ClientArea.Grid.Factor, ActiveGanttVBECtl1.CurrentViewObject.TimeLine.StartDate))
        ActiveGanttVBECtl1.Redraw()
    End Sub

#End Region

#Region "Context Menus"

    Private Sub InitTaskContextMenu()
        mp_oTaskPopUp = New Popup
        mp_oTaskStackPanel = New StackPanel
        mp_oTaskStackPanel.Background = New SolidColorBrush(Color.FromArgb(255, 200, 200, 200))
        mp_mnuEditTask = New TextBlock
        mp_mnuEditTask.Text = "Edit Task"
        mp_mnuEditTask.Padding = New Thickness(5, 5, 5, 5)
        mp_oTaskStackPanel.Children.Add(mp_mnuEditTask)
        mp_mnuConvertToRental = New TextBlock
        mp_mnuConvertToRental.Text = "Convert to rental"
        mp_mnuConvertToRental.Padding = New Thickness(5, 5, 5, 5)
        mp_oTaskStackPanel.Children.Add(mp_mnuConvertToRental)
        mp_mnuDeleteTask = New TextBlock
        mp_mnuDeleteTask.Text = "Delete Task"
        mp_mnuDeleteTask.Padding = New Thickness(5, 5, 5, 5)
        mp_oTaskStackPanel.Children.Add(mp_mnuDeleteTask)
        mp_oTaskPopUp.Child = mp_oTaskStackPanel
        LayoutRoot.Children.Add(mp_oTaskPopUp)
        AddHandler mp_mnuEditTask.MouseLeftButtonUp, AddressOf mnuEditTask_Click
        AddHandler mp_mnuConvertToRental.MouseLeftButtonUp, AddressOf mnuConvertToRental_Click
        AddHandler mp_mnuDeleteTask.MouseLeftButtonUp, AddressOf mnuDeleteTask_Click
    End Sub

    Private Sub InitRowContextMenu()
        mp_oRowPopUp = New Popup
        mp_oRowStackPanel = New StackPanel
        mp_oRowStackPanel.Background = New SolidColorBrush(Color.FromArgb(255, 200, 200, 200))
        mp_mnuEditRow = New TextBlock
        mp_mnuEditRow.Text = "Edit Row"
        mp_mnuEditRow.Padding = New Thickness(5, 5, 5, 5)
        mp_oRowStackPanel.Children.Add(mp_mnuEditRow)
        mp_mnuDeleteRow = New TextBlock
        mp_mnuDeleteRow.Text = "Delete Row"
        mp_mnuDeleteRow.Padding = New Thickness(5, 5, 5, 5)
        mp_oRowStackPanel.Children.Add(mp_mnuDeleteRow)
        mp_oRowPopUp.Child = mp_oRowStackPanel
        LayoutRoot.Children.Add(mp_oRowPopUp)
        AddHandler mp_mnuEditRow.MouseLeftButtonUp, AddressOf mnuEditRow_Click
        AddHandler mp_mnuDeleteRow.MouseLeftButtonUp, AddressOf mnuDeleteRow_Click
    End Sub

    Private Sub mnuEditTask_Click()
        mp_oTaskPopUp.IsOpen = False
        mp_fCarRentalReservation.Mode = PRG_DIALOGMODE.DM_EDIT
        mp_fCarRentalReservation.mp_sTaskID = mp_sEditTaskKey.Replace("K", "")
        mp_fCarRentalReservation.Show()
    End Sub

    Private Sub mnuConvertToRental_Click()
        Dim oTask As clsTask
        oTask = ActiveGanttVBECtl1.Tasks.Item(mp_sEditTaskKey)
        mp_oTaskPopUp.IsOpen = False
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            '//TODO
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            Dim oRental As AG_CR_Rental
            For Each oRental In mp_o_AG_CR_Rentals
                If oRental.lTaskID = System.Convert.ToInt32(mp_sEditTaskKey.Replace("K", "")) Then
                    oRental.yMode = 1
                    Exit For
                End If
            Next
        End If
        oTask.Tag = "1"
        oTask.StyleIndex = "Rental"
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub mnuDeleteTask_Click()
        mp_oTaskPopUp.IsOpen = False
        mp_fYesNoMsgBox.Prompt = "Are you sure you want to delete this item?"
        mp_fYesNoMsgBox.Type = "DeleteTask"
        mp_fYesNoMsgBox.Show()
    End Sub

    Private Sub mnuEditRow_Click()
        Dim oRow As clsRow
        oRow = ActiveGanttVBECtl1.Rows.Item(mp_sEditRowKey)
        mp_oRowPopUp.IsOpen = False
        If oRow.Node.Depth = 1 Then
            mp_fCarRentalVehicle.Mode = PRG_DIALOGMODE.DM_EDIT
            mp_fCarRentalVehicle.mp_sRowID = mp_sEditRowKey.Replace("K", "")
            mp_fCarRentalVehicle.Show()
        ElseIf oRow.Node.Depth = 0 Then
            mp_fCarRentalBranch.Mode = PRG_DIALOGMODE.DM_EDIT
            mp_fCarRentalBranch.mp_sRowID = mp_sEditRowKey.Replace("K", "")
            mp_fCarRentalBranch.Show()
        End If
    End Sub

    Private Sub mnuDeleteRow_Click()
        mp_oRowPopUp.IsOpen = False
        mp_fYesNoMsgBox.Prompt = "Are you sure you want to delete this item?"
        mp_fYesNoMsgBox.Type = "DeleteRow"
        mp_fYesNoMsgBox.Show()
    End Sub

    Private Sub YesNoMsgBox_Closed()
        If mp_fYesNoMsgBox.Type = "DeleteRow" Then
            If mp_fYesNoMsgBox.DialogResult = True Then
                If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                    '//TODO
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                    '//TODO
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                    Dim oCRRow As AG_CR_Row
                    Dim oRental As AG_CR_Rental
                    Dim i As Integer
                    Dim oRentalRemoveList As New List(Of AG_CR_Rental)
                    For Each oCRRow In mp_o_AG_CR_Rows
                        If oCRRow.lRowID = System.Convert.ToInt32(mp_sEditRowKey.Replace("K", "")) Then
                            mp_o_AG_CR_Rows.Remove(oCRRow)
                            Exit For
                        End If
                    Next
                    For Each oRental In mp_o_AG_CR_Rentals
                        If oRental.lRowID = System.Convert.ToInt32(mp_sEditRowKey.Replace("K", "")) Then
                            oRentalRemoveList.Add(oRental)
                        End If
                    Next
                    For i = 0 To oRentalRemoveList.Count - 1
                        oRental = oRentalRemoveList(i)
                        mp_o_AG_CR_Rentals.Remove(oRental)
                    Next
                End If
                ActiveGanttVBECtl1.Rows.Remove(mp_sEditRowKey)
                ActiveGanttVBECtl1.Rows.UpdateTree()
                ActiveGanttVBECtl1.Redraw()
            End If
        ElseIf mp_fYesNoMsgBox.Type = "DeleteTask" Then
            If mp_fYesNoMsgBox.DialogResult = True Then
                Dim oRental As AG_CR_Rental
                If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                    '//TODO
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                    '//TODO
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
                    For Each oRental In mp_o_AG_CR_Rentals
                        If oRental.lTaskID = System.Convert.ToInt32(mp_sEditTaskKey.Replace("K", "")) Then
                            mp_o_AG_CR_Rentals.Remove(oRental)
                            Exit For
                        End If
                    Next
                End If
                ActiveGanttVBECtl1.Tasks.Remove(mp_sEditTaskKey)
                ActiveGanttVBECtl1.Redraw()
            End If
        End If
    End Sub

#End Region

#Region "Load Data"

    Private Sub NoDataSource_LoadRowsAndTasks()
        Dim sRowID As String = ""
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing

        NoDataSorce_Load_Rows()

        Dim oCRRow As AG_CR_Row
        For Each oCRRow In mp_o_AG_CR_Rows
            sRowID = "K" & oCRRow.lRowID.ToString()
            oRow = ActiveGanttVBECtl1.Rows.Add(sRowID)
            oRow.AllowTextEdit = True
            If oCRRow.lDepth = 0 Then
                oRow.Text = oCRRow.sBranchName & ", " & oCRRow.sState & vbCrLf & "Phone: " & oCRRow.sPhone
                oRow.MergeCells = True
                oRow.Container = False
                oRow.StyleIndex = "Branch"
                oRow.ClientAreaStyleIndex = "BranchCA"
                oRow.Node.Depth = 0
                oRow.UseNodeImages = True
                oRow.Node.ExpandedImage = GetImage("CarRental\minus.jpg")
                oRow.Node.Image = GetImage("CarRental\plus.jpg")
                oRow.AllowMove = False
                oRow.AllowSize = False
            ElseIf oCRRow.lDepth = 1 Then
                Dim sDescription As String = ""
                sDescription = GetDescription(oCRRow.lCarTypeID)
                oRow.Cells.Item("1").Text = oCRRow.sLicensePlates
                oRow.Cells.Item("1").AllowTextEdit = True
                oRow.Cells.Item("2").Image = GetImage("CarRental/Small/" & sDescription & ".jpg")
                oRow.Cells.Item("3").Text = sDescription & vbCrLf & oCRRow.sACRISSCode & " - " & oCRRow.dRate.ToString() & " USD"
                oRow.Cells.Item("3").AllowTextEdit = True
                oRow.Node.Depth = 1
                oRow.Tag = oCRRow.sACRISSCode & "|" & oCRRow.dRate.ToString() & "|" & oCRRow.lCarTypeID.ToString()
            End If
        Next


        NoDataSorce_Load_Rentals()

        Dim oRental As AG_CR_Rental
        For Each oRental In mp_o_AG_CR_Rentals
            oTask = ActiveGanttVBECtl1.Tasks.Add("", "K" & oRental.lRowID.ToString(), FromDate(oRental.dtPickUp), FromDate(oRental.dtReturn), "K" & oRental.lTaskID.ToString())
            oTask.AllowTextEdit = True
            If oRental.yMode = 2 Then
                oTask.Text = "Scheduled Maintenance"
                oTask.StyleIndex = "Maintenance"
            Else
                oTask.Text = oRental.sName & vbCrLf & "Phone: " & oRental.sPhone & vbCrLf & "Estimated Total: " & oRental.dEstimatedTotal.ToString("00.00") & " USD"
                If oRental.yMode = 0 Then
                    oTask.StyleIndex = "Reservation"
                ElseIf oRental.yMode = 1 Then
                    oTask.StyleIndex = "Rental"
                End If
            End If
            oTask.Tag = oRental.yMode.ToString()
        Next

    End Sub

    Private Sub NoDataSorce_Load_Rows()
        NoDataSorce_Load_Row(28, 0, 1, "", 0, "", 0.0, "", "Hillsboro Beach", "Hillsboro Beach", "FL", "(175) 157-9697", "Nancy Mcatee", "(175) 554-7615", "113 Bueno Drive", "22454")
        NoDataSorce_Load_Row(29, 1, 2, "CKT-2542", 39, "", 245.0, "FFBV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(30, 1, 3, "XXW-9757", 14, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(31, 1, 4, "HGO-6751", 16, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(32, 1, 5, "QIZ-1491", 17, "", 37.0, "ECAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(33, 1, 6, "WGN-3159", 46, "", 77.0, "LCAR", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(34, 1, 8, "TJS-5515", 37, "", 245.0, "FFBV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(35, 1, 9, "FPN-9487", 31, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(36, 1, 10, "ENU-2926", 26, "", 45.0, "FWAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(37, 1, 11, "MND-5686", 11, "", 39.0, "IDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(38, 1, 12, "ZZY-1567", 18, "", 37.0, "ECAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(39, 0, 13, "", 0, "", 0.0, "", "Woodville", "Woodville", "OK", "(145) 548-2974", "Matthew Risner", "(145) 679-8583", "8 Navarro Junction", "61614")
        NoDataSorce_Load_Row(40, 1, 14, "SGL-3748", 24, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(41, 1, 15, "VYW-1478", 43, "", 51.0, "FVAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(42, 1, 16, "LXV-4412", 27, "", 45.0, "FWAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(44, 1, 7, "IMU-3364", 23, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(45, 1, 17, "FRG-8842", 30, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(46, 1, 18, "OJQ-8553", 14, "", 37.0, "EDAZ", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(47, 1, 19, "INT-3737", 5, "", 223.0, "PWDV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(48, 1, 20, "USM-8758", 47, "", 77.0, "LCAR", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(49, 1, 21, "RRL-2724", 32, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(50, 1, 22, "EMF-3865", 20, "", 37.0, "CDAV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(51, 1, 23, "SRC-5911", 32, "", 37.0, "CDMV", "", "", "", "", "", "", "", "")
        NoDataSorce_Load_Row(52, 1, 24, "VTN-9768", 3, "", 71.0, "IFBV", "", "", "", "", "", "", "", "")
    End Sub

    Private Sub NoDataSorce_Load_Row(ByVal lRowID As Integer, ByVal lDepth As Integer, ByVal lOrder As Integer, ByVal sLicensePlates As String, ByVal lCarTypeID As Integer, ByVal sNotes As String, ByVal dRate As Double, ByVal sACRISSCode As String, ByVal sCity As String, ByVal sBranchName As String, ByVal sState As String, ByVal sPhone As String, ByVal sManagerName As String, ByVal sManagerMobile As String, ByVal sAddress As String, ByVal sZIP As String)
        Dim oRow As New AG_CR_Row()
        oRow.lRowID = lRowID
        oRow.lDepth = lDepth
        oRow.lOrder = lOrder
        oRow.sLicensePlates = sLicensePlates
        oRow.lCarTypeID = lCarTypeID
        oRow.sNotes = sNotes
        oRow.dRate = dRate
        oRow.sACRISSCode = sACRISSCode
        oRow.sCity = sCity
        oRow.sBranchName = sBranchName
        oRow.sState = sState
        oRow.sPhone = sPhone
        oRow.sManagerName = sManagerName
        oRow.sManagerMobile = sManagerMobile
        oRow.sAddress = sAddress
        oRow.sZIP = sZIP
        mp_o_AG_CR_Rows.Add(oRow)
    End Sub

    Private Sub NoDataSorce_Load_Rentals()
        NoDataSorce_Load_Rental(21, 30, 0, "Jeromy Lapham", "33 Mckinley Plaza", "Munds Park", "AZ", "37167", "(532) 463-3173", "(532) 793-8291", New AGVBE.DateTime(2009, 6, 13, 0, 0, 0), New AGVBE.DateTime(2009, 6, 20, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 359.22, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(22, 34, 0, "Colleen Nagle", "21 Graziano Street", "George", "SC", "99234", "(266) 819-5725", "(266) 876-2444", New AGVBE.DateTime(2009, 6, 12, 0, 0, 0), New AGVBE.DateTime(2009, 6, 18, 12, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1923.27, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(23, 36, 0, "Luisa Farrior", "86 Wiegand Courts", "Dayton", "VA", "79821", "(417) 727-8137", "(417) 974-9449", New AGVBE.DateTime(2009, 6, 10, 12, 0, 0), New AGVBE.DateTime(2009, 6, 26, 0, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 941.21, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(25, 32, 0, "Nancy Sandusky", "4 Babcock Street", "Arlington Heights village", "IL", "37895", "(446) 926-4519", "(446) 552-5686", New AGVBE.DateTime(2009, 6, 9, 12, 0, 0), New AGVBE.DateTime(2009, 6, 18, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 461.85, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(26, 29, 0, "Shawn Kidder", "7 Hynes Street", "Vernon Center", "MN", "71625", "(675) 132-8559", "(675) 568-8572", New AGVBE.DateTime(2009, 6, 19, 0, 0, 0), New AGVBE.DateTime(2009, 6, 25, 0, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1847.03, True, False, False, False, False, False)
        NoDataSorce_Load_Rental(27, 33, 1, "Josephina Kuo", "7 Gruber Stravenue", "North Adams", "MA", "29555", "(585) 968-9925", "(585) 789-1551", New AGVBE.DateTime(2009, 6, 11, 12, 0, 0), New AGVBE.DateTime(2009, 6, 22, 12, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1081.84, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(28, 35, 1, "Sherie Gebhard", "241 Booth Lock", "Bauxite", "AR", "73573", "(893) 882-9983", "(893) 854-1831", New AGVBE.DateTime(2009, 6, 11, 0, 0, 0), New AGVBE.DateTime(2009, 6, 21, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 513.16, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(29, 29, 0, "Linda Roscoe", "17 Rosenberry Underpass", "Siler", "NC", "23686", "(929) 872-1524", "(929) 546-9944", New AGVBE.DateTime(2009, 6, 11, 0, 0, 0), New AGVBE.DateTime(2009, 6, 17, 0, 0, 0), 245.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 1775.33, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(30, 37, 0, "Matthew Alfred", "298 Burcham Street", "Kivalina", "AK", "88648", "(896) 563-7588", "(896) 973-8419", New AGVBE.DateTime(2009, 6, 11, 12, 0, 0), New AGVBE.DateTime(2009, 6, 20, 12, 0, 0), 39.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 483.01, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(31, 31, 1, "Betty Ballew", "6 Gillespie Drive", "Souris", "ND", "99572", "(718) 942-2143", "(718) 726-7799", New AGVBE.DateTime(2009, 6, 13, 0, 0, 0), New AGVBE.DateTime(2009, 6, 23, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(32, 40, 0, "Jame Josephson", "52 Danford Circle", "Arkport village", "NY", "16792", "(289) 991-7674", "(289) 669-9184", New AGVBE.DateTime(2009, 6, 10, 12, 0, 0), New AGVBE.DateTime(2009, 6, 21, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(33, 42, 0, "Penny Holsinger", "1 Mariano Fields", "Seneca Knolls", "NY", "58312", "(372) 274-7459", "(372) 576-9947", New AGVBE.DateTime(2009, 6, 9, 12, 0, 0), New AGVBE.DateTime(2009, 6, 17, 0, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 455.42, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(34, 44, 0, "Linda Gabaldon", "4 Lewellen Boulevard", "Cypress Lake", "FL", "71862", "(626) 786-3444", "(626) 591-2811", New AGVBE.DateTime(2009, 6, 10, 0, 0, 0), New AGVBE.DateTime(2009, 6, 25, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 795.41, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(35, 46, 0, "Gale Cottingham", "717 Seaton Way", "Worthington borough", "PA", "91136", "(799) 683-3813", "(799) 827-3616", New AGVBE.DateTime(2009, 6, 11, 0, 0, 0), New AGVBE.DateTime(2009, 6, 23, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 641.46, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(36, 42, 0, "Brian Grayson", "4 Eckert Drive", "Dunlap village", "IL", "29184", "(598) 441-2575", "(598) 191-9179", New AGVBE.DateTime(2009, 6, 19, 0, 0, 0), New AGVBE.DateTime(2009, 6, 25, 12, 0, 0), 45.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 394.7, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(38, 45, 1, "Dessie Hoffer", "6 Clay Way", "Monett", "MO", "54761", "(648) 657-9664", "(648) 481-3828", New AGVBE.DateTime(2009, 6, 10, 0, 0, 0), New AGVBE.DateTime(2009, 6, 22, 0, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 615.8, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(39, 41, 1, "Vickie Cartier", "43 Jordan Way", "Williamston", "MI", "92739", "(682) 266-8395", "(682) 745-8184", New AGVBE.DateTime(2009, 6, 9, 12, 0, 0), New AGVBE.DateTime(2009, 6, 19, 12, 0, 0), 51.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 677.78, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(41, 47, 0, "Brian Lenoir", "1 Betts Ridges", "Morrisville", "NC", "11594", "(319) 241-1851", "(319) 571-6978", New AGVBE.DateTime(2009, 6, 11, 12, 0, 0), New AGVBE.DateTime(2009, 6, 19, 12, 0, 0), 223.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 2160.16, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(43, 38, 2, "", "", "", "", "", "", "", New AGVBE.DateTime(2009, 6, 11, 0, 0, 0), New AGVBE.DateTime(2009, 6, 24, 12, 0, 0), 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(44, 49, 0, "Allison Peck", "169 Massa Street", "Waldorf", "MD", "91846", "(679) 847-1487", "(679) 513-3341", New AGVBE.DateTime(2009, 6, 10, 0, 0, 0), New AGVBE.DateTime(2009, 6, 17, 12, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 940.05, False, True, True, False, False, False)
        NoDataSorce_Load_Rental(45, 48, 0, "Tiffany Arce", "6 Spires Street", "Hartford village", "IL", "36615", "(362) 357-2429", "(362) 488-4141", New AGVBE.DateTime(2009, 6, 19, 0, 0, 0), New AGVBE.DateTime(2009, 6, 24, 0, 0, 0), 77.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 491.75, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(46, 51, 0, "Felipe Vantassel", "56 Ormsby Street", "Cheswold", "DE", "49225", "(714) 757-2167", "(714) 378-9745", New AGVBE.DateTime(2009, 6, 14, 0, 0, 0), New AGVBE.DateTime(2009, 6, 24, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 538.82, False, False, False, False, False, False)
        NoDataSorce_Load_Rental(47, 50, 0, "Patricia Cook", "22 Goulet Drive", "Wesleyville borough", "PA", "15945", "(421) 352-2962", "(421) 682-7189", New AGVBE.DateTime(2009, 6, 19, 0, 0, 0), New AGVBE.DateTime(2009, 6, 25, 12, 0, 0), 37.0, 14.43, 0.09, 0.67, 11.95, 26.99, 4.0, 2.95, 4.0, 0.6, 2.03, 0.07, 333.56, False, False, False, False, False, False)
    End Sub

    Private Sub NoDataSorce_Load_Rental(ByVal lTaskID As Integer, ByVal lRowID As Integer, ByVal yMode As Integer, ByVal sName As String, ByVal sAddress As String, ByVal sCity As String, ByVal sState As String, ByVal sZIP As String, ByVal sPhone As String, ByVal sMobile As String, ByVal dtPickUp As AGVBE.DateTime, ByVal dtReturn As AGVBE.DateTime, ByVal dRate As Double, ByVal dALI As Double, ByVal dCRF As Double, ByVal dERF As Double, ByVal dGPS As Double, ByVal dLDW As Double, ByVal dPAI As Double, ByVal dPEP As Double, ByVal dRCFC As Double, ByVal dVLF As Double, ByVal dWTB As Double, ByVal dTax As Double, ByVal dEstimatedTotal As Double, ByVal bGPS As Boolean, ByVal bFSO As Boolean, ByVal bLDW As Boolean, ByVal bPAI As Boolean, ByVal bPEP As Boolean, ByVal bALI As Boolean)
        Dim oRental As New AG_CR_Rental()
        oRental.lTaskID = lTaskID
        oRental.lRowID = lRowID
        oRental.yMode = yMode
        oRental.sName = sName
        oRental.sAddress = sAddress
        oRental.sCity = sCity
        oRental.sState = sState
        oRental.sZIP = sZIP
        oRental.sPhone = sPhone
        oRental.sMobile = sMobile
        oRental.dtPickUp = dtPickUp.DateTimePart
        oRental.dtReturn = dtReturn.DateTimePart
        oRental.dRate = dRate
        oRental.dALI = dALI
        oRental.dCRF = dCRF
        oRental.dERF = dERF
        oRental.dGPS = dGPS
        oRental.dLDW = dLDW
        oRental.dPAI = dPAI
        oRental.dPEP = dPEP
        oRental.dRCFC = dRCFC
        oRental.dVLF = dVLF
        oRental.dWTB = dWTB
        oRental.dTax = dTax
        oRental.dEstimatedTotal = dEstimatedTotal
        oRental.bGPS = bGPS
        oRental.bFSO = bFSO
        oRental.bLDW = bLDW
        oRental.bPAI = bPAI
        oRental.bPEP = bPEP
        oRental.bALI = bALI
        mp_o_AG_CR_Rentals.Add(oRental)
    End Sub

    Private Sub NoDataSource_Load_Car_Types()
        NoDataSource_Add_Car_Type(1, "Escape Panther Black", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(2, "Escape Hot Red", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(3, "Escape Atlantis Blue", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(4, "Escape Metalic Sand", "IFBV", 71.0)
        NoDataSource_Add_Car_Type(5, "Territory TX RWD Ego", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(6, "Territory TX RWD Kashmir", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(7, "Territory TX RWD Steel", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(8, "Territory TX RWD Silhouette", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(9, "Territory TX RWD Winter White", "PWDV", 223.0)
        NoDataSource_Add_Car_Type(10, "Mondeo LX Sea Grey", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(11, "Mondeo LX Ink Blue", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(12, "Mondeo LX Colorado Red", "IDAV", 39.0)
        NoDataSource_Add_Car_Type(13, "Fiesta CL 5 Door Squeeze", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(14, "Fiesta CL 5 Door Hydro", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(15, "Fiesta CL 5 Door Panther Black", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(16, "Fiesta CL 5 Door Frozen White", "EDAZ", 37.0)
        NoDataSource_Add_Car_Type(17, "Fiesta CL 3 Door Ocean", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(18, "Fiesta CL 3 Door Hydro", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(19, "Fiesta CL 3 Door Panther Black", "ECAZ", 37.0)
        NoDataSource_Add_Car_Type(20, "Focus CL Sedan Satin White", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(21, "Focus CL Sedan Titanium Grey", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(22, "Focus CL Sedan Black Sapphire", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(23, "Focus CL Sedan Tango", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(24, "Focus CL Sedan Ocean", "CDAV", 37.0)
        NoDataSource_Add_Car_Type(25, "Falcon XT Wagon Lightning Strike", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(26, "Falcon XT Wagon Silhoutte", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(27, "Falcon XT Wagon Sensation", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(28, "Falcon XT Wagon Vixen", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(29, "Falcon XT Wagon Steel", "FWAV", 45.0)
        NoDataSource_Add_Car_Type(30, "Focus CL Hatch Ocean", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(31, "Focus CL Hatch Black Sapphire", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(32, "Focus CL Hatch Tonic", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(33, "Focus CL Hatch Colorado Red", "CDMV", 37.0)
        NoDataSource_Add_Car_Type(34, "Range Rover HSE Alaska White", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(35, "Range Rover HSE Rimini", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(36, "Range Rover HSE Galway Green", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(37, "Range Rover HSE Buckingham Blue", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(38, "Range Rover HSE Santorini Black", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(39, "Range Rover HSE Zermatt Silver", "FFBV", 245.0)
        NoDataSource_Add_Car_Type(40, "LR3 Rimini Red", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(41, "LR3 Santorini Black", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(42, "LR3 Alaska White", "FFBV", 232.0)
        NoDataSource_Add_Car_Type(43, "Town and Country Modern Blue", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(44, "Town and Country Melbourne Green", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(45, "Town and Country Inferno Red", "FVAV", 51.0)
        NoDataSource_Add_Car_Type(46, "Chrysler 300 Clearwater Blue", "LCAR", 77.0)
        NoDataSource_Add_Car_Type(47, "Chrysler 300 Brilliant Black", "LCAR", 77.0)
        NoDataSource_Add_Car_Type(48, "Chrysler 300 Bright Silver", "LCAR", 77.0)
    End Sub

    Private Sub NoDataSource_Add_Car_Type(ByVal lCarTypeID As Integer, ByVal sDescription As String, ByVal sACRISSCode As String, ByVal dStdRate As Double)
        Dim oCarType As New AG_CR_Car_Type()
        oCarType.lCarTypeID = lCarTypeID
        oCarType.sDescription = sDescription
        oCarType.sACRISSCode = sACRISSCode
        oCarType.dStdRate = dStdRate
        mp_o_AG_CR_Car_Types.Add(oCarType)
    End Sub

    Private Sub NoDataSource_Load_US_States()
        NoDataSource_Add_US_State("AK ", "Alaska", 0.0)
        NoDataSource_Add_US_State("AL ", "Alabama", 0.01)
        NoDataSource_Add_US_State("AR ", "Arkansas", 0.01)
        NoDataSource_Add_US_State("AZ ", "Arizona", 0.05)
        NoDataSource_Add_US_State("CA ", "California", 0.06)
        NoDataSource_Add_US_State("CO ", "Colorado", 0.03)
        NoDataSource_Add_US_State("CT ", "Connecticut", 0.06)
        NoDataSource_Add_US_State("DE ", "Delaware", 0.02)
        NoDataSource_Add_US_State("FL ", "Florida", 0.07)
        NoDataSource_Add_US_State("GA ", "Georgia", 0.04)
        NoDataSource_Add_US_State("HI ", "Hawaii", 0.05)
        NoDataSource_Add_US_State("IA ", "Iowa", 0.04)
        NoDataSource_Add_US_State("ID ", "Idaho", 0.03)
        NoDataSource_Add_US_State("IL ", "Illinois", 0.02)
        NoDataSource_Add_US_State("IN ", "Indiana", 0.08)
        NoDataSource_Add_US_State("KS ", "Kansas", 0.06)
        NoDataSource_Add_US_State("KY ", "Kentucky", 0.05)
        NoDataSource_Add_US_State("LA ", "Louisiana", 0.03)
        NoDataSource_Add_US_State("MA ", "Massachusetts", 0.06)
        NoDataSource_Add_US_State("MD ", "Maryland", 0.02)
        NoDataSource_Add_US_State("ME ", "Maine", 0.03)
        NoDataSource_Add_US_State("MI ", "Michigan", 0.04)
        NoDataSource_Add_US_State("MN ", "Minnesota", 0.04)
        NoDataSource_Add_US_State("MO ", "Missouri", 0.02)
        NoDataSource_Add_US_State("MS ", "Mississippi", 0.01)
        NoDataSource_Add_US_State("MT ", "Montana", 0.03)
        NoDataSource_Add_US_State("NC", "North Carolina", 0.04)
        NoDataSource_Add_US_State("ND", "North Dakota", 0.08)
        NoDataSource_Add_US_State("NE", "Nebraska", 0.06)
        NoDataSource_Add_US_State("NH", "New Hampshire", 0.07)
        NoDataSource_Add_US_State("NJ", "New Jersey", 0.06)
        NoDataSource_Add_US_State("NM", "New Mexico", 0.03)
        NoDataSource_Add_US_State("NV", "Nevada", 0.02)
        NoDataSource_Add_US_State("NY", "New York", 0.03)
        NoDataSource_Add_US_State("OH", "Ohio", 0.02)
        NoDataSource_Add_US_State("OK", "Oklahoma", 0.03)
        NoDataSource_Add_US_State("OR", "Oregon", 0.04)
        NoDataSource_Add_US_State("PA", "Pennsylvania", 0.05)
        NoDataSource_Add_US_State("RI", "Rhode Island", 0.06)
        NoDataSource_Add_US_State("SC", "South Carolina", 0.05)
        NoDataSource_Add_US_State("SD", "South Dakota", 0.04)
        NoDataSource_Add_US_State("TN", "Tennessee", 0.03)
        NoDataSource_Add_US_State("TX", "Texas", 0.02)
        NoDataSource_Add_US_State("UT", "Utah", 0.05)
        NoDataSource_Add_US_State("VA", "Virginia", 0.06)
        NoDataSource_Add_US_State("VT", "Vermont", 0.05)
        NoDataSource_Add_US_State("WA", "Washington", 0.04)
        NoDataSource_Add_US_State("WI", "Wisconsin", 0.06)
        NoDataSource_Add_US_State("WV", "West Virginia", 0.07)
        NoDataSource_Add_US_State("WY", "Wyoming", 0.08)
    End Sub

    Private Sub NoDataSource_Add_US_State(ByVal sState As String, ByVal sName As String, ByVal dCarRentalTax As Double)
        Dim oState As New AG_CR_US_State()
        oState.sState = sState
        oState.sName = sName
        oState.dCarRentalTax = dCarRentalTax.ToString()
        mp_o_AG_CR_US_States.Add(oState)
    End Sub

    Friend Sub NoDataSource_Load_ACRISS_Codes(ByVal o_AG_CR_ACRISS_Codes As List(Of AG_CR_ACRISS_Code), ByVal lPosition As Integer)
        If lPosition = 0 Or lPosition = 1 Then
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 1, "M", "Mini", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 2, "N", "Mini Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 3, "E", "Economy", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 4, "H", "Economy Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 5, "C", "Compact", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 6, "D", "Compact Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 7, "I", "Intermediate", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 8, "J", "Intermediate Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 9, "S", "Standard", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 10, "R", "Standard Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 11, "F", "Fullsize", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 12, "G", "Fullsize Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 13, "P", "Premium", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 14, "U", "Premium Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 15, "L", "Luxury", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 16, "W", "Luxury Elite", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 17, "O", "Oversize", 1)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 18, "X", "Special", 1)
        End If
        If lPosition = 0 Or lPosition = 2 Then
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 19, "B", "2-3 Door", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 20, "C", "2/4 Door", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 21, "D", "4-5 Door", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 22, "W", "Wagon/Estate", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 23, "V", "Passenger Van", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 24, "L", "Limousine", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 25, "S", "Sport", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 26, "T", "Convertible", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 27, "F", "SUV", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 28, "J", "Open Air All Terrain", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 29, "X", "Special", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 30, "P", "Pick up Regular Cab", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 31, "Q", "Pick up Extended Cab", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 32, "Z", "Special Offer Car", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 33, "E", "Coupe", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 34, "M", "Monospace", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 35, "R", "Recreational Vehicle", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 36, "H", "Motor Home", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 37, "Y", "2 Wheel Vehicle", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 38, "N", "Roadster", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 39, "G", "Crossover", 2)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 40, "K", "Commercial Van/Truck", 2)
        End If
        If lPosition = 0 Or lPosition = 3 Then
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 41, "M", "Manual Unspecified Drive", 3)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 42, "N", "Manual 4WD", 3)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 43, "C", "Manual AWD", 3)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 44, "A", "Auto Unspecified Drive", 3)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 45, "B", "Auto 4WD", 3)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 46, "D", "Auto AWD", 3)
        End If
        If lPosition = 0 Or lPosition = 4 Then
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 47, "R", "Unspecified Fuel/Power With Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 48, "N", "Unspecified Fuel/Power Without Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 49, "D", "Diesel Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 50, "Q", "Diesel No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 51, "H", "Hybrid Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 52, "I", "Hybrid No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 53, "E", "Electric Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 54, "C", "Electric No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 55, "L", "LPG/Compressed Gas Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 56, "S", "LPG/Compressed Gas No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 57, "A", "Hydrogen Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 58, "B", "Hydrogen No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 59, "M", "Multi Fuel/Power Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 60, "F", "Multi Fuel/Power No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 61, "V", "Petrol Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 62, "Z", "Petrol No Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 63, "U", "Ethanol Air", 4)
            NoDataSource_Add_ACRISS_Code(o_AG_CR_ACRISS_Codes, 64, "X", "Ethanol No Air", 4)
        End If
    End Sub

    Private Sub NoDataSource_Add_ACRISS_Code(ByVal o_AG_CR_ACRISS_Codes As List(Of AG_CR_ACRISS_Code), ByVal lID As Integer, ByVal sLetter As String, ByVal sDescription As String, ByVal lPosition As Integer)
        Dim oACRISSCode As New AG_CR_ACRISS_Code()
        oACRISSCode.lID = lID
        oACRISSCode.sLetter = sLetter
        oACRISSCode.sDescription = sDescription
        oACRISSCode.lPosition = lPosition
        o_AG_CR_ACRISS_Codes.Add(oACRISSCode)
    End Sub

    Private Sub NoDataSource_Load_Taxes_Surcharges_Options()
        NoDataSource_Add_Taxes_Surcharges_Options("ALI", "Additional Liability Insurance", 14.43)
        NoDataSource_Add_Taxes_Surcharges_Options("CRF", "Concession Recovery Fee", 0.1)
        NoDataSource_Add_Taxes_Surcharges_Options("ERF", "Energy Recovery Fee", 0.67)
        NoDataSource_Add_Taxes_Surcharges_Options("GPS", "GPS", 11.95)
        NoDataSource_Add_Taxes_Surcharges_Options("LDW", "Loss Damage Waiver", 26.99)
        NoDataSource_Add_Taxes_Surcharges_Options("PAI", "Personal Accident Insurance", 4.0)
        NoDataSource_Add_Taxes_Surcharges_Options("PEP", "Personal Effects Protection", 2.95)
        NoDataSource_Add_Taxes_Surcharges_Options("RCFC", "Rental Car Facility Charge", 4.0)
        NoDataSource_Add_Taxes_Surcharges_Options("VLF", "Vehicle License Fee", 0.6)
        NoDataSource_Add_Taxes_Surcharges_Options("WTB", "Waste Tire/Battery", 2.03)
    End Sub

    Private Sub NoDataSource_Add_Taxes_Surcharges_Options(ByVal sID As String, ByVal sDescription As String, ByVal dRate As Double)
        Dim oTSO As New AG_CR_Tax_Surcharge_Option()
        oTSO.sID = sID
        oTSO.sDescription = sDescription
        oTSO.dRate = dRate
        mp_o_AG_CR_Taxes_Surcharges_Options.Add(oTSO)
    End Sub

#End Region

End Class
