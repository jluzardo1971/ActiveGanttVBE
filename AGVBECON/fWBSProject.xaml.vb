Imports AGVBE
Imports System.Windows.Threading
Imports System.ServiceModel.DomainServices.Client
Imports AGVBECON.Web

Partial Public Class fWBSProject
    Inherits Page

    Private mp_dtStartDate As AGVBE.DateTime
    Private mp_dtEndDate As AGVBE.DateTime
    Private Const mp_sFontName As String = "Tahoma"
    Friend mp_yDataSourceType As E_DATASOURCETYPE
    Private mp_oToolTip As Border
    Private mp_oStackPanel As StackPanel
    Private mp_oTitleGrid As Grid
    Private mp_oToolTipTitle As TextBlock
    Private mp_oToolTipText As TextBlock
    Private mp_oToolTipLine As Line
    Private mp_oToolTipImage As Image
    Private mp_bBluePercentagesVisible As Boolean = True
    Private mp_bGreenPercentagesVisible As Boolean = True
    Private mp_bRedPercentagesVisible As Boolean = True
    Private mp_fWBSPProperties As fWBSPProperties

    Private WithEvents invkGetFileList As InvokeOperation(Of String)
    Private WithEvents invkSetXML As InvokeOperation
    Private oServiceContext As New ActiveGanttXMLContext()
    Private mp_oFileList As List(Of CON_File)
    Private WithEvents mp_fSave As fSave


#Region "Constructors"

    Public Sub New()
        InitializeComponent()
        mp_fWBSPProperties = New fWBSPProperties(Me)
    End Sub

#End Region

#Region "Page Loaded"

    Private Sub fWBSProject_Loaded(sender As Object, e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        InitToolTip()

        invkGetFileList = oServiceContext.GetFileList()

        mp_dtStartDate = New AGVBE.DateTime()
        mp_dtEndDate = New AGVBE.DateTime()
        Dim dtStartDate As AGVBE.DateTime = New AGVBE.DateTime()

        Dim oStyle As clsStyle = Nothing
        Dim oView As clsView = Nothing

        oStyle = ActiveGanttVBECtl1.Styles.Add("ControlStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BackColor = Color.FromArgb(255, 240, 240, 240)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ScrollBar")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ArrowButtons")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ThumbButtonH")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.EndGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ThumbButtonV")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.EndGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ThumbButtonHP")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.StartGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.EndGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ThumbButtonVP")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL
        oStyle.StartGradientColor = Color.FromArgb(255, 165, 186, 207)
        oStyle.EndGradientColor = Color.FromArgb(255, 240, 240, 240)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 138, 145, 153)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ColumnStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 179, 206, 235)
        oStyle.EndGradientColor = Color.FromArgb(255, 161, 193, 232)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Right = True
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 100, 145, 204)

        oStyle = ActiveGanttVBECtl1.Styles.Add("ScrollBarSeparatorStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.BorderColor = Color.FromArgb(255, 150, 150, 150)

        oStyle = ActiveGanttVBECtl1.Styles.Add("TimeLineTiers")
        oStyle.Font = New Font(mp_sFontName, 7, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_TRANSPARENT
        oStyle.CustomBorderStyle.Left = True
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Right = False
        oStyle.CustomBorderStyle.Bottom = True
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)

        oStyle = ActiveGanttVBECtl1.Styles.Add("TimeLine")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Color.FromArgb(255, 179, 206, 235)
        oStyle.EndGradientColor = Color.FromArgb(255, 161, 193, 232)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.BackColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE

        oStyle = ActiveGanttVBECtl1.Styles.Add("NodeRegular")
        oStyle.Font = New Font(mp_sFontName, 8, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("NodeRegularChecked")
        oStyle.Font = New Font(mp_sFontName, 8, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Normal)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Color.FromArgb(255, 176, 196, 222)
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("NodeBold")
        oStyle.Font = New Font(mp_sFontName, 8, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("NodeBoldChecked")
        oStyle.Font = New Font(mp_sFontName, 8, E_FONTSIZEUNITS.FSU_POINTS, System.Windows.FontWeights.Bold)
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Color.FromArgb(255, 176, 196, 222)
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("ClientAreaChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        oStyle.BackColor = Color.FromArgb(255, 176, 196, 222)

        oStyle = ActiveGanttVBECtl1.Styles.Add("NormalTask")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.PredecessorStyle.LineColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND

        oStyle = ActiveGanttVBECtl1.Styles.Add("NormalTaskWarning")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Color.FromArgb(255, 100, 145, 204)
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.PredecessorStyle.LineColor = Colors.Red
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND

        oStyle = ActiveGanttVBECtl1.Styles.Add("SelectedPredecessor")
        oStyle.PredecessorStyle.LineColor = Colors.Green

        oStyle = ActiveGanttVBECtl1.Styles.Add("GreenSummary")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Colors.Green
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TaskStyle.EndFillColor = Colors.Green
        oStyle.TaskStyle.EndBorderColor = Colors.Green
        oStyle.TaskStyle.StartFillColor = Colors.Green
        oStyle.TaskStyle.StartBorderColor = Colors.Green
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBECtl1.Styles.Add("RedSummary")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Red
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 10
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT
        oStyle.StartGradientColor = Colors.White
        oStyle.EndGradientColor = Colors.Red
        oStyle.GradientFillMode = GRE_GRADIENTFILLMODE.GDT_VERTICAL
        oStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TaskStyle.EndFillColor = Colors.Red
        oStyle.TaskStyle.EndBorderColor = Colors.Red
        oStyle.TaskStyle.StartFillColor = Colors.Red
        oStyle.TaskStyle.StartBorderColor = Colors.Red
        oStyle.TaskStyle.StartShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.TaskStyle.EndShapeIndex = GRE_FIGURETYPE.FT_PROJECTDOWN
        oStyle.FillMode = GRE_FILLMODE.FM_UPPERHALFFILLED

        oStyle = ActiveGanttVBECtl1.Styles.Add("BluePercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Blue
        oStyle.BorderColor = Colors.Blue
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 8
        oStyle.OffsetBottom = 4
        oStyle.SelectionRectangleStyle.Visible = True
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBECtl1.Styles.Add("GreenPercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Green
        oStyle.BorderColor = Colors.Green
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBECtl1.Styles.Add("RedPercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.Red
        oStyle.BorderColor = Colors.Red
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBECtl1.Styles.Add("InvisiblePercentages")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Colors.White
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        oStyle.SelectionRectangleStyle.OffsetTop = 0
        oStyle.SelectionRectangleStyle.OffsetLeft = 0
        oStyle.SelectionRectangleStyle.OffsetRight = 0
        oStyle.SelectionRectangleStyle.OffsetBottom = 0
        oStyle.OffsetTop = 5
        oStyle.OffsetBottom = 5
        oStyle.SelectionRectangleStyle.Visible = False
        oStyle.TextVisible = False
        oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID

        oStyle = ActiveGanttVBECtl1.Styles.Add("ClientAreaStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 197, 206, 216)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.CustomBorderStyle.Right = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("CellStyleKeyColumn")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        oStyle = ActiveGanttVBECtl1.Styles.Add("CellStyle")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Colors.White
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        oStyle = ActiveGanttVBECtl1.Styles.Add("CellStyleKeyColumnChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Color.FromArgb(255, 176, 196, 222)
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False
        oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT
        oStyle.TextXMargin = 4

        oStyle = ActiveGanttVBECtl1.Styles.Add("CellStyleChecked")
        oStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        oStyle.BackColor = Color.FromArgb(255, 176, 196, 222)
        oStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)
        oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        oStyle.CustomBorderStyle.Top = False
        oStyle.CustomBorderStyle.Left = False

        ActiveGanttVBECtl1.ControlTag = "WBSProject"
        ActiveGanttVBECtl1.StyleIndex = "ControlStyle"
        ActiveGanttVBECtl1.ScrollBarSeparator.StyleIndex = "ScrollBarSeparatorStyle"
        ActiveGanttVBECtl1.AllowRowMove = True
        ActiveGanttVBECtl1.AllowRowSize = True
        ActiveGanttVBECtl1.AddMode = E_ADDMODE.AT_BOTH

        Dim oColumn As clsColumn

        oColumn = ActiveGanttVBECtl1.Columns.Add("ID", "", 30, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBECtl1.Columns.Add("Task Name", "", 300, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBECtl1.Columns.Add("StartDate", "", 125, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        oColumn = ActiveGanttVBECtl1.Columns.Add("EndDate", "", 125, "")
        oColumn.StyleIndex = "ColumnStyle"
        oColumn.AllowTextEdit = True

        ActiveGanttVBECtl1.TreeviewColumnIndex = 2

        ActiveGanttVBECtl1.Treeview.Images = True
        ActiveGanttVBECtl1.Treeview.CheckBoxes = True
        ActiveGanttVBECtl1.Treeview.FullColumnSelect = True
        ActiveGanttVBECtl1.Treeview.PlusMinusBorderColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBECtl1.Treeview.PlusMinusSignColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBECtl1.Treeview.CheckBoxBorderColor = Color.FromArgb(255, 100, 145, 204)
        ActiveGanttVBECtl1.Treeview.TreeLineColor = Color.FromArgb(255, 100, 145, 204)

        ActiveGanttVBECtl1.Splitter.Type = E_SPLITTERTYPE.SA_USERDEFINED
        ActiveGanttVBECtl1.Splitter.Width = 1
        ActiveGanttVBECtl1.Splitter.SetColor(1, Color.FromArgb(255, 100, 145, 204))
        ActiveGanttVBECtl1.Splitter.Position = 255

        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.TimerInterval = 50
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonV"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonVP"
        ActiveGanttVBECtl1.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonV"

        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        ActiveGanttVBECtl1.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'Access_LoadTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            'XML_LoadTasks()
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_NONE Then
            NoDataSource_LoadTasks()
        End If
        ActiveGanttVBECtl1.Rows.UpdateTree()

        '// Start one month before the first task:
        dtStartDate = ActiveGanttVBECtl1.MathLib.DateTimeAdd(E_INTERVAL.IL_MONTH, -1, mp_dtStartDate)

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 24, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = False
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False

        oView = ActiveGanttVBECtl1.Views.Add(E_INTERVAL.IL_HOUR, 12, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM, E_TIERTYPE.ST_CUSTOM)
        oView.TimeLine.TierArea.UpperTier.Interval = E_INTERVAL.IL_QUARTER
        oView.TimeLine.TierArea.UpperTier.Factor = 1
        oView.TimeLine.TierArea.UpperTier.Height = 17
        oView.TimeLine.TierArea.MiddleTier.Visible = False
        oView.TimeLine.TierArea.LowerTier.Interval = E_INTERVAL.IL_MONTH
        oView.TimeLine.TierArea.LowerTier.Factor = 1
        oView.TimeLine.TierArea.LowerTier.Height = 17
        oView.TimeLine.TickMarkArea.Visible = False
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 12
        oView.TimeLine.TimeLineScrollBar.LargeChange = 240
        oView.TimeLine.TimeLineScrollBar.Max = 2000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
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
        oView.TimeLine.TimeLineScrollBar.StartDate = dtStartDate
        oView.TimeLine.TimeLineScrollBar.Interval = E_INTERVAL.IL_HOUR
        oView.TimeLine.TimeLineScrollBar.Factor = 1
        oView.TimeLine.TimeLineScrollBar.SmallChange = 6
        oView.TimeLine.TimeLineScrollBar.LargeChange = 480
        oView.TimeLine.TimeLineScrollBar.Max = 4000
        oView.TimeLine.TimeLineScrollBar.Value = 0
        oView.TimeLine.TimeLineScrollBar.Enabled = True
        oView.TimeLine.TimeLineScrollBar.Visible = True
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = "ScrollBar"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = "ArrowButtons"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = "ThumbButtonH"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = "ThumbButtonHP"
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = "ThumbButtonH"
        oView.TimeLine.StyleIndex = "TimeLine"
        oView.ClientArea.DetectConflicts = False

        ActiveGanttVBECtl1.CurrentView = "2"

        ActiveGanttVBECtl1.Redraw()

    End Sub

    Private Sub InitToolTip()
        mp_oToolTip = New Border
        mp_oToolTip.BorderThickness = New Thickness(1)
        mp_oToolTip.BorderBrush = New SolidColorBrush(Colors.Black)
        mp_oToolTip.Background = New SolidColorBrush(Color.FromArgb(255, 255, 255, 224))
        mp_oToolTip.Width = 275
        mp_oToolTip.Visibility = Windows.Visibility.Collapsed
        mp_oStackPanel = New StackPanel

        mp_oTitleGrid = New Grid
        Dim oColumn0 As New ColumnDefinition
        Dim oColumn1 As New ColumnDefinition
        oColumn0.Width = New GridLength(20)
        mp_oTitleGrid.ColumnDefinitions.Add(oColumn0)
        oColumn1.Width = New GridLength(255)
        mp_oTitleGrid.ColumnDefinitions.Add(oColumn1)
        Dim oRow1 As New RowDefinition
        oRow1.Height = GridLength.Auto
        mp_oTitleGrid.RowDefinitions.Add(oRow1)

        mp_oToolTipImage = New Image
        mp_oToolTipImage.Width = 16
        mp_oToolTipImage.Height = 16
        mp_oToolTipImage.VerticalAlignment = Windows.VerticalAlignment.Top

        mp_oTitleGrid.Children.Add(mp_oToolTipImage)
        Grid.SetRow(mp_oToolTipImage, 0)
        Grid.SetColumn(mp_oToolTipImage, 0)


        mp_oToolTipTitle = New TextBlock
        mp_oToolTipTitle.TextWrapping = TextWrapping.Wrap
        mp_oToolTipTitle.TextAlignment = TextAlignment.Center

        mp_oTitleGrid.Children.Add(mp_oToolTipTitle)
        Grid.SetRow(mp_oToolTipTitle, 0)
        Grid.SetColumn(mp_oToolTipTitle, 1)

        mp_oStackPanel.Children.Add(mp_oTitleGrid)

        mp_oToolTipLine = New Line
        mp_oToolTipLine.Stroke = New SolidColorBrush(Colors.Black)
        mp_oToolTipLine.X1 = 0
        mp_oToolTipLine.Y1 = 0
        mp_oToolTipLine.X2 = 275
        mp_oToolTipLine.Y2 = 0

        mp_oStackPanel.Children.Add(mp_oToolTipLine)

        mp_oToolTipText = New TextBlock

        mp_oStackPanel.Children.Add(mp_oToolTipText)

        mp_oToolTip.Child = mp_oStackPanel
        ActiveGanttVBECtl1.ControlCanvas.Children.Add(mp_oToolTip)
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
        If e.TierPosition = E_TIERPOSITION.SP_LOWER Then
            e.StyleIndex = "TimeLineTiers"
            ActiveGanttVBECtl1.CustomTierDrawEventArgs.Text = e.StartDate.ToString("MMM")
        ElseIf e.TierPosition = E_TIERPOSITION.SP_UPPER Then
            e.StyleIndex = "TimeLineTiers"
            ActiveGanttVBECtl1.CustomTierDrawEventArgs.Text = e.StartDate.Year & " Q" & e.StartDate.Quarter
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_NodeChecked(sender As Object, e As AGVBE.NodeEventArgs) Handles ActiveGanttVBECtl1.NodeChecked
        Dim oRow As clsRow
        oRow = ActiveGanttVBECtl1.Rows.Item(e.Index.ToString())
        If oRow.Node.Checked = True Then
            oRow.ClientAreaStyleIndex = "ClientAreaChecked"
            oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumnChecked"
            oRow.Cells.Item("3").StyleIndex = "CellStyleChecked"
            oRow.Cells.Item("4").StyleIndex = "CellStyleChecked"
            If oRow.Node.StyleIndex = "NodeBold" Then
                oRow.Node.StyleIndex = "NodeBoldChecked"
            Else
                oRow.Node.StyleIndex = "NodeRegularChecked"
            End If
        Else
            oRow.ClientAreaStyleIndex = "ClientAreaStyle"
            oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
            oRow.Cells.Item("3").StyleIndex = "CellStyle"
            oRow.Cells.Item("4").StyleIndex = "CellStyle"
            If oRow.Node.StyleIndex = "NodeBoldChecked" Then
                oRow.Node.StyleIndex = "NodeBold"
            Else
                oRow.Node.StyleIndex = "NodeRegular"
            End If
        End If
    End Sub

    Private Sub ActiveGanttVBECtl1_ControlMouseDown(sender As Object, e As AGVBE.MouseEventArgs) Handles ActiveGanttVBECtl1.ControlMouseDown
        '// TODO
    End Sub

    Private Sub ActiveGanttVBECtl1_ObjectAdded(sender As Object, e As AGVBE.ObjectAddedEventArgs) Handles ActiveGanttVBECtl1.ObjectAdded
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK, E_EVENTTARGET.EVT_MILESTONE
                Dim oTask As clsTask = Nothing
                oTask = GetTaskByRowKey(ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString()).RowKey)
                oTask.StartDate = ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString()).StartDate
                oTask.EndDate = ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString()).EndDate
                UpdateTask(oTask.Index)
                ActiveGanttVBECtl1.Tasks.Remove(e.TaskIndex.ToString())
            Case E_EVENTTARGET.EVT_PREDECESSOR
                ActiveGanttVBECtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).StyleIndex = "NormalTask"
                ActiveGanttVBECtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).WarningStyleIndex = "NormalTaskWarning"
                ActiveGanttVBECtl1.Predecessors.Item(e.PredecessorObjectIndex.ToString()).SelectedStyleIndex = "SelectedPredecessor"
                InsertPredecessor(e.PredecessorTaskKey, e.TaskKey, e.PredecessorType)
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectMove(sender As Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectMove
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                UpdateTask(e.Index)
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_CompleteObjectSize(sender As Object, e As AGVBE.ObjectStateChangedEventArgs) Handles ActiveGanttVBECtl1.CompleteObjectSize
        Select Case e.EventTarget
            Case E_EVENTTARGET.EVT_TASK
                UpdateTask(e.Index)
            Case E_EVENTTARGET.EVT_PERCENTAGE
                Dim lTaskIndex As Integer = 0
                lTaskIndex = ActiveGanttVBECtl1.Tasks.Item(ActiveGanttVBECtl1.Percentages.Item(e.Index.ToString()).TaskKey).Index
                UpdateTask(lTaskIndex)
        End Select
    End Sub

    Private Sub ActiveGanttVBECtl1_ToolTipOnMouseHover(sender As Object, e As AGVBE.ToolTipEventArgs) Handles ActiveGanttVBECtl1.ToolTipOnMouseHover

        If e.EventTarget = E_EVENTTARGET.EVT_TASK Or e.EventTarget = E_EVENTTARGET.EVT_SELECTEDTASK Or e.EventTarget = E_EVENTTARGET.EVT_PERCENTAGE Or e.EventTarget = E_EVENTTARGET.EVT_SELECTEDPERCENTAGE Then
            mp_oToolTip.Visibility = Windows.Visibility.Visible
            Dim oRow As clsRow
            Dim oTask As clsTask
            Dim oPercentage As clsPercentage
            Dim fPercentage As Single = 0
            oRow = ActiveGanttVBECtl1.Rows.Item(ActiveGanttVBECtl1.MathLib.GetRowIndexByPosition(e.Y).ToString())
            oTask = ActiveGanttVBECtl1.Tasks.Item(ActiveGanttVBECtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y).ToString())
            oPercentage = GetPercentageByTaskKey(oTask.Key)
            If Not oPercentage Is Nothing Then
                fPercentage = oPercentage.Percent * 100
            End If
            mp_oToolTipTitle.Text = oRow.Text
            mp_oToolTipImage.Source = oRow.Node.Image.Source
            mp_oToolTipText.Text = "Duration: " & ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_DAY, oTask.StartDate, oTask.EndDate) & " days" & vbCrLf & _
            "From: " & oTask.StartDate.ToString("ddd MMM d, yyyy") & " To " & oTask.EndDate.ToString("ddd MMM d, yyyy") & vbCrLf & _
            "Percent Completed: " & fPercentage.ToString("00.00") & "%"
            Return
        End If
        mp_oToolTip.Visibility = Windows.Visibility.Collapsed
    End Sub

    Private Sub ActiveGanttVBECtl1_ToolTipOnMouseMove(sender As Object, e As AGVBE.ToolTipEventArgs) Handles ActiveGanttVBECtl1.ToolTipOnMouseMove
        If e.Operation = E_OPERATION.EO_PERCENTAGESIZING Or e.Operation = E_OPERATION.EO_TASKMOVEMENT Or e.Operation = E_OPERATION.EO_TASKSTRETCHLEFT Or e.Operation = E_OPERATION.EO_TASKSTRETCHRIGHT Then
            mp_oToolTip.Visibility = Windows.Visibility.Visible
            Dim oRow As clsRow
            Dim oTask As clsTask = Nothing
            Dim oPercentage As clsPercentage
            Dim fPercentage As Single = 0
            oRow = ActiveGanttVBECtl1.Rows.Item(ActiveGanttVBECtl1.MathLib.GetRowIndexByPosition(e.Y).ToString())
            If ActiveGanttVBECtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y) >= 1 Then
                oTask = ActiveGanttVBECtl1.Tasks.Item(ActiveGanttVBECtl1.MathLib.GetTaskIndexByPosition(e.X, e.Y).ToString())
            End If
            If oTask Is Nothing Then
                oTask = ActiveGanttVBECtl1.Tasks.Item(e.TaskIndex.ToString())
            End If
            oPercentage = GetPercentageByTaskKey(oTask.Key)
            If e.Operation = E_OPERATION.EO_PERCENTAGESIZING Then
                fPercentage = System.Convert.ToSingle((e.X - e.XStart) / (e.XEnd - e.XStart) * 100)
            Else
                If Not oPercentage Is Nothing Then
                    fPercentage = oPercentage.Percent * 100
                End If
            End If
            mp_oToolTipTitle.Text = oRow.Text
            mp_oToolTipImage.Source = oRow.Node.Image.Source
            mp_oToolTipText.Text = "Duration: " & ActiveGanttVBECtl1.MathLib.DateTimeDiff(E_INTERVAL.IL_DAY, e.StartDate, e.EndDate) & " days" & vbCrLf & _
            "From: " & e.StartDate.ToString("ddd MMM d, yyyy") & " To " & e.EndDate.ToString("ddd MMM d, yyyy") & vbCrLf & _
            "Percent Completed: " & fPercentage.ToString("00.00") & "%"
            Return
        End If
        mp_oToolTip.Visibility = Windows.Visibility.Collapsed
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

    Private Sub ActiveGanttVBECtl1_MouseMove(sender As Object, e As System.Windows.Input.MouseEventArgs) Handles ActiveGanttVBECtl1.MouseMove
        mp_oToolTip.SetValue(Canvas.LeftProperty, e.GetPosition(ActiveGanttVBECtl1.ControlCanvas).X + 32)
        mp_oToolTip.SetValue(Canvas.TopProperty, e.GetPosition(ActiveGanttVBECtl1.ControlCanvas).Y + 32)
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

#End Region

#Region "Functions"

    Private Sub UpdateTask(ByVal Index As Integer)
        Dim oPercentage As AGVBE.clsPercentage = GetPercentageByTaskKey(ActiveGanttVBECtl1.Tasks.Item(Index.ToString()).Key)
        Dim oTask As clsTask
        oTask = ActiveGanttVBECtl1.Tasks.Item(Index.ToString())
        SetTaskGridColumns(oTask)
        Dim sRowKey As String = oTask.RowKey
        Dim dtStartDate As AGVBE.DateTime = oTask.StartDate
        Dim dtEndDate As AGVBE.DateTime = oTask.EndDate
        Dim oNode As clsNode = ActiveGanttVBECtl1.Rows.Item(sRowKey).Node
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            '    Dim oCmd As OleDbCommand = Nothing
            '    Dim sSQL As String = "UPDATE tb_GuysStThomas SET " & _
            '    "StartDate = " & g_DST_ACCESS_ConvertDate(dtStartDate) & _
            '    ", EndDate = " & g_DST_ACCESS_ConvertDate(dtEndDate) & _
            '    ", PercentCompleted = " & oPercentage.Percent & _
            '    " WHERE ID = " & sRowKey.Replace("K", "")
            '    oConn.Open()
            '    oCmd = New OleDbCommand(sSQL, oConn)
            '    oCmd.ExecuteNonQuery()
            '    oConn.Close()
            'End Using
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            'Dim oDataRow As DataRow = Nothing
            'oDataRow = mp_otb_GuysStThomas.Tables(1).Rows.Find(sRowKey.Replace("K", ""))
            'oDataRow("StartDate") = dtStartDate
            'oDataRow("EndDate") = dtEndDate
            'oDataRow("PercentCompleted") = oPercentage.Percent
            'mp_otb_GuysStThomas.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xml")
        End If
        UpdateSummary(oNode)
    End Sub

    Private Sub InsertPredecessor(ByVal PredecessorKey As String, ByVal SuccessorKey As String, ByVal PredecessorType As E_CONSTRAINTTYPE)
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'Using oConn As New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            '    Dim oCmd As OleDbCommand = Nothing
            '    PredecessorKey = PredecessorKey.Replace("T", "")
            '    SuccessorKey = SuccessorKey.Replace("T", "")
            '    Dim sSQL As String = "INSERT INTO tb_GuysStThomas_Predecessors (lPredecessorID, lSuccessorID, yType) VALUES (" & PredecessorKey.Replace("T", "") & "," & SuccessorKey.Replace("T", "") & "," & PredecessorType & ")"
            '    oConn.Open()
            '    oCmd = New OleDbCommand(sSQL, oConn)
            '    oCmd.ExecuteNonQuery()
            '    oConn.Close()
            'End Using
        ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
            'Dim oDataRow As DataRow = Nothing
            'Dim oLastRow As DataRow = Nothing
            'oLastRow = mp_otb_GuysStThomas_Predecessors.Tables(1).Rows(mp_otb_GuysStThomas_Predecessors.Tables(1).Rows.Count - 1)
            'oDataRow = mp_otb_GuysStThomas_Predecessors.Tables(1).NewRow()
            'oDataRow("lID") = DirectCast(oLastRow.Item("ID"), System.Int32) + 1
            'oDataRow("lPredecessorID") = PredecessorKey.Replace("T", "")
            'oDataRow("lSuccessorID") = SuccessorKey.Replace("T", "")
            'oDataRow("yType") = PredecessorType
            'mp_otb_GuysStThomas_Predecessors.Tables(1).Rows.Add(oDataRow)
            'mp_otb_GuysStThomas_Predecessors.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas_Predecessors.xml")
        End If
    End Sub

    Private Sub UpdateSummary(ByRef oNode As clsNode)
        'Dim oConn As OleDbConnection = Nothing
        'Dim oCmd As OleDbCommand = Nothing
        Dim sSQL As String = ""
        Dim oParentNode As clsNode = Nothing
        Dim oSummaryTask As clsTask = Nothing
        Dim oSummaryPercentage As clsPercentage = Nothing
        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'oConn = New OleDbConnection(g_DST_ACCESS_GetConnectionString())
            'oConn.Open()
        End If
        oParentNode = oNode.Parent
        While Not oParentNode Is Nothing
            oSummaryTask = GetTaskByRowKey(oParentNode.Row.Key)
            oSummaryPercentage = GetPercentageByTaskKey(oSummaryTask.Key)
            If Not oSummaryTask Is Nothing Then
                Dim oChildTask As clsTask = Nothing
                Dim oChildPercentage As clsPercentage = Nothing
                Dim oChildNode As clsNode = Nothing
                Dim dtSumStartDate As AGVBE.DateTime = New AGVBE.DateTime()
                Dim dtSumEndDate As AGVBE.DateTime = New AGVBE.DateTime()
                Dim lPercentagesCount As Integer = 0
                Dim fPercentagesSum As Single = 0
                Dim fPercentageAvg As Single = 0
                oChildNode = oParentNode.Child
                While Not oChildNode Is Nothing
                    oChildTask = GetTaskByRowKey(oChildNode.Row.Key)
                    oChildPercentage = GetPercentageByTaskKey(oChildTask.Key)
                    lPercentagesCount = lPercentagesCount + 1
                    fPercentagesSum = fPercentagesSum + oChildPercentage.Percent
                    If Not oChildTask Is Nothing Then
                        If dtSumStartDate.DateTimePart.Ticks() = 0 Then
                            dtSumStartDate = oChildTask.StartDate
                        Else
                            If oChildTask.StartDate < dtSumStartDate Then
                                dtSumStartDate = oChildTask.StartDate
                            End If
                        End If
                        If dtSumEndDate.DateTimePart.Ticks() = 0 Then
                            dtSumEndDate = oChildTask.EndDate
                        Else
                            If oChildTask.EndDate > dtSumEndDate Then
                                dtSumEndDate = oChildTask.EndDate
                            End If
                        End If
                    End If
                    oChildNode = oChildNode.NextSibling
                End While
                fPercentageAvg = fPercentagesSum / lPercentagesCount
                oSummaryTask.StartDate = dtSumStartDate
                oSummaryTask.EndDate = dtSumEndDate
                SetTaskGridColumns(oSummaryTask)
                oSummaryPercentage.Percent = fPercentageAvg
                If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
                    'sSQL = "UPDATE tb_GuysStThomas SET " & _
                    '"StartDate = " & g_DST_ACCESS_ConvertDate(dtSumStartDate) & _
                    '", EndDate = " & g_DST_ACCESS_ConvertDate(dtSumEndDate) & _
                    '", PercentCompleted = " & oSummaryPercentage.Percent & _
                    '" WHERE ID = " & oSummaryTask.RowKey.Replace("K", "")
                    'oCmd = New OleDbCommand(sSQL, oConn)
                    'oCmd.ExecuteNonQuery()
                ElseIf mp_yDataSourceType = E_DATASOURCETYPE.DST_XML Then
                    'Dim oDataRow As DataRow = Nothing
                    'oDataRow = mp_otb_GuysStThomas.Tables(1).Rows.Find(oSummaryTask.RowKey.Replace("K", ""))
                    'oDataRow("StartDate") = dtSumStartDate
                    'oDataRow("EndDate") = dtSumEndDate
                    'oDataRow("PercentCompleted") = oSummaryPercentage.Percent
                    'mp_otb_GuysStThomas.WriteXml(g_GetAppLocation() & "\HPM_XML\tb_GuysStThomas.xml")
                End If
            End If
            oParentNode = oParentNode.Parent
        End While

        If mp_yDataSourceType = E_DATASOURCETYPE.DST_ACCESS Then
            'oConn.Close()
        End If

    End Sub

    Private Function GetTaskByRowKey(ByVal sRowKey As String) As clsTask
        Dim i As Integer = 0
        Dim oTask As clsTask = Nothing
        For i = 1 To ActiveGanttVBECtl1.Tasks.Count
            oTask = ActiveGanttVBECtl1.Tasks.Item(i.ToString())
            If oTask.RowKey = sRowKey Then
                Return oTask
            End If
        Next
        Return Nothing
    End Function

    Private Function GetPercentageByTaskKey(ByVal sTaskKey As String) As clsPercentage
        Dim i As Integer
        Dim oPercentage As clsPercentage
        For i = 1 To ActiveGanttVBECtl1.Percentages.Count
            oPercentage = ActiveGanttVBECtl1.Percentages.Item(i.ToString())
            If oPercentage.TaskKey = sTaskKey Then
                Return oPercentage
            End If
        Next
        Return Nothing
    End Function

    Private Function GetImage(ByVal sImage As String) As Image
        Dim oReturnImage = New Image
        Dim oURI As System.Uri = New System.Uri("AGVBECON;component/Images/WBS/" & sImage, UriKind.Relative)
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

#Region "cmdSave"

    Private Sub cmdSaveXML_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdSaveXML.Click
        mp_fSave = New fSave()
        mp_fSave.sSuggestedFileName = "AGVBE_WBSP"
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
        If CType(ActiveGanttVBECtl1.CurrentView, System.Int32) < 3 Then
            ActiveGanttVBECtl1.CurrentView = CType(CType(ActiveGanttVBECtl1.CurrentView, System.Int32) + 1, System.String)
            ActiveGanttVBECtl1.Redraw()
        End If
    End Sub

    Private Sub cmdZoomout_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdZoomout.Click
        If CType(ActiveGanttVBECtl1.CurrentView, System.Int32) > 1 Then
            ActiveGanttVBECtl1.CurrentView = CType(CType(ActiveGanttVBECtl1.CurrentView, System.Int32) - 1, System.String)
            ActiveGanttVBECtl1.Redraw()
        End If
    End Sub

    Private Sub cmdBluePercentages_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdBluePercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bBluePercentagesVisible = Not mp_bBluePercentagesVisible
        For i = 1 To ActiveGanttVBECtl1.Percentages.Count
            oPercentage = ActiveGanttVBECtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "BluePercentages" Then
                oPercentage.Visible = mp_bBluePercentagesVisible
            End If
        Next
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdGreenPercentages_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdGreenPercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bGreenPercentagesVisible = Not mp_bGreenPercentagesVisible
        For i = 1 To ActiveGanttVBECtl1.Percentages.Count
            oPercentage = ActiveGanttVBECtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "GreenPercentages" Then
                oPercentage.Visible = mp_bGreenPercentagesVisible
            End If
        Next
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdRedPercentages_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdRedPercentages.Click
        Dim i As Integer
        Dim oPercentage As clsPercentage
        mp_bRedPercentagesVisible = Not mp_bRedPercentagesVisible
        For i = 1 To ActiveGanttVBECtl1.Percentages.Count
            oPercentage = ActiveGanttVBECtl1.Percentages.Item(i.ToString())
            If oPercentage.StyleIndex = "RedPercentages" Then
                oPercentage.Visible = mp_bRedPercentagesVisible
            End If
        Next
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdProperties_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdProperties.Click
        mp_fWBSPProperties.Show()
    End Sub

    Private Sub cmdCheckPredecessors_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdCheckPredecessors.Click
        ActiveGanttVBECtl1.CheckPredecessors()
        ActiveGanttVBECtl1.Redraw()
    End Sub

    Private Sub cmdHelp_Click(sender As Object, e As System.Windows.RoutedEventArgs) Handles cmdHelp.Click
        System.Windows.Browser.HtmlPage.Window.Navigate(New Uri("http://www.sourcecodestore.com/Article.aspx?ID=17"), "_blank")
    End Sub

#End Region

#Region "Load Data"

    Public Sub NoDataSource_LoadTasks()
        AddRow_Task(1, 0, "A", "Capital Plan", New AGVBE.DateTime(2007, 3, 8, 12, 0, 0), New AGVBE.DateTime(2007, 10, 19, 0, 0, 0), 0.4, False, True)
        AddRow_Task(2, 0, "F", "Strategic Projects", New AGVBE.DateTime(2006, 11, 1, 12, 0, 0), New AGVBE.DateTime(2007, 9, 14, 0, 0, 0), 0.75, True, True)
        AddRow_Task(3, 1, "F", "Infrastructure Work Team", New AGVBE.DateTime(2007, 2, 1, 12, 0, 0), New AGVBE.DateTime(2007, 9, 5, 0, 0, 0), 0.77, True, True)
        AddRow_Task(4, 2, "A", "Guys Tower Façade Feasability", New AGVBE.DateTime(2007, 2, 1, 12, 0, 0), New AGVBE.DateTime(2007, 8, 1, 0, 0, 0), 0.6, False, True)
        AddRow_Task(5, 2, "A", "East Wing Cladding (inc Ward Refurbisments)", New AGVBE.DateTime(2007, 4, 21, 0, 0, 0), New AGVBE.DateTime(2007, 9, 5, 0, 0, 0), 0.94, False, True)
        AddRow_Task(6, 1, "F", "Modernisation Workstream", New AGVBE.DateTime(2007, 1, 22, 0, 0, 0), New AGVBE.DateTime(2007, 3, 27, 12, 0, 0), 0.72, True, True)
        AddRow_Task(7, 2, "A", "A&E Reconfiguration", New AGVBE.DateTime(2007, 1, 22, 0, 0, 0), New AGVBE.DateTime(2007, 3, 27, 12, 0, 0), 0.69, False, True)
        AddRow_Task(8, 2, "A", "St. Thomas Main Theatres Study", New AGVBE.DateTime(2007, 1, 28, 0, 0, 0), New AGVBE.DateTime(2007, 3, 18, 12, 0, 0), 0.75, False, True)
        AddRow_Task(9, 1, "F", "Ambulatory Workstream", New AGVBE.DateTime(2007, 3, 9, 12, 0, 0), New AGVBE.DateTime(2007, 6, 5, 12, 0, 0), 0.73, True, True)
        AddRow_Task(10, 2, "A", "PET Feasability", New AGVBE.DateTime(2007, 3, 9, 12, 0, 0), New AGVBE.DateTime(2007, 6, 5, 12, 0, 0), 0.73, False, True)
        AddRow_Task(11, 1, "F", "Cancer Workstream", New AGVBE.DateTime(2006, 11, 1, 12, 0, 0), New AGVBE.DateTime(2007, 9, 14, 0, 0, 0), 0.78, True, True)
        AddRow_Task(12, 2, "A", "Redevelopment of Guys Site Incorporating Cancer Feasability", New AGVBE.DateTime(2007, 1, 11, 0, 0, 0), New AGVBE.DateTime(2007, 8, 11, 12, 0, 0), 0.74, False, True)
        AddRow_Task(13, 2, "A", "Radiotherapy and Chemotherapy Center", New AGVBE.DateTime(2006, 11, 1, 12, 0, 0), New AGVBE.DateTime(2007, 3, 30, 12, 0, 0), 0.94, False, True)
        AddRow_Task(14, 2, "A", "Decant Facilities", New AGVBE.DateTime(2007, 5, 24, 12, 0, 0), New AGVBE.DateTime(2007, 9, 14, 0, 0, 0), 0.65, False, True)
        AddRow_Task(15, 0, "F", "Capital Projects", New AGVBE.DateTime(2006, 9, 1, 12, 0, 0), New AGVBE.DateTime(2007, 12, 12, 0, 0, 0), 0.87, True, True)
        AddRow_Task(16, 1, "A", "4th Floor Block & Refurbishment", New AGVBE.DateTime(2006, 9, 1, 12, 0, 0), New AGVBE.DateTime(2007, 2, 1, 0, 0, 0), 0.93, False, True)
        AddRow_Task(17, 1, "A", "Bio Medical Research Center & CRF", New AGVBE.DateTime(2007, 3, 2, 0, 0, 0), New AGVBE.DateTime(2007, 7, 4, 0, 0, 0), 0.91, False, True)
        AddRow_Task(18, 1, "A", "Blundell Ward Relocation Florence + Aston Key", New AGVBE.DateTime(2007, 8, 7, 12, 0, 0), New AGVBE.DateTime(2007, 11, 12, 12, 0, 0), 0.62, False, True)
        AddRow_Task(19, 1, "A", "Bostock Ward Replacement of Water Treatment Plant", New AGVBE.DateTime(2007, 3, 7, 0, 0, 0), New AGVBE.DateTime(2007, 6, 23, 12, 0, 0), 0.84, False, True)
        AddRow_Task(20, 1, "A", "Centralisation Health Record Storage", New AGVBE.DateTime(2007, 6, 22, 0, 0, 0), New AGVBE.DateTime(2007, 11, 12, 0, 0, 0), 0.78, False, True)
        AddRow_Task(21, 1, "A", "ENT & Audiology Suite Phase II", New AGVBE.DateTime(2006, 12, 31, 12, 0, 0), New AGVBE.DateTime(2007, 3, 10, 0, 0, 0), 0.75, False, True)
        AddRow_Task(22, 1, "A", "GLI Structural Monitoring & Repair", New AGVBE.DateTime(2007, 2, 12, 12, 0, 0), New AGVBE.DateTime(2007, 5, 9, 12, 0, 0), 0.91, False, True)
        AddRow_Task(23, 1, "A", "Pathology Labs (Phase 1A)", New AGVBE.DateTime(2007, 4, 2, 0, 0, 0), New AGVBE.DateTime(2007, 10, 23, 0, 0, 0), 0.95, False, True)
        AddRow_Task(24, 1, "A", "Pathology Labs (Phase 2)", New AGVBE.DateTime(2007, 1, 15, 0, 0, 0), New AGVBE.DateTime(2007, 7, 29, 12, 0, 0), 0.92, False, True)
        AddRow_Task(25, 1, "A", "Pathology: NW5 - CSR Haematology & CSR Labs", New AGVBE.DateTime(2007, 4, 9, 0, 0, 0), New AGVBE.DateTime(2007, 9, 5, 0, 0, 0), 0.88, False, True)
        AddRow_Task(26, 1, "A", "Pathology: Haematology Day Care Center Transfer (NW4 to GT4)", New AGVBE.DateTime(2006, 10, 19, 0, 0, 0), New AGVBE.DateTime(2007, 1, 12, 0, 0, 0), 0.85, False, True)
        AddRow_Task(27, 1, "A", "HDR", New AGVBE.DateTime(2007, 6, 1, 0, 0, 0), New AGVBE.DateTime(2007, 9, 3, 0, 0, 0), 0.85, False, True)
        AddRow_Task(28, 1, "A", "Kidney Treatment Center", New AGVBE.DateTime(2007, 6, 25, 0, 0, 0), New AGVBE.DateTime(2007, 11, 18, 0, 0, 0), 0.76, False, True)
        AddRow_Task(29, 1, "A", "Maternity Expansion Business Case", New AGVBE.DateTime(2006, 11, 9, 12, 0, 0), New AGVBE.DateTime(2007, 4, 6, 0, 0, 0), 0.93, False, True)
        AddRow_Task(30, 1, "A", "New Laminar Flow Theatre at Guy's", New AGVBE.DateTime(2007, 4, 25, 12, 0, 0), New AGVBE.DateTime(2007, 11, 29, 12, 0, 0), 0.89, False, True)
        AddRow_Task(31, 1, "A", "North Wing Basement Entance - Phase 2", New AGVBE.DateTime(2007, 9, 7, 0, 0, 0), New AGVBE.DateTime(2007, 11, 30, 0, 0, 0), 0.88, False, True)
        AddRow_Task(32, 1, "A", "Paediatric Neurosciences Feasibility", New AGVBE.DateTime(2006, 11, 29, 0, 0, 0), New AGVBE.DateTime(2007, 2, 10, 0, 0, 0), 0.9, False, True)
        AddRow_Task(33, 1, "A", "Fluroscopy (Imaging 2) at St. Thomas", New AGVBE.DateTime(2007, 1, 24, 0, 0, 0), New AGVBE.DateTime(2007, 6, 8, 12, 0, 0), 0.94, False, True)
        AddRow_Task(34, 1, "A", "Interventional Radiology Suite (Imaging 3) at GT3 Phase 1", New AGVBE.DateTime(2007, 6, 17, 0, 0, 0), New AGVBE.DateTime(2007, 12, 12, 0, 0, 0), 0.91, False, True)
        AddRow_Task(35, 1, "A", "Interventional Radiology Suite (Imaging 3) at GT3 Phase 2", New AGVBE.DateTime(2007, 8, 12, 0, 0, 0), New AGVBE.DateTime(2007, 12, 1, 12, 0, 0), 0.92, False, True)
        AddRow_Task(36, 1, "A", "Imaging: Radiology Environment & Waiting Areas (Imaging 2) Phases 1 & 2", New AGVBE.DateTime(2006, 11, 27, 12, 0, 0), New AGVBE.DateTime(2007, 1, 25, 12, 0, 0), 1.0, False, True)
        AddRow_Task(37, 1, "A", "Imaging: Radiology Environment & Waiting Areas (Imaging 2) Phase 3", New AGVBE.DateTime(2006, 12, 21, 0, 0, 0), New AGVBE.DateTime(2007, 1, 9, 0, 0, 0), 1.0, False, True)
        AddRow_Task(38, 1, "A", "Relocation of Pharmacy Manufacturing & QC Laboratories", New AGVBE.DateTime(2007, 6, 7, 12, 0, 0), New AGVBE.DateTime(2007, 8, 20, 12, 0, 0), 0.93, False, True)
        AddRow_Task(39, 1, "A", "Samaritan Ward - Bone marrow transplant beds", New AGVBE.DateTime(2007, 6, 1, 0, 0, 0), New AGVBE.DateTime(2007, 8, 18, 0, 0, 0), 0.94, False, True)
        AddRow_Task(40, 1, "A", "Sexual Health Relocation", New AGVBE.DateTime(2007, 1, 10, 12, 0, 0), New AGVBE.DateTime(2007, 4, 12, 12, 0, 0), 1.0, False, True)
        AddRow_Task(41, 1, "A", "St. Thomas HV Upgrade", New AGVBE.DateTime(2007, 5, 2, 12, 0, 0), New AGVBE.DateTime(2007, 6, 20, 12, 0, 0), 0.52, False, True)
        AddRow_Task(42, 1, "A", "Ultrasound (Imaging 2) at Guy's", New AGVBE.DateTime(2007, 6, 5, 12, 0, 0), New AGVBE.DateTime(2007, 6, 22, 12, 0, 0), 1.0, False, True)
        AddRow_Task(43, 1, "F", "New Schemes Approved in Year", New AGVBE.DateTime(2006, 11, 15, 12, 0, 0), New AGVBE.DateTime(2007, 9, 4, 12, 0, 0), 0.78, True, True)
        AddRow_Task(44, 2, "A", "Modular Theatres", New AGVBE.DateTime(2006, 11, 15, 12, 0, 0), New AGVBE.DateTime(2007, 1, 1, 12, 0, 0), 0.84, False, True)
        AddRow_Task(45, 2, "A", "ECH - Theatre Ventilation", New AGVBE.DateTime(2006, 12, 24, 0, 0, 0), New AGVBE.DateTime(2007, 9, 4, 12, 0, 0), 0.77, False, True)
        AddRow_Task(46, 2, "A", "Modular Pharmacy Aseptic Unit", New AGVBE.DateTime(2006, 12, 22, 12, 0, 0), New AGVBE.DateTime(2007, 1, 28, 12, 0, 0), 0.82, False, True)
        AddRow_Task(47, 2, "A", "Acute Stroke Unit Bid", New AGVBE.DateTime(2007, 4, 11, 0, 0, 0), New AGVBE.DateTime(2007, 7, 20, 0, 0, 0), 0.74, False, True)
        AddRow_Task(48, 2, "A", "Chemo Centralisation", New AGVBE.DateTime(2006, 12, 26, 0, 0, 0), New AGVBE.DateTime(2007, 3, 30, 0, 0, 0), 0.9, False, True)
        AddRow_Task(49, 2, "A", "Feasability of MRI at Guy's", New AGVBE.DateTime(2007, 5, 12, 0, 0, 0), New AGVBE.DateTime(2007, 7, 25, 0, 0, 0), 0.59, False, True)
        AddRow_Task(50, 0, "F", "Engineering", New AGVBE.DateTime(2006, 10, 17, 0, 0, 0), New AGVBE.DateTime(2007, 9, 15, 12, 0, 0), 0.7, True, True)
        AddRow_Task(51, 1, "A", "Borough Wing Theatre Ductwork and Heater Batteries", New AGVBE.DateTime(2007, 5, 2, 0, 0, 0), New AGVBE.DateTime(2007, 6, 20, 0, 0, 0), 0.85, False, True)
        AddRow_Task(52, 1, "A", "Combined Heat and Power System at Guy's", New AGVBE.DateTime(2007, 1, 20, 12, 0, 0), New AGVBE.DateTime(2007, 4, 15, 12, 0, 0), 0.88, False, True)
        AddRow_Task(53, 1, "A", "Combined Heat and Power System at St. Thomas", New AGVBE.DateTime(2007, 3, 10, 12, 0, 0), New AGVBE.DateTime(2007, 9, 15, 12, 0, 0), 0.74, False, True)
        AddRow_Task(54, 1, "A", "Electrical Power Monitoring", New AGVBE.DateTime(2006, 11, 20, 0, 0, 0), New AGVBE.DateTime(2007, 8, 22, 12, 0, 0), 0.88, False, True)
        AddRow_Task(55, 1, "A", "Guy's Lifts 101-105 (Guys Tower)", New AGVBE.DateTime(2006, 12, 6, 0, 0, 0), New AGVBE.DateTime(2007, 3, 3, 0, 0, 0), 0.88, False, True)
        AddRow_Task(56, 1, "A", "Guy's Lifts 110-114 (Guys Tower)", New AGVBE.DateTime(2007, 5, 15, 12, 0, 0), New AGVBE.DateTime(2007, 7, 1, 12, 0, 0), 0.5, False, True)
        AddRow_Task(57, 1, "A", "Motor Control Panel Refurbishment", New AGVBE.DateTime(2007, 1, 9, 0, 0, 0), New AGVBE.DateTime(2007, 6, 13, 0, 0, 0), 0.7, False, True)
        AddRow_Task(58, 1, "A", "North Wing / Lambeth Wing Air Supply Plants", New AGVBE.DateTime(2007, 1, 13, 0, 0, 0), New AGVBE.DateTime(2007, 4, 19, 0, 0, 0), 0.21, False, True)
        AddRow_Task(59, 1, "A", "North Wing Chiller Replacement", New AGVBE.DateTime(2007, 1, 9, 0, 0, 0), New AGVBE.DateTime(2007, 6, 16, 0, 0, 0), 0.5, False, True)
        AddRow_Task(60, 1, "A", "North Wing Replacement Generator", New AGVBE.DateTime(2006, 12, 10, 12, 0, 0), New AGVBE.DateTime(2007, 6, 11, 0, 0, 0), 0.76, False, True)
        AddRow_Task(61, 1, "A", "NW/LW Riser Refurbishment", New AGVBE.DateTime(2007, 1, 20, 12, 0, 0), New AGVBE.DateTime(2007, 3, 17, 12, 0, 0), 0.5, False, True)
        AddRow_Task(62, 1, "A", "Satchwell BMS Upgrade", New AGVBE.DateTime(2006, 12, 16, 12, 0, 0), New AGVBE.DateTime(2007, 7, 18, 12, 0, 0), 0.91, False, True)
        AddRow_Task(63, 1, "A", "St. Thomas Increase Standby Capacity - Phase 2", New AGVBE.DateTime(2007, 1, 2, 0, 0, 0), New AGVBE.DateTime(2007, 6, 18, 0, 0, 0), 0.8, False, True)
        AddRow_Task(64, 1, "A", "Substation 3 HV Works (St. Thomas)", New AGVBE.DateTime(2007, 2, 27, 0, 0, 0), New AGVBE.DateTime(2007, 8, 10, 12, 0, 0), 0.78, False, True)
        AddRow_Task(65, 1, "A", "TB Electrical Distribution", New AGVBE.DateTime(2006, 10, 17, 0, 0, 0), New AGVBE.DateTime(2007, 6, 29, 12, 0, 0), 0.73, False, True)
        AddRow_Task(66, 1, "A", "Tower Wing Dental Theatre Air Handling Unit", New AGVBE.DateTime(2006, 12, 30, 12, 0, 0), New AGVBE.DateTime(2007, 3, 24, 12, 0, 0), 0.75, False, True)
        AddRow_Task(67, 1, "A", "Tower Wing Recovery Air Handling Unit", New AGVBE.DateTime(2007, 3, 2, 0, 0, 0), New AGVBE.DateTime(2007, 8, 8, 0, 0, 0), 0.7, False, True)
        AddRow_Task(68, 1, "A", "Water Booster Pumps - Phase 1 & 2", New AGVBE.DateTime(2007, 1, 8, 12, 0, 0), New AGVBE.DateTime(2007, 6, 14, 12, 0, 0), 0.64, False, True)
        AddRow_Task(69, 1, "A", "Water Softner - Boiler House", New AGVBE.DateTime(2007, 2, 12, 12, 0, 0), New AGVBE.DateTime(2007, 7, 30, 12, 0, 0), 0.66, False, True)
        AddRow_Task(70, 1, "A", "Energy Efficiency", New AGVBE.DateTime(2007, 3, 31, 12, 0, 0), New AGVBE.DateTime(2007, 9, 4, 12, 0, 0), 0.72, False, True)
        AddRow_Task(71, 0, "F", "PEAT Plan", New AGVBE.DateTime(2006, 11, 5, 0, 0, 0), New AGVBE.DateTime(2008, 1, 21, 0, 0, 0), 0.82, True, True)
        AddRow_Task(72, 1, "A", "Hilliers Ward Refurb St. Thomas", New AGVBE.DateTime(2007, 3, 28, 0, 0, 0), New AGVBE.DateTime(2007, 5, 23, 12, 0, 0), 0.79, False, True)
        AddRow_Task(73, 1, "A", "William Gull Ward St. Thomas", New AGVBE.DateTime(2007, 3, 20, 0, 0, 0), New AGVBE.DateTime(2007, 8, 23, 0, 0, 0), 0.77, False, True)
        AddRow_Task(74, 1, "A", "Henry Ward Day Room", New AGVBE.DateTime(2007, 4, 29, 0, 0, 0), New AGVBE.DateTime(2007, 6, 1, 0, 0, 0), 0.8, False, True)
        AddRow_Task(75, 1, "A", "Sarah Swift Ward", New AGVBE.DateTime(2006, 11, 5, 0, 0, 0), New AGVBE.DateTime(2007, 2, 3, 0, 0, 0), 0.78, False, True)
        AddRow_Task(76, 1, "A", "Victoria Ward", New AGVBE.DateTime(2007, 5, 10, 12, 0, 0), New AGVBE.DateTime(2007, 7, 14, 12, 0, 0), 0.91, False, True)
        AddRow_Task(77, 1, "A", "Appointment Center Staff Toilets", New AGVBE.DateTime(2007, 1, 16, 0, 0, 0), New AGVBE.DateTime(2007, 4, 7, 12, 0, 0), 0.77, False, True)
        AddRow_Task(78, 1, "A", "Page Ward", New AGVBE.DateTime(2007, 5, 19, 12, 0, 0), New AGVBE.DateTime(2007, 7, 16, 12, 0, 0), 0.74, False, True)
        AddRow_Task(79, 1, "A", "Nightingdale Ward - Side Rooms", New AGVBE.DateTime(2007, 2, 18, 0, 0, 0), New AGVBE.DateTime(2007, 4, 28, 0, 0, 0), 0.77, False, True)
        AddRow_Task(80, 1, "A", "Luke Ward - Side Rooms", New AGVBE.DateTime(2007, 11, 14, 12, 0, 0), New AGVBE.DateTime(2007, 12, 31, 12, 0, 0), 0.8, False, True)
        AddRow_Task(81, 1, "A", "Therapies Department - Disabled Toilets", New AGVBE.DateTime(2007, 7, 31, 12, 0, 0), New AGVBE.DateTime(2007, 9, 26, 12, 0, 0), 0.81, False, True)
        AddRow_Task(82, 1, "A", "Northumberland Ward Side Rooms", New AGVBE.DateTime(2007, 4, 18, 0, 0, 0), New AGVBE.DateTime(2007, 6, 6, 0, 0, 0), 0.83, False, True)
        AddRow_Task(83, 1, "A", "General Outpatients", New AGVBE.DateTime(2007, 10, 17, 0, 0, 0), New AGVBE.DateTime(2008, 1, 21, 0, 0, 0), 0.86, False, True)
        AddRow_Task(84, 1, "A", "Rheumatology Clinic", New AGVBE.DateTime(2007, 5, 3, 0, 0, 0), New AGVBE.DateTime(2007, 5, 28, 0, 0, 0), 0.84, False, True)
        AddRow_Task(85, 1, "A", "Diabetes Clinic", New AGVBE.DateTime(2007, 1, 8, 12, 0, 0), New AGVBE.DateTime(2007, 3, 18, 12, 0, 0), 0.86, False, True)
        AddRow_Task(86, 1, "A", "ENT Clinic", New AGVBE.DateTime(2007, 4, 14, 12, 0, 0), New AGVBE.DateTime(2007, 10, 28, 12, 0, 0), 0.91, False, True)
        AddRow_Task(87, 0, "F", "Buildings Improvement Programs", New AGVBE.DateTime(2006, 10, 18, 12, 0, 0), New AGVBE.DateTime(2007, 10, 28, 0, 0, 0), 0.75, True, True)
        AddRow_Task(88, 1, "F", "Environmental Improvement Plan", New AGVBE.DateTime(2006, 10, 18, 12, 0, 0), New AGVBE.DateTime(2007, 10, 28, 0, 0, 0), 0.75, False, False)
        AddRow_Task(89, 2, "A", "Ward Improvementrs", New AGVBE.DateTime(2006, 10, 18, 12, 0, 0), New AGVBE.DateTime(2007, 10, 15, 12, 0, 0), 0.61, False, True)
        AddRow_Task(90, 2, "A", "Outpatient / Clinics", New AGVBE.DateTime(2006, 12, 29, 0, 0, 0), New AGVBE.DateTime(2007, 8, 11, 0, 0, 0), 0.74, False, True)
        AddRow_Task(91, 2, "A", "Circulation Areas", New AGVBE.DateTime(2007, 4, 14, 12, 0, 0), New AGVBE.DateTime(2007, 10, 28, 0, 0, 0), 0.74, False, True)
        AddRow_Task(92, 2, "A", "St. Thomas Main Entrance", New AGVBE.DateTime(2007, 2, 28, 0, 0, 0), New AGVBE.DateTime(2007, 6, 8, 0, 0, 0), 0.76, False, True)
        AddRow_Task(93, 2, "A", "St. Thomas Retail Mall", New AGVBE.DateTime(2007, 1, 1, 0, 0, 0), New AGVBE.DateTime(2007, 2, 6, 0, 0, 0), 0.81, False, True)
        AddRow_Task(94, 2, "A", "Guys Main Entrance Revolving Door", New AGVBE.DateTime(2007, 3, 28, 12, 0, 0), New AGVBE.DateTime(2007, 4, 25, 12, 0, 0), 0.83, False, True)

        AddPredecessor(16, 17, E_CONSTRAINTTYPE.PCT_END_TO_START, 696, E_INTERVAL.IL_HOUR)     '//End-To-Start with lag (down)
        AddPredecessor(13, 5, E_CONSTRAINTTYPE.PCT_END_TO_START, 516, E_INTERVAL.IL_HOUR)      '//End-To-Start with lag (up)
        AddPredecessor(21, 22, E_CONSTRAINTTYPE.PCT_END_TO_START, -612, E_INTERVAL.IL_HOUR)    '//End-To-Start with lead (down)
        AddPredecessor(24, 19, E_CONSTRAINTTYPE.PCT_END_TO_START, -3468, E_INTERVAL.IL_HOUR)   '//End-To-Start with lead (up)

        AddPredecessor(18, 20, E_CONSTRAINTTYPE.PCT_START_TO_END, 2316, E_INTERVAL.IL_HOUR)    '//Start-To-End with lag (down)
        AddPredecessor(29, 26, E_CONSTRAINTTYPE.PCT_START_TO_END, 1524, E_INTERVAL.IL_HOUR)    '//Start-To-End with lag (up)
        AddPredecessor(27, 32, E_CONSTRAINTTYPE.PCT_START_TO_END, -2664, E_INTERVAL.IL_HOUR)   '//Start-To-End with lead (down)
        AddPredecessor(38, 36, E_CONSTRAINTTYPE.PCT_START_TO_END, -3192, E_INTERVAL.IL_HOUR)   '//Start-To-End with lead (up)

        AddPredecessor(12, 14, E_CONSTRAINTTYPE.PCT_START_TO_START, 3204, E_INTERVAL.IL_HOUR)  '//Start-To-Start with lag (down)
        AddPredecessor(48, 47, E_CONSTRAINTTYPE.PCT_START_TO_START, 2544, E_INTERVAL.IL_HOUR)  '//Start-To-Start with lag (up)
        AddPredecessor(52, 55, E_CONSTRAINTTYPE.PCT_START_TO_START, -1092, E_INTERVAL.IL_HOUR) '//Start-To-Start with lead (down)
        AddPredecessor(56, 53, E_CONSTRAINTTYPE.PCT_START_TO_START, -1584, E_INTERVAL.IL_HOUR) '//Start-To-Start with lead (up)

        AddPredecessor(40, 41, E_CONSTRAINTTYPE.PCT_END_TO_END, 1656, E_INTERVAL.IL_HOUR)      '//End-To-End with lag (down)
        AddPredecessor(58, 57, E_CONSTRAINTTYPE.PCT_END_TO_END, 1320, E_INTERVAL.IL_HOUR)      '//End-To-End with lag (up)
        AddPredecessor(62, 63, E_CONSTRAINTTYPE.PCT_END_TO_END, -732, E_INTERVAL.IL_HOUR)      '//End-To-End with lead (down)
        AddPredecessor(67, 65, E_CONSTRAINTTYPE.PCT_END_TO_END, -948, E_INTERVAL.IL_HOUR)      '//End-To-End with lead (up)

    End Sub

    Public Sub AddPredecessor(ByVal lPredecessorID As Integer, ByVal lSuccessorID As Integer, ByVal yType As E_CONSTRAINTTYPE, ByVal lLagFactor As Integer, ByVal yLagInterval As E_INTERVAL)
        Dim oPredecessor As clsPredecessor
        oPredecessor = ActiveGanttVBECtl1.Predecessors.Add("T" & lSuccessorID.ToString(), "T" & lPredecessorID.ToString(), yType, "", "NormalTask")
        oPredecessor.WarningStyleIndex = "NormalTaskWarning"
        oPredecessor.SelectedStyleIndex = "SelectedPredecessor"
        oPredecessor.LagFactor = lLagFactor
        oPredecessor.LagInterval = yLagInterval
    End Sub

    Public Sub AddRow_Task(ByVal lID As Integer, ByVal lDepth As Integer, ByVal sTaskType As String, ByVal sDescription As String, ByVal dtStartDate As AGVBE.DateTime, ByVal dtEndDate As AGVBE.DateTime, ByVal fPercentCompleted As Single, ByVal bSummary As Boolean, ByVal bHasTasks As Boolean)
        Dim oRow As clsRow = Nothing
        Dim oTask As clsTask = Nothing
        oRow = ActiveGanttVBECtl1.Rows.Add("K" & lID.ToString(), sDescription)
        oRow.Cells.Item("1").Text = lID.ToString()
        oRow.Cells.Item("1").StyleIndex = "CellStyleKeyColumn"
        oRow.Node.StyleIndex = "CellStyle"
        oRow.Cells.Item("3").StyleIndex = "CellStyle"
        oRow.Cells.Item("4").StyleIndex = "CellStyle"
        oRow.Height = 20
        oRow.ClientAreaStyleIndex = "ClientAreaStyle"
        oRow.Node.AllowTextEdit = True
        If sTaskType = "F" Then
            If lDepth = 0 Then
                oRow.Node.Image = GetImage("folderclosed.png")
                oRow.Node.ExpandedImage = GetImage("folderopen.png")
                oRow.Node.StyleIndex = "NodeBold"
            Else
                oRow.Node.Image = GetImage("modules.png")
                oRow.Node.StyleIndex = "NodeRegular"
            End If
        ElseIf sTaskType = "A" Then
            oRow.Node.StyleIndex = "NodeRegular"
            oRow.Node.Image = GetImage("task.png")
            oRow.Node.CheckBoxVisible = True
        End If
        oRow.Node.Depth = lDepth
        oRow.Node.ImageVisible = True
        oRow.Node.AllowTextEdit = True
        If (mp_dtStartDate.DateTimePart.Ticks() = 0) Then
            mp_dtStartDate = dtStartDate
        Else
            If (dtStartDate < mp_dtStartDate) Then
                mp_dtStartDate = dtStartDate
            End If
        End If
        If (mp_dtEndDate.DateTimePart.Ticks() = 0) Then
            mp_dtEndDate = dtEndDate
        Else
            If (dtEndDate > mp_dtEndDate) Then
                mp_dtEndDate = dtEndDate
            End If
        End If
        oTask = ActiveGanttVBECtl1.Tasks.Add("", "K" & lID, dtStartDate, dtEndDate, "T" & lID.ToString())
        SetTaskGridColumns(oTask)
        If bSummary = True Then
            '// Prevent user from moving/sizing summary tasks
            oTask.AllowedMovement = E_MOVEMENTTYPE.MT_MOVEMENTDISABLED
            oTask.AllowStretchLeft = False
            oTask.AllowStretchRight = False
            '// Prevent user from adding tasks in these Rows
            oRow.Container = False
            '// Apply Summary Style 
            If oRow.Node.Depth = 0 Then
                oTask.StyleIndex = "RedSummary"
                ActiveGanttVBECtl1.Percentages.Add("T" & lID.ToString(), "RedPercentages", fPercentCompleted)
            ElseIf oRow.Node.Depth = 1 Then
                oTask.StyleIndex = "GreenSummary"
                ActiveGanttVBECtl1.Percentages.Add("T" & lID.ToString(), "GreenPercentages", fPercentCompleted)
            End If
            ActiveGanttVBECtl1.Percentages.Item(ActiveGanttVBECtl1.Percentages.Count.ToString()).AllowSize = False
        Else
            oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW
            oTask.StyleIndex = "NormalTask"
            oTask.WarningStyleIndex = "NormalTaskWarning"
            If bHasTasks = False Then
                oTask.Visible = False
                '// Prevent user from adding tasks in these rows
                oRow.Container = False
                ActiveGanttVBECtl1.Percentages.Add("T" & lID.ToString(), "InvisiblePercentages", fPercentCompleted)
                ActiveGanttVBECtl1.Percentages.Item(ActiveGanttVBECtl1.Percentages.Count.ToString()).AllowSize = False
            Else
                ActiveGanttVBECtl1.Percentages.Add("T" & lID.ToString(), "BluePercentages", fPercentCompleted)
            End If
        End If
    End Sub

    Private Sub SetTaskGridColumns(ByVal oTask As clsTask)
        oTask.Row.Cells.Item("3").Text = oTask.StartDate.ToString("MM/dd/yyyy")
        oTask.Row.Cells.Item("4").Text = oTask.EndDate.ToString("MM/dd/yyyy")
    End Sub

#End Region

End Class
