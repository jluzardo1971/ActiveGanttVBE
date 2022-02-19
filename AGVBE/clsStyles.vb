Option Explicit On

Public Class clsStyles

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oCollection As clsCollectionBase
    Private mp_oDefaultControlStyle As clsStyle
    Private mp_oDefaultTaskStyle As clsStyle
    Private mp_oDefaultRowStyle As clsStyle
    Private mp_oDefaultClientAreaStyle As clsStyle
    Private mp_oDefaultCellStyle As clsStyle
    Private mp_oDefaultColumnStyle As clsStyle
    Private mp_oDefaultPercentageStyle As clsStyle
    Private mp_oDefaultPredecessorStyle As clsStyle
    Private mp_oDefaultTimeLineStyle As clsStyle
    Private mp_oDefaultTimeBlockStyle As clsStyle
    Private mp_oDefaultTickMarkAreaStyle As clsStyle
    Private mp_oDefaultSplitterStyle As clsStyle
    Private mp_oDefaultNodeStyle As clsStyle
    Private mp_oDefaultTierStyle As clsStyle
    Private mp_oDefaultScrollBarStyle As clsStyle
    Private mp_oDefaultSBSeparator As clsStyle
    Private mp_oDefaultSBNormalStyle As clsStyle
    Private mp_oDefaultSBPressedStyle As clsStyle
    Private mp_oDefaultSBDisabledStyle As clsStyle

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_oCollection = New clsCollectionBase(Value, "Style")

        mp_oDefaultControlStyle = New clsStyle(Value)
        mp_oDefaultControlStyle.Appearance = E_STYLEAPPEARANCE.SA_SUNKEN
        mp_oDefaultControlStyle.BackColor = Colors.White

        mp_oDefaultTaskStyle = New clsStyle(Value)
        mp_oDefaultTaskStyle.MilestoneStyle.ShapeIndex = GRE_FIGURETYPE.FT_DIAMOND

        mp_oDefaultRowStyle = New clsStyle(Value)

        mp_oDefaultClientAreaStyle = New clsStyle(Value)
        mp_oDefaultClientAreaStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultClientAreaStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_TRANSPARENT
        mp_oDefaultClientAreaStyle.BorderColor = Colors.Gray
        mp_oDefaultClientAreaStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        mp_oDefaultClientAreaStyle.CustomBorderStyle.Top = False
        mp_oDefaultClientAreaStyle.CustomBorderStyle.Left = False
        mp_oDefaultClientAreaStyle.CustomBorderStyle.Right = False

        mp_oDefaultCellStyle = New clsStyle(Value)
        mp_oDefaultCellStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultCellStyle.BackColor = Colors.White
        mp_oDefaultCellStyle.BorderColor = Colors.Gray
        mp_oDefaultCellStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        mp_oDefaultCellStyle.CustomBorderStyle.Top = False
        mp_oDefaultCellStyle.CustomBorderStyle.Left = False
        mp_oDefaultCellStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        mp_oDefaultCellStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        mp_oDefaultCellStyle.TextYMargin = 5
        mp_oDefaultCellStyle.TextXMargin = 5
        mp_oDefaultCellStyle.Font = New Font("Tahoma", 8, E_FONTSIZEUNITS.FSU_POINTS, FontWeights.Bold)
        mp_oDefaultCellStyle.ImageAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT
        mp_oDefaultCellStyle.ImageAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP
        mp_oDefaultCellStyle.ImageXMargin = 0
        mp_oDefaultCellStyle.ImageYMargin = 0

        mp_oDefaultNodeStyle = New clsStyle(Value)
        mp_oDefaultNodeStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultNodeStyle.BackColor = Colors.White
        mp_oDefaultNodeStyle.BorderColor = Colors.Gray
        mp_oDefaultNodeStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM
        mp_oDefaultNodeStyle.CustomBorderStyle.Top = False
        mp_oDefaultNodeStyle.CustomBorderStyle.Left = False

        mp_oDefaultColumnStyle = New clsStyle(Value)
        mp_oDefaultColumnStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED
        mp_oDefaultColumnStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM
        mp_oDefaultColumnStyle.TextYMargin = 5
        mp_oDefaultColumnStyle.Font = New Font("Tahoma", 8, E_FONTSIZEUNITS.FSU_POINTS, FontWeights.Bold)

        mp_oDefaultPercentageStyle = New clsStyle(Value)
        mp_oDefaultPercentageStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultPercentageStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID
        mp_oDefaultPercentageStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT
        mp_oDefaultPercentageStyle.OffsetTop = 10
        mp_oDefaultPercentageStyle.OffsetBottom = 15
        mp_oDefaultPercentageStyle.BackColor = Color.FromArgb(0, 0, 0, 255)

        mp_oDefaultPredecessorStyle = New clsStyle(Value)
        mp_oDefaultTimeLineStyle = New clsStyle(Value)

        mp_oDefaultTimeBlockStyle = New clsStyle(Value)
        mp_oDefaultTimeBlockStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultTimeBlockStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH
        mp_oDefaultTimeBlockStyle.HatchBackColor = Colors.White
        mp_oDefaultTimeBlockStyle.HatchForeColor = Colors.Gray
        mp_oDefaultTimeBlockStyle.HatchStyle = GRE_HATCHSTYLE.HS_PERCENT50

        mp_oDefaultTickMarkAreaStyle = New clsStyle(Value)

        mp_oDefaultSplitterStyle = New clsStyle(Value)
        mp_oDefaultSplitterStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultSplitterStyle.BackColor = Colors.Black

        mp_oDefaultTierStyle = New clsStyle(Value)
        mp_oDefaultTierStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultTierStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE
        mp_oDefaultTierStyle.DrawTextInVisibleArea = True

        mp_oDefaultScrollBarStyle = New clsStyle(Value)
        mp_oDefaultScrollBarStyle.Appearance = E_STYLEAPPEARANCE.SA_FLAT
        mp_oDefaultScrollBarStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH
        mp_oDefaultScrollBarStyle.HatchStyle = GRE_HATCHSTYLE.HS_PERCENT50
        mp_oDefaultScrollBarStyle.HatchForeColor = Color.FromArgb(255, 192, 192, 192)
        mp_oDefaultScrollBarStyle.HatchBackColor = Colors.White
        mp_oDefaultScrollBarStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE
        mp_oDefaultScrollBarStyle.BorderColor = Color.FromArgb(255, 192, 192, 192)

        mp_oDefaultSBSeparator = New clsStyle(Value)
        mp_oDefaultSBSeparator.Appearance = E_STYLEAPPEARANCE.SA_RAISED

        mp_oDefaultSBNormalStyle = New clsStyle(Value)
        mp_oDefaultSBNormalStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED

        mp_oDefaultSBPressedStyle = New clsStyle(Value)
        mp_oDefaultSBPressedStyle.Appearance = E_STYLEAPPEARANCE.SA_SUNKEN

        mp_oDefaultSBDisabledStyle = New clsStyle(Value)
        mp_oDefaultSBDisabledStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED
        mp_oDefaultSBDisabledStyle.ScrollBarStyle.ArrowColor = Color.FromArgb(255, 192, 192, 192)

    End Sub

    Protected Overrides Sub Finalize()
        mp_oCollection = Nothing
        mp_oControl = Nothing
        MyBase.Finalize()
    End Sub

    
    Public ReadOnly Property Count() As Integer
        Get
            Return mp_oCollection.m_lCount
        End Get
    End Property

    
    Public Function Item(ByVal Index As String) As clsStyle
        Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.STYLES_ITEM_1, SYS_ERRORS.STYLES_ITEM_2, SYS_ERRORS.STYLES_ITEM_3, SYS_ERRORS.STYLES_ITEM_4), clsStyle)
    End Function

    Friend Function FItem(ByVal Index As String) As clsStyle
        If Index = "DS_TASK" Then
            Return mp_oDefaultTaskStyle
        ElseIf Index = "DS_ROW" Then
            Return mp_oDefaultRowStyle
        ElseIf Index = "DS_CLIENTAREA" Then
            Return mp_oDefaultClientAreaStyle
        ElseIf Index = "DS_CELL" Then
            Return mp_oDefaultCellStyle
        ElseIf Index = "DS_COLUMN" Then
            Return mp_oDefaultColumnStyle
        ElseIf Index = "DS_PERCENTAGE" Then
            Return mp_oDefaultPercentageStyle
        ElseIf Index = "DS_PREDECESSOR" Then
            Return mp_oDefaultPredecessorStyle
        ElseIf Index = "DS_TIMELINE" Then
            Return mp_oDefaultTimeLineStyle
        ElseIf Index = "DS_TIMEBLOCK" Then
            Return mp_oDefaultTimeBlockStyle
        ElseIf Index = "DS_TICKMARKAREA" Then
            Return mp_oDefaultTickMarkAreaStyle
        ElseIf Index = "DS_SPLITTER" Then
            Return mp_oDefaultSplitterStyle
        ElseIf Index = "DS_CONTROL" Then
            Return mp_oDefaultControlStyle
        ElseIf Index = "DS_NODE" Then
            Return mp_oDefaultNodeStyle
        ElseIf Index = "DS_TIER" Then
            Return mp_oDefaultTierStyle
        ElseIf Index = "DS_SCROLLBAR" Then
            Return mp_oDefaultScrollBarStyle
        ElseIf Index = "DS_SB_NORMAL" Then
            Return mp_oDefaultSBNormalStyle
        ElseIf Index = "DS_SB_PRESSED" Then
            Return mp_oDefaultSBPressedStyle
        ElseIf Index = "DS_SB_DISABLED" Then
            Return mp_oDefaultSBDisabledStyle
        ElseIf Index = "DS_SB_SEPARATOR" Then
            Return mp_oDefaultSBSeparator
        Else
            Return DirectCast(mp_oCollection.m_oItem(Index, SYS_ERRORS.STYLES_ITEM_1, SYS_ERRORS.STYLES_ITEM_2, SYS_ERRORS.STYLES_ITEM_3, SYS_ERRORS.STYLES_ITEM_4), clsStyle)
        End If
    End Function

    Friend ReadOnly Property oCollection() As clsCollectionBase
        Get
            Return mp_oCollection
        End Get
    End Property

    
    Public Function Add(ByVal Key As String) As clsStyle
        mp_oCollection.AddMode = True
        Dim oStyle As New clsStyle(mp_oControl)
        Key = mp_oControl.StrLib.StrTrim(Key)
        oStyle.Key = Key
        mp_oCollection.m_Add(oStyle, Key, SYS_ERRORS.STYLES_ADD_1, SYS_ERRORS.STYLES_ADD_2, True, SYS_ERRORS.STYLES_ADD_3)
        Return oStyle
    End Function

    
    Public Sub Clear()
        Dim lIndex As Integer
        Dim lIndex2 As Integer
        Dim oColumn As clsColumn
        Dim oRow As clsRow
        Dim oCell As clsCell
        Dim oTask As clsTask
        Dim oPredecessor As clsPredecessor
        Dim oTimeBlock As clsTimeBlock
        Dim oPercentage As clsPercentage
        Dim oView As clsView

        mp_oControl.StyleIndex = ""
        mp_oControl.Splitter.StyleIndex = ""

        mp_oControl.VerticalScrollBar.ScrollBar.StyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = ""
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = ""

        mp_oControl.HorizontalScrollBar.ScrollBar.StyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = ""
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = ""

        mp_oControl.ScrollBarSeparator.StyleIndex = ""

        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            oColumn.StyleIndex = ""
        Next
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.StyleIndex = ""
            oRow.ClientAreaStyleIndex = ""
            oRow.Node.StyleIndex = ""
            For lIndex2 = 1 To oRow.Cells.Count
                oCell = DirectCast(oRow.Cells.oCollection.m_oReturnArrayElement(lIndex2), clsCell)
                oCell.StyleIndex = ""
            Next
        Next lIndex
        For lIndex = 1 To mp_oControl.Tasks.Count
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            oTask.StyleIndex = ""
            oTask.WarningStyleIndex = ""
        Next
        For lIndex = 1 To mp_oControl.TimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TimeBlocks.oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            oTimeBlock.StyleIndex = ""
        Next
        For lIndex = 1 To mp_oControl.Percentages.Count
            oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            oPercentage.StyleIndex = ""
        Next
        For lIndex = 1 To mp_oControl.Predecessors.Count
            oPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
            oPredecessor.StyleIndex = ""
            oPredecessor.WarningStyleIndex = ""
            oPredecessor.SelectedStyleIndex = ""
        Next
        For lIndex = 1 To mp_oControl.Views.Count
            oView = DirectCast(mp_oControl.Views.oCollection.m_oReturnArrayElement(lIndex), clsView)
            oView.TimeLine.StyleIndex = ""
            oView.TimeLine.TickMarkArea.StyleIndex = ""
            oView.TimeLine.TierArea.UpperTier.StyleIndex = ""
            oView.TimeLine.TierArea.MiddleTier.StyleIndex = ""
            oView.TimeLine.TierArea.LowerTier.StyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = ""
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = ""
        Next
        oView = mp_oControl.Views.FItem("0")
        oView.TimeLine.StyleIndex = ""
        oView.TimeLine.TickMarkArea.StyleIndex = ""
        oView.TimeLine.TierArea.UpperTier.StyleIndex = ""
        oView.TimeLine.TierArea.MiddleTier.StyleIndex = ""
        oView.TimeLine.TierArea.LowerTier.StyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = ""
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = ""

        mp_oCollection.m_Clear()
    End Sub

    Public Sub Remove(ByVal Index As String)
        Dim sRIndex As String = ""
        Dim sRKey As String = ""
        mp_oCollection.m_GetKeyAndIndex(Index, sRKey, sRIndex)
        Dim lIndex As Integer
        Dim lIndex2 As Integer
        Dim oColumn As clsColumn
        Dim oRow As clsRow
        Dim oCell As clsCell
        Dim oTask As clsTask
        Dim oPredecessor As clsPredecessor
        Dim oTimeBlock As clsTimeBlock
        Dim oPercentage As clsPercentage
        Dim oView As clsView

        mp_oControl.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.StyleIndex)
        mp_oControl.Splitter.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.Splitter.StyleIndex)

        mp_oControl.VerticalScrollBar.ScrollBar.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.StyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex)
        mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.VerticalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex)

        mp_oControl.HorizontalScrollBar.ScrollBar.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.StyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.NormalStyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.PressedStyleIndex)
        mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.HorizontalScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex)

        mp_oControl.ScrollBarSeparator.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, mp_oControl.ScrollBarSeparator.StyleIndex)

        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            oColumn.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oColumn.StyleIndex)
        Next
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oRow.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oRow.StyleIndex)
            oRow.ClientAreaStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oRow.ClientAreaStyleIndex)
            oRow.Node.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oRow.Node.StyleIndex)
            For lIndex2 = 1 To oRow.Cells.Count
                oCell = DirectCast(oRow.Cells.oCollection.m_oReturnArrayElement(lIndex2), clsCell)
                oCell.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oCell.StyleIndex)
            Next
        Next lIndex
        For lIndex = 1 To mp_oControl.Tasks.Count
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            oTask.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oTask.StyleIndex)
            oTask.WarningStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oTask.WarningStyleIndex)
        Next
        For lIndex = 1 To mp_oControl.Predecessors.Count
            oPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor)
            oPredecessor.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor.StyleIndex)
            oPredecessor.WarningStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor.WarningStyleIndex)
            oPredecessor.SelectedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oPredecessor.SelectedStyleIndex)
        Next
        For lIndex = 1 To mp_oControl.TimeBlocks.Count
            oTimeBlock = DirectCast(mp_oControl.TimeBlocks.oCollection.m_oReturnArrayElement(lIndex), clsTimeBlock)
            oTimeBlock.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oTimeBlock.StyleIndex)
        Next
        For lIndex = 1 To mp_oControl.Percentages.Count
            oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            oPercentage.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oPercentage.StyleIndex)
        Next
        For lIndex = 1 To mp_oControl.Views.Count
            oView = DirectCast(mp_oControl.Views.oCollection.m_oReturnArrayElement(lIndex), clsView)
            oView.TimeLine.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.StyleIndex)
            oView.TimeLine.TickMarkArea.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TickMarkArea.StyleIndex)
            oView.TimeLine.TierArea.UpperTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.UpperTier.StyleIndex)
            oView.TimeLine.TierArea.MiddleTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.MiddleTier.StyleIndex)
            oView.TimeLine.TierArea.LowerTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.LowerTier.StyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex)
            oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex)
        Next
        oView = mp_oControl.Views.FItem("0")
        oView.TimeLine.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.StyleIndex)
        oView.TimeLine.TickMarkArea.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TickMarkArea.StyleIndex)
        oView.TimeLine.TierArea.UpperTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.UpperTier.StyleIndex)
        oView.TimeLine.TierArea.MiddleTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.MiddleTier.StyleIndex)
        oView.TimeLine.TierArea.LowerTier.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TierArea.LowerTier.StyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.StyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.NormalStyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.PressedStyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ArrowButtons.DisabledStyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.NormalStyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.PressedStyleIndex)
        oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex = mp_GetNewStyleIndex(sRKey, sRIndex, oView.TimeLine.TimeLineScrollBar.ScrollBar.ThumbButton.DisabledStyleIndex)

        mp_oCollection.m_Remove(Index, SYS_ERRORS.STYLES_REMOVE_1, SYS_ERRORS.STYLES_REMOVE_2, SYS_ERRORS.STYLES_REMOVE_3, SYS_ERRORS.STYLES_REMOVE_4)
    End Sub

    Private Function mp_GetNewStyleIndex(ByVal sKey As String, ByVal sIndex As String, ByVal sStyleIndex As String) As String
        If sIndex = sStyleIndex Then
            Return ""
        End If
        If sKey = sStyleIndex Then
            Return ""
        End If
        Return sStyleIndex
    End Function

    
    Public Function GetXML() As String
        Dim lIndex As Integer
        Dim oStyle As clsStyle
        Dim oXML As New clsXML(mp_oControl, "Styles")
        oXML.InitializeWriter()
        For lIndex = 1 To Count
            oStyle = DirectCast(mp_oCollection.m_oReturnArrayElement(lIndex), clsStyle)
            oXML.WriteObject(oStyle.GetXML)
        Next lIndex
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim lIndex As Integer
        Dim oXML As New clsXML(mp_oControl, "Styles")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        mp_oCollection.m_Clear()
        For lIndex = 1 To oXML.ReadCollectionCount
            Dim oStyle As New clsStyle(mp_oControl)
            oStyle.SetXML(oXML.ReadCollectionObject(lIndex))
            mp_oCollection.AddMode = True
            mp_oCollection.m_Add(oStyle, oStyle.Key, SYS_ERRORS.STYLES_ADD_1, SYS_ERRORS.STYLES_ADD_2, True, SYS_ERRORS.STYLES_ADD_3)
            oStyle = Nothing
        Next lIndex
    End Sub


End Class
