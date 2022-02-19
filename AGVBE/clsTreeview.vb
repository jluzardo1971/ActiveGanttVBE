Option Explicit On

Public Class clsTreeview

    Private Structure S_CHECKBOXCLICK
        Public lNodeIndex As Integer
        Public Sub Clear()
            lNodeIndex = 0
        End Sub
    End Structure

    Private Structure S_SIGNCLICK
        Public lNodeIndex As Integer
        Public Sub Clear()
            lNodeIndex = 0
        End Sub
    End Structure

    Private Structure S_ROWMOVEMENT
        Public lRowIndex As Integer
        Public lDestinationRowIndex As Integer
        Public Sub Clear()
            lRowIndex = 0
            lDestinationRowIndex = 0
        End Sub
    End Structure

    Private Structure S_ROWSIZING
        Public lRowIndex As Integer
        Public Sub Clear()
            lRowIndex = 0
        End Sub
    End Structure

    Private Structure S_ROWSELECTION
        Public lRowIndex As Integer
        Public lCellIndex As Integer
        Public Sub Clear()
            lRowIndex = 0
            lCellIndex = 0
        End Sub
    End Structure

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lLastVisibleNode As Integer
    Private mp_lIndentation As Integer
    Private mp_clrBackColor As Color
    Private mp_clrCheckBoxBorderColor As Color
    Private mp_clrCheckBoxColor As Color
    Private mp_clrCheckBoxMarkColor As Color
    Private mp_clrSelectedBackColor As Color
    Private mp_clrSelectedForeColor As Color
    Private mp_clrTreeLineColor As Color
    Private mp_clrPlusMinusBorderColor As Color
    Private mp_clrPlusMinusSignColor As Color
    Private mp_bCheckBoxes As Boolean
    Private mp_bTreeLines As Boolean
    Private mp_bImages As Boolean
    Private mp_bPlusMinusSigns As Boolean
    Private mp_bFullColumnSelect As Boolean
    Private mp_bExpansionOnSelection As Boolean
    Private mp_sPathSeparator As String
    Private mp_yOperation As E_OPERATION
    Private s_chkCLK As S_CHECKBOXCLICK
    Private s_sgnCLK As S_SIGNCLICK
    Private s_rowMVT As S_ROWMOVEMENT
    Private s_rowSZ As S_ROWSIZING
    Private s_rowSEL As S_ROWSELECTION

    Public Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_lLastVisibleNode = 0
        mp_lIndentation = 20
        mp_clrBackColor = Colors.White
        mp_clrCheckBoxBorderColor = Colors.Gray
        mp_clrCheckBoxColor = Colors.White
        mp_clrCheckBoxMarkColor = Colors.Black
        mp_clrSelectedBackColor = Colors.Blue
        mp_clrSelectedForeColor = Colors.White
        mp_clrTreeLineColor = Colors.Gray
        mp_clrPlusMinusBorderColor = Colors.Gray
        mp_clrPlusMinusSignColor = Colors.Black
        mp_bCheckBoxes = False
        mp_bTreeLines = True
        mp_bImages = True
        mp_bPlusMinusSigns = True
        mp_bFullColumnSelect = False
        mp_bExpansionOnSelection = False
        mp_sPathSeparator = "/"
        mp_yOperation = E_OPERATION.EO_NONE
    End Sub

    Friend Function OverControl(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        Dim lIndex As Integer
        If mp_oControl.TreeviewColumnIndex = 0 Then
            Return False
        End If
        If Not (X >= LeftTrim And X <= RightTrim) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Visible = True Then
                If Y >= oRow.Top And Y <= oRow.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function CursorPosition(ByVal X As Integer, ByVal Y As Integer) As E_EVENTTARGET
        If mp_bOverCheckBox(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TREEVIEWCHECKBOX
        ElseIf mp_bOverPlusMinusSign(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TREEVIEWSIGN
        ElseIf mp_oControl.MouseKeyboardEvents.mp_bOverSelectedRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDROW
        ElseIf mp_oControl.MouseKeyboardEvents.mp_bOverRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_ROW
        End If
        Return E_EVENTTARGET.EVT_NONE
    End Function

    Friend Sub OnMouseHover(ByVal X As Integer, ByVal Y As Integer)
        Select Case CursorPosition(X, Y)
            Case E_EVENTTARGET.EVT_TREEVIEWCHECKBOX
                mp_EO_CHECKBOXCLICK(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_TREEVIEWSIGN
                mp_EO_SIGNCLICK(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDROW
                If mp_bCursorEditTextNode(X, Y) = True Then
                    Return
                End If
                Select Case mp_oControl.MouseKeyboardEvents.mp_yRowArea(X, Y)
                    Case E_AREA.EA_BOTTOM
                        mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                    Case E_AREA.EA_CENTER
                        mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                End Select
            Case E_EVENTTARGET.EVT_ROW
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
        End Select
        System.Diagnostics.Debug.Assert(mp_yOperation = E_OPERATION.EO_NONE)
    End Sub

    Friend Sub OnMouseDown(ByVal X As Integer, ByVal Y As Integer)
        System.Diagnostics.Debug.Assert(mp_yOperation = E_OPERATION.EO_NONE)
        Select Case CursorPosition(X, Y)
            Case E_EVENTTARGET.EVT_TREEVIEWCHECKBOX
                mp_EO_CHECKBOXCLICK(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_CHECKBOXCLICK
            Case E_EVENTTARGET.EVT_TREEVIEWSIGN
                mp_EO_SIGNCLICK(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_SIGNCLICK
            Case E_EVENTTARGET.EVT_SELECTEDROW
                If mp_bShowEditTextNode(X, Y) = True Then
                    Return
                End If
                Select Case mp_oControl.MouseKeyboardEvents.mp_yRowArea(X, Y)
                    Case E_AREA.EA_BOTTOM
                        mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                        mp_yOperation = E_OPERATION.EO_ROWSIZING
                    Case E_AREA.EA_CENTER
                        mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                        mp_yOperation = E_OPERATION.EO_ROWMOVEMENT
                End Select
            Case E_EVENTTARGET.EVT_ROW
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_ROWSELECTION
        End Select
    End Sub

    Friend Sub OnMouseMove(ByVal X As Integer, ByVal Y As Integer)
        Dim yOperation As E_OPERATION = mp_yOperation
        Select Case mp_yOperation
            Case E_OPERATION.EO_CHECKBOXCLICK
                mp_EO_CHECKBOXCLICK(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_SIGNCLICK
                mp_EO_SIGNCLICK(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWMOVEMENT
                mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWSIZING
                mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWSELECTION
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
        End Select
        System.Diagnostics.Debug.Assert(yOperation = mp_yOperation)
    End Sub

    Friend Sub OnMouseUp(ByVal X As Integer, ByVal Y As Integer)
        Select Case mp_yOperation
            Case E_OPERATION.EO_CHECKBOXCLICK
                mp_EO_CHECKBOXCLICK(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_SIGNCLICK
                mp_EO_SIGNCLICK(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWMOVEMENT
                mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWSIZING
                mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWSELECTION
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
        End Select
        mp_yOperation = E_OPERATION.EO_NONE
    End Sub

    Private Sub mp_EO_CHECKBOXCLICK(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        Dim oNode As clsNode = Nothing
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_chkCLK.lNodeIndex = mp_oControl.MathLib.GetNodeIndexByCheckBoxPosition(X, Y)
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_chkCLK.lNodeIndex), clsRow)
                oNode = oRow.Node
                oNode.Checked = Not oNode.Checked
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.NodeEventArgs.Clear()
                mp_oControl.NodeEventArgs.Index = s_chkCLK.lNodeIndex
                mp_oControl.FireNodeChecked()
                mp_oControl.Redraw()
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_SIGNCLICK(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        Dim oNode As clsNode = Nothing
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_sgnCLK.lNodeIndex = mp_oControl.MathLib.GetNodeIndexBySignPosition(X, Y)
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_sgnCLK.lNodeIndex), clsRow)
                oNode = oRow.Node
                oNode.Expanded = Not oNode.Expanded
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.NodeEventArgs.Clear()
                mp_oControl.NodeEventArgs.Index = s_sgnCLK.lNodeIndex
                mp_oControl.FireNodeExpanded()
                mp_oControl.Redraw()
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_ROWMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oDestinationRow As clsRow = Nothing
        If mp_oControl.AllowRowMove = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_MOVEROW)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_rowMVT.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                System.Diagnostics.Debug.Assert(s_rowMVT.lRowIndex >= 1)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                mp_oControl.ObjectStateChangedEventArgs.Index = s_rowMVT.lRowIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectMove()
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.MouseKeyboardEvents.mp_DynamicRowMove(Y)
                    s_rowMVT.lDestinationRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                    If s_rowMVT.lDestinationRowIndex >= 1 Then
                        mp_oControl.clsG.EraseReversibleFrames()
                        mp_oControl.ObjectStateChangedEventArgs.DestinationIndex = s_rowMVT.lDestinationRowIndex
                        mp_oControl.FireObjectMove()
                        If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                            oDestinationRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex), clsRow)
                            mp_oControl.MouseKeyboardEvents.mp_DrawMovingReversibleFrame(oDestinationRow.Left, oDestinationRow.Top, oDestinationRow.Right, oDestinationRow.Bottom, E_FOCUSTYPE.FCT_NORMAL)
                        End If
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireEndObjectMove()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        s_rowMVT.lDestinationRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                        If s_rowMVT.lDestinationRowIndex >= 1 And (s_rowMVT.lRowIndex <> s_rowMVT.lDestinationRowIndex) Then
                            Dim oRow As clsRow = Nothing
                            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_rowMVT.lRowIndex), clsRow)
                            oDestinationRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex), clsRow)
                            oRow.Node.Depth = oDestinationRow.Node.Depth
                            mp_oControl.SelectedRowIndex = mp_oControl.Rows.oCollection.m_lCopyAndMoveItems(s_rowMVT.lRowIndex, s_rowMVT.lDestinationRowIndex)
                            mp_oControl.FireCompleteObjectMove()
                        End If
                    End If
                End If
                mp_oControl.Redraw()
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_ROWSIZING(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        If mp_oControl.AllowRowSize = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_ROWHEIGHT)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_rowSZ.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                System.Diagnostics.Debug.Assert(s_rowSZ.lRowIndex >= 1)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                mp_oControl.ObjectStateChangedEventArgs.Index = s_rowSZ.lRowIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectSize()
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireBeginObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        mp_oControl.MouseKeyboardEvents.mp_DrawMovingReversibleFrame(0, Y, mp_oControl.clsG.Width(), Y + 2, E_FOCUSTYPE.FCT_NORMAL)
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireEndObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_rowSZ.lRowIndex), clsRow)
                        oRow.Height = oRow.Height + (Y - oRow.Bottom)
                        If oRow.Height < mp_oControl.MinRowHeight Then
                            oRow.Height = mp_oControl.MinRowHeight
                        End If
                        mp_oControl.FireCompleteObjectSize()
                    End If
                End If
                mp_oControl.Redraw()
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_ROWSELECTION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_rowSEL.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                s_rowSEL.lCellIndex = mp_oControl.MathLib.GetCellIndexByPosition(X)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedRowIndex = s_rowSEL.lRowIndex
                mp_oControl.SelectedCellIndex = s_rowSEL.lCellIndex
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
                mp_oControl.ObjectSelectedEventArgs.Clear()
                mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedRowIndex
                mp_oControl.FireObjectSelected()
                mp_oControl.Redraw()
                mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Function mp_bOverCheckBox(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        Dim bReturn As Boolean
        If mp_bCheckBoxes = False Then
            Return False
        End If
        bReturn = False
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And X >= (oNode.CheckBoxLeft) And X <= (oNode.CheckBoxLeft + 13) And Y <= (oNode.YCenter + 6) And Y >= (oNode.YCenter - 7) Then
                bReturn = True
            End If
        Next lIndex
        Return bReturn
    End Function

    Private Function mp_bOverPlusMinusSign(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim lIndex As Integer
        Dim oNode As clsNode = Nothing
        Dim oRow As clsRow = Nothing
        Dim bReturn As Boolean
        If mp_bPlusMinusSigns = False Then
            Return False
        End If
        bReturn = False
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            oNode = oRow.Node
            If oRow.ClientAreaVisibility = E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA And X >= (oNode.Left - 5) And X <= (oNode.Left + 5) And Y <= (oNode.YCenter + 5) And Y >= (oNode.YCenter - 5) Then
                bReturn = True
            End If
        Next lIndex
        Return bReturn
    End Function

    Friend Sub Draw()
        If mp_oControl.TreeviewColumnIndex = 0 Then
            Return
        End If
        If mp_oControl.Columns.Item(mp_oControl.TreeviewColumnIndex.ToString()).Visible = False Then
            Return
        End If
        mp_oControl.clsG.ClipRegion(LeftTrim, mp_oControl.CurrentViewObject.ClientArea.Top, RightTrim, mp_oControl.clsG.Height() - mt_BorderThickness - 1, False)
        mp_oControl.Rows.NodesDrawBackground()
        mp_oControl.clsG.ClipRegion(LeftTrim, mp_oControl.CurrentViewObject.ClientArea.Top, RightTrim - 2, mp_oControl.clsG.Height() - mt_BorderThickness - 1, False)
        mp_oControl.Rows.NodesDraw()
        mp_oControl.Rows.NodesDrawTreeLines()
        mp_oControl.Rows.NodesDrawElements()
        mp_oControl.clsG.ClearClipRegion()
    End Sub

    Friend ReadOnly Property f_FirstVisibleNode() As Integer
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return 0
            Else
                Return mp_oControl.VerticalScrollBar.Value
            End If
        End Get
    End Property

    Public Property FirstVisibleNode() As Integer
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return 0
            Else
                Return mp_oControl.Rows.RealFirstVisibleRow
            End If
        End Get
        Set(ByVal Value As Integer)
            If Value < 1 Then
                Value = 1
            ElseIf ((Value > mp_oControl.Rows.Count) And (mp_oControl.Rows.Count <> 0)) Then
                Value = mp_oControl.Rows.Count
            End If
            mp_oControl.VerticalScrollBar.Value = Value
        End Set
    End Property


    Public ReadOnly Property LastVisibleNode() As Integer
        Get
            Return mp_lLastVisibleNode
        End Get
    End Property

    Friend WriteOnly Property f_LastVisibleNode() As Integer
        Set(ByVal Value As Integer)
            mp_lLastVisibleNode = Value
        End Set
    End Property

    Friend ReadOnly Property mt_BorderThickness() As Integer
        Get
            Return mp_oControl.mt_BorderThickness
        End Get
    End Property


    Public Property Indentation() As Integer
        Get
            Return mp_lIndentation
        End Get
        Set(ByVal Value As Integer)
            mp_lIndentation = Value
        End Set
    End Property


    Public Sub ClearSelections()
        mp_oControl.SelectedRowIndex = 0
    End Sub


    Public Property CheckBoxBorderColor() As Color
        Get
            Return mp_clrCheckBoxBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrCheckBoxBorderColor = Value
        End Set
    End Property


    Public Property CheckBoxColor() As Color
        Get
            Return mp_clrCheckBoxColor
        End Get
        Set(ByVal Value As Color)
            mp_clrCheckBoxColor = Value
        End Set
    End Property


    Public Property CheckBoxMarkColor() As Color
        Get
            Return mp_clrCheckBoxMarkColor
        End Get
        Set(ByVal Value As Color)
            mp_clrCheckBoxMarkColor = Value
        End Set
    End Property


    Public Property BackColor() As Color
        Get
            Return mp_clrBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrBackColor = Value
        End Set
    End Property


    Public Property PathSeparator() As String
        Get
            Return mp_sPathSeparator
        End Get
        Set(ByVal Value As String)
            mp_sPathSeparator = Value
        End Set
    End Property


    Public Property TreeLines() As Boolean
        Get
            Return mp_bTreeLines
        End Get
        Set(ByVal Value As Boolean)
            mp_bTreeLines = Value
        End Set
    End Property


    Public Property PlusMinusSigns() As Boolean
        Get
            Return mp_bPlusMinusSigns
        End Get
        Set(ByVal Value As Boolean)
            mp_bPlusMinusSigns = Value
        End Set
    End Property


    Public Property Images() As Boolean
        Get
            Return mp_bImages
        End Get
        Set(ByVal Value As Boolean)
            mp_bImages = Value
        End Set
    End Property


    Public Property CheckBoxes() As Boolean
        Get
            Return mp_bCheckBoxes
        End Get
        Set(ByVal Value As Boolean)
            mp_bCheckBoxes = Value
        End Set
    End Property


    Public Property FullColumnSelect() As Boolean
        Get
            Return mp_bFullColumnSelect
        End Get
        Set(ByVal Value As Boolean)
            mp_bFullColumnSelect = Value
        End Set
    End Property


    Public Property ExpansionOnSelection() As Boolean
        Get
            Return mp_bExpansionOnSelection
        End Get
        Set(ByVal Value As Boolean)
            mp_bExpansionOnSelection = Value
        End Set
    End Property

    Public Property SelectedBackColor() As Color
        Get
            Return mp_clrSelectedBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrSelectedBackColor = Value
        End Set
    End Property


    Public Property SelectedForeColor() As Color
        Get
            Return mp_clrSelectedForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrSelectedForeColor = Value
        End Set
    End Property


    Public Property TreeLineColor() As Color
        Get
            Return mp_clrTreeLineColor
        End Get
        Set(ByVal Value As Color)
            mp_clrTreeLineColor = Value
        End Set
    End Property


    Public Property PlusMinusBorderColor() As Color
        Get
            Return mp_clrPlusMinusBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrPlusMinusBorderColor = Value
        End Set
    End Property


    Public Property PlusMinusSignColor() As Color
        Get
            Return mp_clrPlusMinusSignColor
        End Get
        Set(ByVal Value As Color)
            mp_clrPlusMinusSignColor = Value
        End Set
    End Property

    Friend ReadOnly Property Left() As Integer
        Get
            If mp_oControl.TreeviewColumnIndex = 0 Then
                Return 0
            End If
            Return mp_oControl.Columns.Item(mp_oControl.TreeviewColumnIndex.ToString()).Left
        End Get
    End Property

    Friend ReadOnly Property Right() As Integer
        Get
            If mp_oControl.TreeviewColumnIndex = 0 Then
                Return 0
            End If
            Return mp_oControl.Columns.Item(mp_oControl.TreeviewColumnIndex.ToString()).Right
        End Get
    End Property

    Friend ReadOnly Property LeftTrim() As Integer
        Get
            If mp_oControl.TreeviewColumnIndex = 0 Then
                Return 0
            End If
            Return mp_oControl.Columns.Item(mp_oControl.TreeviewColumnIndex.ToString()).LeftTrim
        End Get
    End Property

    Friend ReadOnly Property RightTrim() As Integer
        Get
            If mp_oControl.TreeviewColumnIndex = 0 Then
                Return 0
            End If
            Return mp_oControl.Columns.Item(mp_oControl.TreeviewColumnIndex.ToString()).RightTrim
        End Get
    End Property

    Friend Function mp_bCursorEditTextNode(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oNode As clsNode
        Dim oRow As clsRow
        oRow = mp_oControl.Rows.Item(mp_oControl.SelectedRowIndex.ToString())
        oNode = oRow.Node
        If oNode.AllowTextEdit = True Then
            If X >= oNode.mp_lTextLeft And X <= oNode.mp_lTextRight Then
                If Y >= oNode.mp_lTextTop And Y <= oNode.mp_lTextBottom Then
                    mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_IBEAM)
                    Return True
                End If
            End If
        End If
        mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        Return False
    End Function

    Friend Function mp_bShowEditTextNode(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oNode As clsNode
        Dim oRow As clsRow
        oRow = mp_oControl.Rows.Item(mp_oControl.SelectedRowIndex.ToString())
        oNode = oRow.Node
        If oNode.AllowTextEdit = True Then
            If X >= oNode.mp_lTextLeft And X <= oNode.mp_lTextRight Then
                If Y >= oNode.mp_lTextTop And Y <= oNode.mp_lTextBottom Then
                    mp_oControl.mp_oTextBox.Initialize(mp_oControl.SelectedRowIndex, 0, E_TEXTOBJECTTYPE.TOT_NODE, X, Y)
                    Return True
                End If
            End If
        End If
        Return False
    End Function


    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Treeview")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("BackColor", mp_clrBackColor)
        oXML.ReadProperty("CheckBoxBorderColor", mp_clrCheckBoxBorderColor)
        oXML.ReadProperty("CheckBoxColor", mp_clrCheckBoxColor)
        oXML.ReadProperty("CheckBoxes", mp_bCheckBoxes)
        oXML.ReadProperty("CheckBoxMarkColor", mp_clrCheckBoxMarkColor)
        oXML.ReadProperty("ExpansionOnSelection", mp_bExpansionOnSelection)
        oXML.ReadProperty("FullColumnSelect", mp_bFullColumnSelect)
        oXML.ReadProperty("Images", mp_bImages)
        oXML.ReadProperty("Indentation", mp_lIndentation)
        oXML.ReadProperty("PathSeparator", mp_sPathSeparator)
        oXML.ReadProperty("PlusMinusBorderColor", mp_clrPlusMinusBorderColor)
        oXML.ReadProperty("PlusMinusSignColor", mp_clrPlusMinusSignColor)
        oXML.ReadProperty("PlusMinusSigns", mp_bPlusMinusSigns)
        oXML.ReadProperty("SelectedBackColor", mp_clrSelectedBackColor)
        oXML.ReadProperty("SelectedForeColor", mp_clrSelectedForeColor)
        oXML.ReadProperty("TreeLineColor", mp_clrTreeLineColor)
        oXML.ReadProperty("TreeLines", mp_bTreeLines)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Treeview")
        oXML.InitializeWriter()
        oXML.WriteProperty("BackColor", mp_clrBackColor)
        oXML.WriteProperty("CheckBoxBorderColor", mp_clrCheckBoxBorderColor)
        oXML.WriteProperty("CheckBoxColor", mp_clrCheckBoxColor)
        oXML.WriteProperty("CheckBoxes", mp_bCheckBoxes)
        oXML.WriteProperty("CheckBoxMarkColor", mp_clrCheckBoxMarkColor)
        oXML.WriteProperty("ExpansionOnSelection", mp_bExpansionOnSelection)
        oXML.WriteProperty("FullColumnSelect", mp_bFullColumnSelect)
        oXML.WriteProperty("Images", mp_bImages)
        oXML.WriteProperty("Indentation", mp_lIndentation)
        oXML.WriteProperty("PathSeparator", mp_sPathSeparator)
        oXML.WriteProperty("PlusMinusBorderColor", mp_clrPlusMinusBorderColor)
        oXML.WriteProperty("PlusMinusSignColor", mp_clrPlusMinusSignColor)
        oXML.WriteProperty("PlusMinusSigns", mp_bPlusMinusSigns)
        oXML.WriteProperty("SelectedBackColor", mp_clrSelectedBackColor)
        oXML.WriteProperty("SelectedForeColor", mp_clrSelectedForeColor)
        oXML.WriteProperty("TreeLineColor", mp_clrTreeLineColor)
        oXML.WriteProperty("TreeLines", mp_bTreeLines)
        Return oXML.GetXML()
    End Function

End Class
