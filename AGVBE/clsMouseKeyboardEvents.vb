Option Explicit On 

Friend Class clsMouseKeyboardEvents

    Private Enum E_OPERATIONMODIFIER
        OPM_NONE = 0
        OPM_PREDECESSORMODE = 1
    End Enum

    Private Structure S_TIMELINEMOVEMENT
        Public lX As Integer
        Public Sub Clear()
            lX = 0
        End Sub
    End Structure

    Private Structure S_COLUMNMOVEMENT
        Public lColumnIndex As Integer
        Public lDestinationColumnIndex As Integer
        Public Sub Clear()
            lColumnIndex = 0
            lDestinationColumnIndex = 0
        End Sub
    End Structure

    Private Structure S_COLUMNSIZING
        Public lColumnIndex As Integer
        Public Sub Clear()
            lColumnIndex = 0
        End Sub
    End Structure

    Private Structure S_COLUMNSELECTION
        Public lColumnIndex As Integer
        Public Sub Clear()
            lColumnIndex = 0
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

    Private Structure S_TASKMOVEMENT
        Public lInitialRowIndex As Integer
        Public sInitialRowKey As String
        Public lFinalRowIndex As Integer
        Public dtInitialStartDate As AGVBE.DateTime
        Public dtInitialEndDate As AGVBE.DateTime
        Public lDeltaLeft As Integer
        Public lDeltaRight As Integer
        Public lDurationFactor As Integer
        Public Sub Clear()
            lInitialRowIndex = 0
            sInitialRowKey = ""
            lFinalRowIndex = 0
            dtInitialStartDate = New AGVBE.DateTime()
            dtInitialEndDate = New AGVBE.DateTime()
            lDeltaLeft = 0
            lDeltaRight = 0
            lDurationFactor = 0
        End Sub
    End Structure

    Private Structure S_TASKSTRETCHLEFT
        Public dtInitialStartDate As AGVBE.DateTime
        Public dtInitialEndDate As AGVBE.DateTime
        Public dtFinalStartDate As AGVBE.DateTime
        Public lRowIndex As Integer
        Public Sub Clear()
            dtInitialStartDate = New AGVBE.DateTime()
            dtInitialEndDate = New AGVBE.DateTime()
            dtFinalStartDate = New AGVBE.DateTime()
            lRowIndex = 0
        End Sub
    End Structure

    Private Structure S_TASKSTRETCHRIGHT
        Public dtInitialStartDate As AGVBE.DateTime
        Public dtInitialEndDate As AGVBE.DateTime
        Public dtFinalEndDate As AGVBE.DateTime
        Public lRowIndex As Integer
        Public Sub Clear()
            dtInitialStartDate = New AGVBE.DateTime()
            dtInitialEndDate = New AGVBE.DateTime()
            dtFinalEndDate = New AGVBE.DateTime()
            lRowIndex = 0
        End Sub
    End Structure

    Private Structure S_TASKSELECTION
        Public lTaskIndex As Integer
        Public Sub Clear()
            lTaskIndex = 0
        End Sub
    End Structure

    Private Structure S_TASKADDITION
        Public bCancel As Boolean
        Public bInConflict As Boolean
        Public dtStartDate As AGVBE.DateTime
        Public dtEndDate As AGVBE.DateTime
        Public lRowIndex As Integer
        Public Sub Clear()
            bCancel = False
            bInConflict = False
            dtStartDate = New AGVBE.DateTime()
            dtEndDate = New AGVBE.DateTime()
            lRowIndex = 0
        End Sub
    End Structure

    Private Structure S_PERCENTAGESELECTION
        Public lPercentageIndex As Integer
        Public Sub Clear()
            lPercentageIndex = 0
        End Sub
    End Structure

    Private Structure S_PERCENTAGESIZING
        Public lX As Integer
        Public lTaskStart As Integer
        Public lTaskEnd As Integer
        Public lTaskIndex As Integer
        Public bMouseMove As Boolean
        Public Sub Clear()
            lX = 0
            lTaskStart = 0
            lTaskEnd = 0
            lTaskIndex = 0
            bMouseMove = True
        End Sub
    End Structure

    Private Structure S_PREDECESSORSELECTION
        Public lPredecessorIndex As Integer
        Public Sub Clear()
            lPredecessorIndex = 0
        End Sub
    End Structure

    Private Structure S_PREDECESSORADDITION
        Public lXStart As Integer
        Public lYStart As Integer
        Public lTaskIndex As Integer
        Public sTaskPosition As String
        Public lPredecessorIndex As Integer
        Public sPredecessorKey As String
        Public sPredecessorPosition As String
        Public bCancel As Boolean
        Public bCantAccept As Boolean
        Public Sub Clear()
            lXStart = 0
            lYStart = 0
            lTaskIndex = 0
            sTaskPosition = ""
            lPredecessorIndex = 0
            sPredecessorKey = ""
            sPredecessorPosition = ""
            bCancel = False
            bCantAccept = False
        End Sub
    End Structure

    Friend mp_yOperation As E_OPERATION = E_OPERATION.EO_NONE
    Private mp_yOperationModifier As E_OPERATIONMODIFIER = E_OPERATIONMODIFIER.OPM_NONE
    Private mp_oControl As ActiveGanttVBECtl
    Private s_tmlnMVT As S_TIMELINEMOVEMENT
    Private s_clmnMVT As S_COLUMNMOVEMENT
    Private s_clmnSZ As S_COLUMNSIZING
    Private s_clmnSEL As S_COLUMNSELECTION
    Private s_rowMVT As S_ROWMOVEMENT
    Private s_rowSZ As S_ROWSIZING
    Private s_rowSEL As S_ROWSELECTION
    Private s_tskMVT As S_TASKMOVEMENT
    Private s_tskSTL As S_TASKSTRETCHLEFT
    Private s_tskSTR As S_TASKSTRETCHRIGHT
    Private s_tskSEL As S_TASKSELECTION
    Private s_tskADD As S_TASKADDITION
    Private s_perSEL As S_PERCENTAGESELECTION
    Private s_perSZ As S_PERCENTAGESIZING
    Private s_preADD As S_PREDECESSORADDITION
    Private s_preSEL As S_PREDECESSORSELECTION

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
    End Sub

    Friend Sub KeyDown(ByVal KeyCode As System.Windows.Input.Key)
        mp_oControl.KeyEventArgs.Clear()
        mp_oControl.KeyEventArgs.KeyCode = KeyCode
        mp_oControl.KeyEventArgs.Cancel = False
        mp_oControl.FireControlKeyDown()
        If mp_oControl.KeyEventArgs.Cancel = True Then
            Return
        End If
        If KeyCode = mp_oControl.PredecessorAddModeKey Then
            mp_yOperationModifier = E_OPERATIONMODIFIER.OPM_PREDECESSORMODE
        End If
    End Sub

    Friend Sub KeyUp(ByVal KeyCode As System.Windows.Input.Key)
        mp_oControl.KeyEventArgs.Clear()
        mp_oControl.KeyEventArgs.KeyCode = KeyCode
        mp_oControl.KeyEventArgs.Cancel = False
        mp_oControl.FireControlKeyUp()
        mp_yOperationModifier = E_OPERATIONMODIFIER.OPM_NONE
    End Sub

    Friend Sub KeyPress(ByVal Key As Char)
        mp_oControl.KeyEventArgs.Clear()
        mp_oControl.KeyEventArgs.CharacterCode = Key
        mp_oControl.KeyEventArgs.Cancel = False
        mp_oControl.FireControlKeyDown()
        If mp_oControl.KeyEventArgs.Cancel = True Then
            Return
        End If
    End Sub

    Friend Sub OnMouseClick()
        mp_oControl.FireControlClick()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            Return
        End If
    End Sub

    Friend Sub OnMouseDblClick()
        mp_oControl.FireControlDblClick()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            Return
        End If
    End Sub

    Friend Sub OnMouseMoveGeneral(ByVal X As Integer, ByVal Y As Integer)
        If mp_yOperation = E_OPERATION.EO_NONE Then
            OnMouseHover(X, Y)
        Else
            OnMouseMove(X, Y)
        End If
    End Sub

    Friend Sub OnMouseHover(ByVal X As Integer, ByVal Y As Integer)
        Dim yEventTarget As E_EVENTTARGET = E_EVENTTARGET.EVT_NONE
        yEventTarget = CursorPosition(X, Y)
        mp_oControl.MouseEventArgs.Clear()
        mp_oControl.MouseEventArgs.X = X
        mp_oControl.MouseEventArgs.Y = Y
        mp_oControl.MouseEventArgs.EventTarget = yEventTarget
        mp_oControl.MouseEventArgs.Cancel = False
        mp_oControl.FireControlMouseHover()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            Return
        End If
        mp_oControl.ToolTipEventArgs.Clear()
        mp_oControl.ToolTipEventArgs.X = X
        mp_oControl.ToolTipEventArgs.Y = Y
        mp_oControl.FireToolTipOnMouseHover(yEventTarget)
        Select Case yEventTarget
            Case E_EVENTTARGET.EVT_VSCROLLBAR
                mp_oControl.VerticalScrollBar.ScrollBar.OnMouseHover(X, Y)
            Case E_EVENTTARGET.EVT_HSCROLLBAR
                mp_oControl.HorizontalScrollBar.ScrollBar.OnMouseHover(X, Y)
            Case E_EVENTTARGET.EVT_TIMELINESCROLLBAR
                mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OnMouseHover(X, Y)
            Case E_EVENTTARGET.EVT_TREEVIEW
                mp_oControl.Treeview.OnMouseHover(X, Y)
            Case E_EVENTTARGET.EVT_SPLITTER
                mp_EO_SPLITTERMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_TIMELINE
                mp_EO_TIMELINEMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDCOLUMN
                If mp_bCursorEditTextColumn(X, Y) = True Then
                    Return
                End If
                Select Case mp_yColumnArea(X, Y)
                    Case E_AREA.EA_RIGHT
                        mp_EO_COLUMNSIZING(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                    Case E_AREA.EA_CENTER
                        mp_EO_COLUMNMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                End Select
            Case E_EVENTTARGET.EVT_COLUMN
                mp_EO_COLUMNSELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDROW
                If mp_bCursorEditTextRow(X, Y) = True Then
                    Return
                End If
                Select Case mp_yRowArea(X, Y)
                    Case E_AREA.EA_BOTTOM
                        mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                    Case E_AREA.EA_CENTER
                        mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                End Select
            Case E_EVENTTARGET.EVT_ROW
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDPERCENTAGE
                mp_EO_PERCENTAGESIZING(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_PERCENTAGE
                mp_EO_PERCENTAGESELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDTASK
                If mp_yOperationModifier = E_OPERATIONMODIFIER.OPM_PREDECESSORMODE Then
                    mp_EO_PREDECESSORADDITION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                Else
                    If mp_bCursorEditTextTask(X, Y) = True Then
                        Return
                    End If
                    Select Case mp_yTaskArea(X, Y)
                        Case E_AREA.EA_CENTER
                            mp_EO_TASKMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                        Case E_AREA.EA_LEFT
                            mp_EO_TASKSTRETCHLEFT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                        Case E_AREA.EA_RIGHT
                            mp_EO_TASKSTRETCHRIGHT(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
                        Case E_AREA.EA_NONE
                    End Select
                End If
            Case E_EVENTTARGET.EVT_TASK
                mp_EO_TASKSELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_SELECTEDPREDECESSOR
                '//
            Case E_EVENTTARGET.EVT_PREDECESSOR
                mp_EO_PREDECESSORSELECTION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_CLIENTAREA
                If mp_bCursorEditTextTask(X, Y) = True Then
                    Return
                End If
                mp_EO_TASKADDITION(E_MOUSEKEYBOARDEVENTS.MouseHover, X, Y)
            Case E_EVENTTARGET.EVT_NONE
        mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        'Moving over empty space
        End Select
        System.Diagnostics.Debug.Assert(mp_yOperation = E_OPERATION.EO_NONE)
    End Sub

    Friend Sub OnMouseDown(ByVal X As Integer, ByVal Y As Integer, ByVal Button As E_MOUSEBUTTONS)
        mp_oControl.mp_oTextBox.Terminate()
        If mp_yOperation <> E_OPERATION.EO_NONE Then
            mp_yOperation = E_OPERATION.EO_NONE
        End If
        System.Diagnostics.Debug.Assert(mp_yOperation = E_OPERATION.EO_NONE)
        Dim EventTarget As E_EVENTTARGET = E_EVENTTARGET.EVT_NONE
        EventTarget = CursorPosition(X, Y)
        mp_oControl.MouseEventArgs.Clear()
        mp_oControl.MouseEventArgs.X = X
        mp_oControl.MouseEventArgs.Y = Y
        mp_oControl.MouseEventArgs.EventTarget = EventTarget
        mp_oControl.MouseEventArgs.Button = Button
        mp_oControl.MouseEventArgs.Cancel = False
        mp_oControl.FireControlMouseDown()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            Return
        End If
        Select Case CursorPosition(X, Y)
            Case E_EVENTTARGET.EVT_VSCROLLBAR
                mp_oControl.VerticalScrollBar.ScrollBar.OnMouseDown(X, Y)
                mp_yOperation = E_OPERATION.EO_VERTICALSCROLLBAR
            Case E_EVENTTARGET.EVT_HSCROLLBAR
                mp_oControl.HorizontalScrollBar.ScrollBar.OnMouseDown(X, Y)
                mp_yOperation = E_OPERATION.EO_HORIZONTALSCROLLBAR
            Case E_EVENTTARGET.EVT_TIMELINESCROLLBAR
                mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OnMouseDown(X, Y)
                mp_yOperation = E_OPERATION.EO_TIMELINESCROLLBAR
            Case E_EVENTTARGET.EVT_TREEVIEW
                mp_oControl.Treeview.OnMouseDown(X, Y)
                mp_yOperation = E_OPERATION.EO_TREEVIEW
            Case E_EVENTTARGET.EVT_SPLITTER
                mp_EO_SPLITTERMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_SPLITTERMOVEMENT
            Case E_EVENTTARGET.EVT_TIMELINE
                mp_EO_TIMELINEMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_TIMELINEMOVEMENT
            Case E_EVENTTARGET.EVT_SELECTEDCOLUMN
                If mp_bShowEditTextColumn(X, Y) = True Then
                    Return
                End If
                Select Case mp_yColumnArea(X, Y)
                    Case E_AREA.EA_RIGHT
                        mp_EO_COLUMNSIZING(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                        mp_yOperation = E_OPERATION.EO_COLUMNSIZING
                    Case E_AREA.EA_CENTER
                        mp_EO_COLUMNMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                        mp_yOperation = E_OPERATION.EO_COLUMNMOVEMENT
                End Select
            Case E_EVENTTARGET.EVT_COLUMN
                mp_EO_COLUMNSELECTION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_COLUMNSELECTION
            Case E_EVENTTARGET.EVT_SELECTEDROW
                If mp_bShowEditTextRow(X, Y) = True Then
                    Return
                End If
                Select Case mp_yRowArea(X, Y)
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
            Case E_EVENTTARGET.EVT_SELECTEDPERCENTAGE
                mp_EO_PERCENTAGESIZING(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_PERCENTAGESIZING
            Case E_EVENTTARGET.EVT_PERCENTAGE
                mp_EO_PERCENTAGESELECTION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_PERCENTAGESELECTION
            Case E_EVENTTARGET.EVT_SELECTEDTASK
                If mp_yOperationModifier = E_OPERATIONMODIFIER.OPM_PREDECESSORMODE Then
                    mp_EO_PREDECESSORADDITION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                    mp_yOperation = E_OPERATION.EO_PREDECESSORADDITION
                Else
                    If mp_bShowEditTextTask(X, Y) = True Then
                        Return
                    End If
                    Select Case mp_yTaskArea(X, Y)
                        Case E_AREA.EA_CENTER
                            mp_EO_TASKMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                            mp_yOperation = E_OPERATION.EO_TASKMOVEMENT
                        Case E_AREA.EA_LEFT
                            mp_EO_TASKSTRETCHLEFT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                            mp_yOperation = E_OPERATION.EO_TASKSTRETCHLEFT
                        Case E_AREA.EA_RIGHT
                            mp_EO_TASKSTRETCHRIGHT(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                            mp_yOperation = E_OPERATION.EO_TASKSTRETCHRIGHT
                        Case E_AREA.EA_NONE
                            mp_yOperation = E_OPERATION.EO_NONE
                    End Select
                End If
            Case E_EVENTTARGET.EVT_TASK
                mp_EO_TASKSELECTION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_TASKSELECTION
            Case E_EVENTTARGET.EVT_SELECTEDPREDECESSOR
                '//
            Case E_EVENTTARGET.EVT_PREDECESSOR
                mp_EO_PREDECESSORSELECTION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                mp_yOperation = E_OPERATION.EO_PREDECESSORSELECTION
            Case E_EVENTTARGET.EVT_CLIENTAREA
                If mp_bShowEditTextTask(X, Y) = True Then
                    Return
                End If
                mp_EO_TASKADDITION(E_MOUSEKEYBOARDEVENTS.MouseDown, X, Y)
                Select Case mp_oControl.AddMode
                    Case E_ADDMODE.AT_TASKADD, E_ADDMODE.AT_DURATION_TASKADD
                        mp_yOperation = E_OPERATION.EO_TASKADDITION
                    Case E_ADDMODE.AT_MILESTONEADD, E_ADDMODE.AT_DURATION_MILESTONEADD
                        mp_yOperation = E_OPERATION.EO_MILESTONEADDITION
                    Case E_ADDMODE.AT_BOTH, E_ADDMODE.AT_DURATION_BOTH
                        mp_yOperation = E_OPERATION.EO_TASKADDITION
                End Select
        End Select
    End Sub

    Friend Sub OnMouseMove(ByVal X As Integer, ByVal Y As Integer)
        Dim yOperation As E_OPERATION = mp_yOperation
        mp_oControl.MouseEventArgs.X = X
        mp_oControl.MouseEventArgs.Y = Y
        mp_oControl.MouseEventArgs.Operation = mp_yOperation
        mp_oControl.FireControlMouseMove()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            mp_yOperation = E_OPERATION.EO_NONE
            Return
        End If
        Select Case mp_yOperation
            Case E_OPERATION.EO_VERTICALSCROLLBAR
                mp_oControl.VerticalScrollBar.ScrollBar.OnMouseMove(X, Y)
            Case E_OPERATION.EO_HORIZONTALSCROLLBAR
                mp_oControl.HorizontalScrollBar.ScrollBar.OnMouseMove(X, Y)
            Case E_OPERATION.EO_TIMELINESCROLLBAR
                mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OnMouseMove(X, Y)
            Case E_OPERATION.EO_TREEVIEW
                mp_oControl.Treeview.OnMouseMove(X, Y)
            Case E_OPERATION.EO_SPLITTERMOVEMENT
                mp_EO_SPLITTERMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TIMELINEMOVEMENT
                mp_EO_TIMELINEMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_COLUMNMOVEMENT
                mp_EO_COLUMNMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_COLUMNSIZING
                mp_EO_COLUMNSIZING(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_COLUMNSELECTION
                mp_EO_COLUMNSELECTION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWMOVEMENT
                mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWSIZING
                mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_ROWSELECTION
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_PERCENTAGESIZING
                mp_EO_PERCENTAGESIZING(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_PERCENTAGESELECTION
                mp_EO_PERCENTAGESELECTION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_PREDECESSORADDITION
                mp_EO_PREDECESSORADDITION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TASKMOVEMENT
                mp_EO_TASKMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TASKSTRETCHLEFT
                mp_EO_TASKSTRETCHLEFT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TASKSTRETCHRIGHT
                mp_EO_TASKSTRETCHRIGHT(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TASKSELECTION
                mp_EO_TASKSELECTION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
            Case E_OPERATION.EO_TASKADDITION
                mp_EO_TASKADDITION(E_MOUSEKEYBOARDEVENTS.MouseMove, X, Y)
        End Select
        System.Diagnostics.Debug.Assert(yOperation = mp_yOperation)
    End Sub

    Friend Sub OnMouseUp(ByVal X As Integer, ByVal Y As Integer)
        mp_oControl.MouseEventArgs.X = X
        mp_oControl.MouseEventArgs.Y = Y
        mp_oControl.FireControlMouseUp()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            mp_yOperation = E_OPERATION.EO_NONE
            Return
        End If
        Select Case mp_yOperation
            Case E_OPERATION.EO_VERTICALSCROLLBAR
                mp_oControl.VerticalScrollBar.ScrollBar.OnMouseUp()
            Case E_OPERATION.EO_HORIZONTALSCROLLBAR
                mp_oControl.HorizontalScrollBar.ScrollBar.OnMouseUp()
            Case E_OPERATION.EO_TIMELINESCROLLBAR
                mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OnMouseUp()
            Case E_OPERATION.EO_TREEVIEW
                mp_oControl.Treeview.OnMouseUp(X, Y)
            Case E_OPERATION.EO_SPLITTERMOVEMENT
                mp_EO_SPLITTERMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TIMELINEMOVEMENT
                mp_EO_TIMELINEMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_COLUMNMOVEMENT
                mp_EO_COLUMNMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_COLUMNSIZING
                mp_EO_COLUMNSIZING(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_COLUMNSELECTION
                mp_EO_COLUMNSELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWMOVEMENT
                mp_EO_ROWMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWSIZING
                mp_EO_ROWSIZING(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_ROWSELECTION
                mp_EO_ROWSELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_PERCENTAGESELECTION
                mp_EO_PERCENTAGESELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_PERCENTAGESIZING
                mp_EO_PERCENTAGESIZING(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_PREDECESSORADDITION
                mp_EO_PREDECESSORADDITION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TASKMOVEMENT
                mp_EO_TASKMOVEMENT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TASKSTRETCHLEFT
                mp_EO_TASKSTRETCHLEFT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TASKSTRETCHRIGHT
                mp_EO_TASKSTRETCHRIGHT(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TASKSELECTION
                mp_EO_TASKSELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_TASKADDITION
                mp_EO_TASKADDITION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
            Case E_OPERATION.EO_PREDECESSORSELECTION
                mp_EO_PREDECESSORSELECTION(E_MOUSEKEYBOARDEVENTS.MouseUp, X, Y)
        End Select
        mp_yOperation = E_OPERATION.EO_NONE
    End Sub

    Private Function CursorPosition(ByVal X As Integer, ByVal Y As Integer) As E_EVENTTARGET
        If mp_oControl.VerticalScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_VSCROLLBAR
        ElseIf mp_oControl.HorizontalScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_HSCROLLBAR
        ElseIf mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TIMELINESCROLLBAR
        ElseIf mp_oControl.Treeview.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TREEVIEW
        ElseIf mp_bOverSplitter(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SPLITTER
        ElseIf mp_bOverEmptySpace(Y) = True Then
            Return E_EVENTTARGET.EVT_NONE
        ElseIf mp_bOverTimeLine(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TIMELINE
        ElseIf mp_bOverSelectedColumn(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDCOLUMN
        ElseIf mp_bOverColumn(X, Y) = True Then
            Return E_EVENTTARGET.EVT_COLUMN
        ElseIf mp_bOverSelectedRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDROW
        ElseIf mp_bOverRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_ROW
        ElseIf mp_bOverSelectedPercentage(X, Y) = True And mp_yOperationModifier <> E_OPERATIONMODIFIER.OPM_PREDECESSORMODE Then
            Return E_EVENTTARGET.EVT_SELECTEDPERCENTAGE
        ElseIf mp_bOverPercentage(X, Y) = True And mp_yOperationModifier <> E_OPERATIONMODIFIER.OPM_PREDECESSORMODE Then
            Return E_EVENTTARGET.EVT_PERCENTAGE
        ElseIf mp_bOverSelectedTask(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDTASK
        ElseIf mp_bOverTask(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TASK
        ElseIf mp_bOverSelectedPredecessor(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDPREDECESSOR
        ElseIf mp_bOverPredecessor(X, Y) = True Then
            Return E_EVENTTARGET.EVT_PREDECESSOR
        ElseIf mp_bOverClientArea(X, Y) = True Then
            Return E_EVENTTARGET.EVT_CLIENTAREA
        Else
            Return E_EVENTTARGET.EVT_NONE
        End If
    End Function

    Private Sub mp_EO_SPLITTERMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        If mp_oControl.AllowSplitterMove = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_MOVESPLITTER)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                mp_oControl.clsG.EraseReversibleFrames()
                mp_DrawMovingReversibleFrame(X, 0, X + 2, 0, E_FOCUSTYPE.FCT_VERTICALSPLITTER)
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                Dim lWidth As Integer = 0
                If X > (mp_oControl.clsG.Width() - 10) Then
                    X = mp_oControl.clsG.Width() - 10
                End If
                If X < 10 Then
                    X = 10
                End If
                lWidth = mp_oControl.Columns.Width
                If (X > lWidth) Then
                    X = lWidth
                    mp_oControl.Splitter.Position = X
                    mp_oControl.HorizontalScrollBar.Value = 0
                ElseIf (X > (lWidth - mp_oControl.HorizontalScrollBar.Value)) Then
                    X = lWidth - mp_oControl.HorizontalScrollBar.Value
                    mp_oControl.Splitter.Position = X
                Else
                    mp_oControl.Splitter.Position = X
                End If

                mp_oControl.clsG.EraseReversibleFrames()
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TIMELINEMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        If mp_oControl.AllowTimeLineScroll = False Then
            Return
        End If
        If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Enabled = True Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                mp_SetCursor(E_CURSORTYPE.CT_SCROLLTIMELINE)
                s_tmlnMVT.lX = X
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.Interval, (s_tmlnMVT.lX - X) * mp_oControl.CurrentViewObject.Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_COLUMNMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oColumn As clsColumn = Nothing
        Dim oDestinationColumn As clsColumn = Nothing
        Dim oRow As clsRow = Nothing
        Dim lIndex As Integer
        If mp_oControl.AllowColumnMove = False Then
            Return
        End If
        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If oColumn.AllowMove = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_MOVECOLUMN)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_clmnMVT.lColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
                System.Diagnostics.Debug.Assert(s_clmnMVT.lColumnIndex >= 1)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_COLUMN
                mp_oControl.ObjectStateChangedEventArgs.Index = s_clmnMVT.lColumnIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectMove()
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_DynamicColumnMove(X)
                    s_clmnMVT.lDestinationColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
                    If s_clmnMVT.lDestinationColumnIndex >= 1 Then
                        mp_oControl.clsG.EraseReversibleFrames()
                        mp_oControl.ObjectStateChangedEventArgs.DestinationIndex = s_clmnMVT.lDestinationColumnIndex
                        mp_oControl.FireObjectMove()
                        If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                            oDestinationColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(s_clmnMVT.lDestinationColumnIndex), clsColumn)
                            mp_DrawMovingReversibleFrame(oDestinationColumn.LeftTrim, oDestinationColumn.Top, oDestinationColumn.RightTrim, oDestinationColumn.Bottom, E_FOCUSTYPE.FCT_NORMAL)
                        End If
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireEndObjectMove()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        s_clmnMVT.lDestinationColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
                        If s_clmnMVT.lDestinationColumnIndex >= 1 And (s_clmnMVT.lColumnIndex <> s_clmnMVT.lDestinationColumnIndex) Then
                            mp_oControl.FireCompleteObjectMove()
                            If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                                If mp_oControl.TreeviewColumnIndex > 0 Then
                                    For lIndex = 1 To mp_oControl.Columns.Count
                                        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
                                        If lIndex = mp_oControl.TreeviewColumnIndex Then
                                            oColumn.mp_bTreeViewColumnIndex = True
                                        Else
                                            oColumn.mp_bTreeViewColumnIndex = False
                                        End If
                                    Next
                                End If
                                mp_oControl.SelectedColumnIndex = mp_oControl.Columns.oCollection.m_lCopyAndMoveItems(s_clmnMVT.lColumnIndex, s_clmnMVT.lDestinationColumnIndex)
                                For lIndex = 1 To mp_oControl.Rows.Count
                                    oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
                                    oRow.Cells.oCollection.m_lCopyAndMoveItems(s_clmnMVT.lColumnIndex, s_clmnMVT.lDestinationColumnIndex)
                                Next lIndex
                                If mp_oControl.TreeviewColumnIndex > 0 Then
                                    For lIndex = 1 To mp_oControl.Columns.Count
                                        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
                                        If oColumn.mp_bTreeViewColumnIndex = True Then
                                            mp_oControl.TreeviewColumnIndex = lIndex
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_COLUMNSIZING(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oColumn As clsColumn = Nothing
        If mp_oControl.AllowColumnSize = False Then
            Return
        End If
        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If oColumn.AllowSize = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_COLUMNWIDTH)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_clmnSZ.lColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
                System.Diagnostics.Debug.Assert(s_clmnSZ.lColumnIndex >= 1)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_COLUMN
                mp_oControl.ObjectStateChangedEventArgs.Index = s_clmnSZ.lColumnIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectSize()
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        mp_DrawMovingReversibleFrame(X, 0, X + 2, mp_oControl.clsG.Height(), E_FOCUSTYPE.FCT_NORMAL)
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireEndObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        If X < mp_oControl.mt_BorderThickness Then
                            X = mp_oControl.mt_BorderThickness
                        End If
                        If X > mp_oControl.Splitter.Position Then
                            mp_oControl.Splitter.Position = X
                        End If
                        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(s_clmnSZ.lColumnIndex), clsColumn)
                        oColumn.Width = oColumn.Width + (X - oColumn.Right)
                        If oColumn.Width < mp_oControl.MinColumnWidth Then
                            oColumn.Width = mp_oControl.MinColumnWidth
                        End If
                        If mp_oControl.Splitter.Position > mp_oControl.Columns.Width Then
                            mp_oControl.Splitter.Position = mp_oControl.Columns.Width
                            mp_oControl.HorizontalScrollBar.Value = 0
                        End If
                        mp_oControl.FireCompleteObjectSize()
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_COLUMNSELECTION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_clmnSEL.lColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedColumnIndex = s_clmnSEL.lColumnIndex
                mp_oControl.ObjectSelectedEventArgs.Clear()
                mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_COLUMN
                mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedColumnIndex
                mp_oControl.FireObjectSelected()
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_ROWMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        Dim oDestinationRow As clsRow = Nothing
        If mp_oControl.AllowRowMove = False Then
            Return
        End If
        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If oRow.AllowMove = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_MOVEROW)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_rowMVT.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                System.Diagnostics.Debug.Assert(s_rowMVT.lRowIndex >= 1)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                mp_oControl.ObjectStateChangedEventArgs.Index = s_rowMVT.lRowIndex
                mp_oControl.ObjectStateChangedEventArgs.InitialRowIndex = s_rowMVT.lRowIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectMove()
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_DynamicRowMove(Y)
                    s_rowMVT.lDestinationRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                    If s_rowMVT.lDestinationRowIndex >= 1 Then
                        mp_oControl.clsG.EraseReversibleFrames()
                        mp_oControl.ObjectStateChangedEventArgs.DestinationIndex = s_rowMVT.lDestinationRowIndex
                        mp_oControl.ObjectStateChangedEventArgs.FinalRowIndex = s_rowMVT.lDestinationRowIndex
                        mp_oControl.FireObjectMove()
                        If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                            oDestinationRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_rowMVT.lDestinationRowIndex), clsRow)
                            mp_DrawMovingReversibleFrame(oDestinationRow.Left, oDestinationRow.Top, oDestinationRow.Right, oDestinationRow.Bottom, E_FOCUSTYPE.FCT_NORMAL)
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
                            mp_oControl.SelectedRowIndex = mp_oControl.Rows.oCollection.m_lCopyAndMoveItems(s_rowMVT.lRowIndex, s_rowMVT.lDestinationRowIndex)
                            mp_oControl.FireCompleteObjectMove()
                        End If
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
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
        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If oRow.AllowSize = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_ROWHEIGHT)
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
                        mp_DrawMovingReversibleFrame(0, Y, mp_oControl.clsG.Width(), Y + 2, E_FOCUSTYPE.FCT_NORMAL)
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
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
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
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_rowSEL.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                s_rowSEL.lCellIndex = mp_oControl.MathLib.GetCellIndexByPosition(X)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedRowIndex = s_rowSEL.lRowIndex
                mp_oControl.SelectedCellIndex = s_rowSEL.lCellIndex
                oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
                If oRow.MergeCells = True Then
                    mp_oControl.ObjectSelectedEventArgs.Clear()
                    mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_ROW
                    mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedRowIndex
                    mp_oControl.FireObjectSelected()
                Else
                    mp_oControl.ObjectSelectedEventArgs.Clear()
                    mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_CELL
                    mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedCellIndex
                    mp_oControl.ObjectSelectedEventArgs.ParentObjectIndex = mp_oControl.SelectedRowIndex
                    mp_oControl.FireObjectSelected()
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TASKMOVEMENT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oTask As clsTask = Nothing
        Dim oRow As clsRow = Nothing
        If mp_oControl.AllowEdit = False Then
            Return
        End If
        oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If oTask.AllowedMovement = E_MOVEMENTTYPE.MT_MOVEMENTDISABLED Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_MOVETASK)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_tskMVT.Clear()
                s_tskMVT.lInitialRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                X = mp_fSnapX(X)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                mp_oControl.ObjectStateChangedEventArgs.Index = mp_oControl.SelectedTaskIndex
                mp_oControl.ObjectStateChangedEventArgs.InitialRowIndex = s_tskMVT.lInitialRowIndex
                mp_oControl.ObjectStateChangedEventArgs.StartDate = oTask.StartDate
                mp_oControl.ObjectStateChangedEventArgs.EndDate = oTask.EndDate
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectMove()
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    s_tskMVT.lDeltaLeft = X - mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate)
                    s_tskMVT.lDeltaRight = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate) - X
                    s_tskMVT.dtInitialStartDate = oTask.StartDate
                    s_tskMVT.dtInitialEndDate = oTask.EndDate
                    s_tskMVT.sInitialRowKey = oTask.RowKey
                    s_tskMVT.lDurationFactor = oTask.DurationFactor
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If oTask.TaskType = E_TASKTYPE.TT_START_END Then
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        X = mp_fSnapX(X)
                        s_tskMVT.lFinalRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                        mp_oControl.clsG.EraseReversibleFrames()
                        mp_oControl.ObjectStateChangedEventArgs.FinalRowIndex = s_tskMVT.lFinalRowIndex
                        mp_oControl.ObjectStateChangedEventArgs.StartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft)
                        mp_oControl.ObjectStateChangedEventArgs.EndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight)
                        mp_oControl.FireObjectMove()
                        mp_DynamicRowMove(Y)
                        If mp_oControl.ObjectStateChangedEventArgs.Cancel = False And (s_tskMVT.lFinalRowIndex >= 1 And s_tskMVT.lFinalRowIndex <= mp_oControl.Rows.Count) Then
                            If X < 0 Or X > mp_oControl.clsG.Width() Or Y < 0 Or Y > mp_oControl.clsG.Height() Then
                                mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                                mp_oControl.clsG.DrawReversibleFrameEx()
                            End If
                            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskMVT.lFinalRowIndex), clsRow)
                            If (s_tskMVT.lInitialRowIndex <> s_tskMVT.lFinalRowIndex And oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                                mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                                mp_oControl.clsG.DrawReversibleFrameEx()
                                Return
                            End If
                            mp_SetCursor(E_CURSORTYPE.CT_MOVETASK)
                            mp_DynamicTimeLineMove(X)
                            mp_oControl.ToolTipEventArgs.Clear()
                            mp_oControl.ToolTipEventArgs.TaskIndex = mp_oControl.SelectedTaskIndex
                            mp_oControl.ToolTipEventArgs.InitialRowIndex = s_tskMVT.lInitialRowIndex
                            mp_oControl.ToolTipEventArgs.FinalRowIndex = s_tskMVT.lFinalRowIndex
                            mp_oControl.ToolTipEventArgs.InitialStartDate = s_tskMVT.dtInitialStartDate
                            mp_oControl.ToolTipEventArgs.InitialEndDate = s_tskMVT.dtInitialEndDate
                            mp_oControl.ToolTipEventArgs.StartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft)
                            mp_oControl.ToolTipEventArgs.EndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight)
                            mp_oControl.ToolTipEventArgs.X = X
                            mp_oControl.ToolTipEventArgs.Y = Y
                            mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                            mp_DrawMovingReversibleFrame(X - s_tskMVT.lDeltaLeft, oRow.Top, X + s_tskMVT.lDeltaRight, oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                        End If
                    End If
                ElseIf oTask.TaskType = E_TASKTYPE.TT_DURATION Then
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        Dim dtStartDate As AGVBE.DateTime
                        Dim dtEndDate As AGVBE.DateTime
                        X = mp_fSnapX(X)
                        s_tskMVT.lFinalRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                        mp_oControl.clsG.EraseReversibleFrames()
                        mp_oControl.ObjectStateChangedEventArgs.FinalRowIndex = s_tskMVT.lFinalRowIndex
                        dtStartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft)
                        dtEndDate = mp_oControl.MathLib.GetEndDate(dtStartDate, oTask.DurationInterval, oTask.DurationFactor)
                        mp_oControl.ObjectStateChangedEventArgs.StartDate = dtStartDate
                        mp_oControl.ObjectStateChangedEventArgs.EndDate = dtEndDate
                        mp_oControl.FireObjectMove()
                        mp_DynamicRowMove(Y)
                        If mp_oControl.ObjectStateChangedEventArgs.Cancel = False And (s_tskMVT.lFinalRowIndex >= 1 And s_tskMVT.lFinalRowIndex <= mp_oControl.Rows.Count) Then
                            If X < 0 Or X > mp_oControl.clsG.Width() Or Y < 0 Or Y > mp_oControl.clsG.Height() Then
                                mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                                mp_oControl.clsG.DrawReversibleFrameEx()
                            End If
                            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskMVT.lFinalRowIndex), clsRow)
                            If (s_tskMVT.lInitialRowIndex <> s_tskMVT.lFinalRowIndex And oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                                mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                                mp_oControl.clsG.DrawReversibleFrameEx()
                                Return
                            End If
                            mp_SetCursor(E_CURSORTYPE.CT_MOVETASK)
                            mp_DynamicTimeLineMove(X)
                            mp_oControl.ToolTipEventArgs.Clear()
                            mp_oControl.ToolTipEventArgs.TaskIndex = mp_oControl.SelectedTaskIndex
                            mp_oControl.ToolTipEventArgs.InitialRowIndex = s_tskMVT.lInitialRowIndex
                            mp_oControl.ToolTipEventArgs.FinalRowIndex = s_tskMVT.lFinalRowIndex
                            mp_oControl.ToolTipEventArgs.InitialStartDate = s_tskMVT.dtInitialStartDate
                            mp_oControl.ToolTipEventArgs.InitialEndDate = s_tskMVT.dtInitialEndDate
                            mp_oControl.ToolTipEventArgs.StartDate = dtStartDate
                            mp_oControl.ToolTipEventArgs.EndDate = dtEndDate
                            mp_oControl.ToolTipEventArgs.X = X
                            mp_oControl.ToolTipEventArgs.Y = Y
                            mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                            mp_DrawMovingReversibleFrame(mp_oControl.MathLib.GetXCoordinateFromDate(dtStartDate), oRow.Top, mp_oControl.MathLib.GetXCoordinateFromDate(dtEndDate), oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                        End If
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If oTask.TaskType = E_TASKTYPE.TT_START_END Then
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False And (s_tskMVT.lFinalRowIndex >= 1 And s_tskMVT.lFinalRowIndex <= mp_oControl.Rows.Count) Then
                        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskMVT.lFinalRowIndex), clsRow)
                        mp_oControl.clsG.EraseReversibleFrames()
                        If (s_tskMVT.lInitialRowIndex <> s_tskMVT.lFinalRowIndex And oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                        Else
                            X = mp_fSnapX(X)
                            mp_oControl.FireEndObjectMove()
                            If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                                oTask.StartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft)
                                oTask.EndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X + s_tskMVT.lDeltaRight)
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.StartDate)
                                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.EndDate)
                                End If
                                oTask.RowKey = oRow.Key
                                If oTask.StartDate <> s_tskMVT.dtInitialStartDate Or oTask.EndDate <> s_tskMVT.dtInitialEndDate Or oTask.RowKey <> s_tskMVT.sInitialRowKey Then
                                    mp_oControl.FireCompleteObjectMove()
                                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = True Then
                                        '//Account for Duration
                                        oTask.StartDate = s_tskMVT.dtInitialStartDate
                                        oTask.EndDate = s_tskMVT.dtInitialEndDate
                                        oTask.RowKey = s_tskMVT.sInitialRowKey
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If mp_oControl.EnforcePredecessors = True Then
                        mp_oControl.CheckPredecessors()
                    End If
                    mp_oControl.Redraw()
                    mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
                ElseIf oTask.TaskType = E_TASKTYPE.TT_DURATION Then
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False And (s_tskMVT.lFinalRowIndex >= 1 And s_tskMVT.lFinalRowIndex <= mp_oControl.Rows.Count) Then
                        Dim dtStartDate As AGVBE.DateTime
                        Dim dtEndDate As AGVBE.DateTime
                        Dim lDuration As Integer
                        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskMVT.lFinalRowIndex), clsRow)
                        mp_oControl.clsG.EraseReversibleFrames()
                        If (s_tskMVT.lInitialRowIndex <> s_tskMVT.lFinalRowIndex And oTask.AllowedMovement = E_MOVEMENTTYPE.MT_RESTRICTEDTOROW) Or (oRow.Container = False) Or (mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X - s_tskMVT.lDeltaLeft)), mp_oControl.MathLib.GetDateFromXCoordinate(mp_fSnapX(X + s_tskMVT.lDeltaRight)), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                        Else
                            X = mp_fSnapX(X)
                            mp_oControl.FireEndObjectMove()
                            If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                                dtStartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X - s_tskMVT.lDeltaLeft)
                                dtEndDate = mp_oControl.MathLib.GetEndDate(dtStartDate, oTask.DurationInterval, oTask.DurationFactor)
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    dtStartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, dtStartDate)
                                End If
                                lDuration = mp_oControl.MathLib.CalculateDuration(dtStartDate, dtEndDate, oTask.DurationInterval)
                                If lDuration <> s_tskMVT.lDurationFactor Then
                                    mp_oControl.mp_ErrorReport(SYS_ERRORS.ERR_DURATION_INCONSISTENT, "Inconsistent duration", "clsMouseKeyboardEvents.mp_EO_TASKMOVEMENT")
                                End If
                                oTask.StartDate = dtStartDate
                                oTask.RowKey = oRow.Key
                                If oTask.StartDate <> s_tskMVT.dtInitialStartDate Or oTask.EndDate <> s_tskMVT.dtInitialEndDate Or oTask.RowKey <> s_tskMVT.sInitialRowKey Then
                                    mp_oControl.FireCompleteObjectMove()
                                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = True Then
                                        oTask.StartDate = s_tskMVT.dtInitialStartDate
                                        oTask.RowKey = s_tskMVT.sInitialRowKey
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If mp_oControl.EnforcePredecessors = True Then
                        mp_oControl.CheckPredecessors()
                    End If
                    mp_oControl.Redraw()
                    mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TASKSTRETCHLEFT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oTask As clsTask = Nothing
        Dim oRow As clsRow = Nothing
        If mp_oControl.AllowEdit = False Then
            Return
        End If
        oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If oTask.AllowStretchLeft = False Or oTask.TaskType = E_TASKTYPE.TT_DURATION Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_SIZETASK)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                s_tskSTL.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                X = mp_fSnapX(X)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                mp_oControl.ObjectStateChangedEventArgs.Index = mp_oControl.SelectedTaskIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectSize()
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    s_tskSTL.dtInitialStartDate = oTask.StartDate
                    s_tskSTL.dtInitialEndDate = oTask.EndDate
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    X = mp_fSnapX(X)
                    mp_oControl.clsG.EraseReversibleFrames()
                    oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskSTL.lRowIndex), clsRow)
                    s_tskSTL.dtFinalStartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                    mp_oControl.FireObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        If mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), oTask.EndDate, oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                            mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                        Else
                            mp_SetCursor(E_CURSORTYPE.CT_SIZETASK)
                        End If
                        mp_DynamicTimeLineMove(X)
                        mp_oControl.ToolTipEventArgs.Clear()
                        mp_oControl.ToolTipEventArgs.TaskIndex = mp_oControl.SelectedTaskIndex
                        mp_oControl.ToolTipEventArgs.RowIndex = s_tskSTL.lRowIndex
                        mp_oControl.ToolTipEventArgs.InitialStartDate = s_tskSTL.dtInitialStartDate
                        mp_oControl.ToolTipEventArgs.InitialEndDate = s_tskSTL.dtInitialEndDate
                        mp_oControl.ToolTipEventArgs.StartDate = s_tskSTL.dtFinalStartDate
                        mp_oControl.ToolTipEventArgs.EndDate = s_tskSTL.dtInitialEndDate
                        mp_oControl.ToolTipEventArgs.X = X
                        mp_oControl.ToolTipEventArgs.Y = Y
                        mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                        mp_DrawMovingReversibleFrame(X, oRow.Top, mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate), oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If s_tskSTL.dtFinalStartDate >= s_tskSTL.dtInitialEndDate Then
                    mp_oControl.ObjectStateChangedEventArgs.Cancel = True
                End If
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    If s_tskSTL.dtFinalStartDate < mp_oControl.CurrentViewObject.TimeLine.StartDate Then
                        s_tskSTL.dtFinalStartDate = mp_oControl.CurrentViewObject.TimeLine.StartDate
                    End If
                    mp_oControl.FireEndObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskSTL.lRowIndex), clsRow)
                        If mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), oTask.EndDate, oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                        Else
                            oTask.StartDate = s_tskSTL.dtFinalStartDate
                            If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.StartDate)
                                oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.EndDate)
                            End If
                            If oTask.StartDate <> s_tskSTL.dtInitialStartDate Then
                                mp_oControl.FireCompleteObjectSize()
                                If mp_oControl.ObjectStateChangedEventArgs.Cancel = True Then
                                    oTask.StartDate = s_tskSTL.dtInitialStartDate
                                End If
                            End If
                        End If
                    End If
                End If
                If mp_oControl.EnforcePredecessors = True Then
                    mp_oControl.CheckPredecessors()
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TASKSTRETCHRIGHT(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oTask As clsTask = Nothing
        Dim oRow As clsRow = Nothing
        If mp_oControl.AllowEdit = False Then
            Return
        End If
        oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If oTask.AllowStretchRight = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_SIZETASK)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_tskSTR.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
                X = mp_fSnapX(X)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                mp_oControl.ObjectStateChangedEventArgs.Index = mp_oControl.SelectedTaskIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                mp_oControl.FireBeginObjectSize()
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    s_tskSTR.dtInitialStartDate = oTask.StartDate
                    s_tskSTR.dtInitialEndDate = oTask.EndDate
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    X = mp_fSnapX(X)
                    mp_oControl.clsG.EraseReversibleFrames()
                    oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskSTR.lRowIndex), clsRow)
                    s_tskSTR.dtFinalEndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                    mp_oControl.FireObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        If mp_oControl.MathLib.DetectConflict(oTask.StartDate, mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                            mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                        Else
                            mp_SetCursor(E_CURSORTYPE.CT_SIZETASK)
                        End If
                        mp_DynamicTimeLineMove(X)
                        mp_oControl.ToolTipEventArgs.Clear()
                        mp_oControl.ToolTipEventArgs.TaskIndex = mp_oControl.SelectedTaskIndex
                        mp_oControl.ToolTipEventArgs.RowIndex = s_tskSTR.lRowIndex
                        mp_oControl.ToolTipEventArgs.InitialStartDate = s_tskSTR.dtInitialStartDate
                        mp_oControl.ToolTipEventArgs.InitialEndDate = s_tskSTR.dtInitialEndDate
                        mp_oControl.ToolTipEventArgs.StartDate = s_tskSTR.dtInitialStartDate
                        mp_oControl.ToolTipEventArgs.EndDate = s_tskSTR.dtFinalEndDate
                        mp_oControl.ToolTipEventArgs.X = X
                        mp_oControl.ToolTipEventArgs.Y = Y
                        mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                        mp_DrawMovingReversibleFrame(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oRow.Top, X, oRow.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If s_tskSTR.dtFinalEndDate <= s_tskSTR.dtInitialStartDate Then
                    mp_oControl.ObjectStateChangedEventArgs.Cancel = True
                End If
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    If s_tskSTR.dtFinalEndDate > mp_oControl.CurrentViewObject.TimeLine.EndDate Then
                        s_tskSTR.dtFinalEndDate = mp_oControl.CurrentViewObject.TimeLine.EndDate
                    End If
                    mp_oControl.FireEndObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskSTR.lRowIndex), clsRow)
                        If mp_oControl.MathLib.DetectConflict(oTask.StartDate, mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, mp_oControl.SelectedTaskIndex, oTask.LayerIndex) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                        Else
                            If oTask.TaskType = E_TASKTYPE.TT_START_END Then
                                oTask.EndDate = s_tskSTR.dtFinalEndDate
                                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True Then
                                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.StartDate)
                                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.EndDate)
                                End If
                            ElseIf oTask.TaskType = E_TASKTYPE.TT_DURATION Then
                                oTask.DurationFactor = mp_oControl.MathLib.CalculateDuration(s_tskSTR.dtInitialStartDate, s_tskSTR.dtFinalEndDate, oTask.DurationInterval)
                            End If
                            If oTask.EndDate <> s_tskSTR.dtInitialEndDate Then
                                mp_oControl.FireCompleteObjectSize()
                                If mp_oControl.ObjectStateChangedEventArgs.Cancel = True Then
                                    oTask.EndDate = s_tskSTR.dtInitialEndDate
                                End If
                            End If
                        End If
                    End If
                End If
                If mp_oControl.EnforcePredecessors = True Then
                    mp_oControl.CheckPredecessors()
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TASKSELECTION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oTask As clsTask = Nothing
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                'ToolTipText(E_OPERATION.EO_TASKSELECTION, mp_oControl.MathLib.GetTaskIndexByPosition(X, Y), X, Y, Nothing, Nothing, "")
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_tskSEL.lTaskIndex = mp_oControl.MathLib.GetTaskIndexByPosition(X, Y)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedTaskIndex = s_tskSEL.lTaskIndex
                oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
                If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = True And mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGridOnSelection = True Then
                    oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
                    oTask.StartDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.StartDate)
                    oTask.EndDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, oTask.EndDate)
                End If
                mp_oControl.ObjectSelectedEventArgs.Clear()
                mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedTaskIndex
                mp_oControl.FireObjectSelected()
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_TASKADDITION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oRow As clsRow = Nothing
        If mp_oControl.AllowAdd = False Then
            mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Return
        End If
        s_tskADD.lRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
        If s_tskADD.lRowIndex <= 0 Then
            Return
        End If
        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(s_tskADD.lRowIndex), clsRow)
        If oRow.Container = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_CLIENTAREA)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_tskADD.bCancel = False
                X = mp_fSnapX(X)
                If (oRow.Container = False) Or (mp_oControl.MathLib.DetectConflict(mp_oControl.MathLib.GetDateFromXCoordinate(X), mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, 0, mp_oControl.CurrentLayer) = True And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True) Then
                    s_tskADD.bCancel = True
                Else
                    s_tskADD.dtStartDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If s_tskADD.bCancel = False Then
                    X = mp_fSnapX(X)
                    mp_oControl.clsG.EraseReversibleFrames()
                    If (mp_oControl.MathLib.DetectConflict(s_tskADD.dtStartDate, mp_oControl.MathLib.GetDateFromXCoordinate(X), oRow.Key, 0, mp_oControl.CurrentLayer) = True Or oRow.Container = False) And mp_oControl.CurrentViewObject.ClientArea.DetectConflicts = True Then
                        mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                        s_tskADD.bInConflict = True
                    Else
                        s_tskADD.bInConflict = False
                        mp_SetCursor(E_CURSORTYPE.CT_CLIENTAREA)
                        s_tskADD.dtEndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                        mp_DynamicTimeLineMove(X)
                        mp_oControl.ToolTipEventArgs.Clear()
                        mp_oControl.ToolTipEventArgs.RowIndex = s_tskADD.lRowIndex
                        mp_oControl.ToolTipEventArgs.StartDate = s_tskADD.dtStartDate
                        mp_oControl.ToolTipEventArgs.EndDate = s_tskADD.dtEndDate
                        mp_oControl.ToolTipEventArgs.X = X
                        mp_oControl.ToolTipEventArgs.Y = Y
                        mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                        mp_DrawMovingReversibleFrame(mp_oControl.MathLib.GetXCoordinateFromDate(s_tskADD.dtStartDate), oRow.Top, X, oRow.Bottom, E_FOCUSTYPE.FCT_ADD)
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.clsG.EraseReversibleFrames()
                If s_tskADD.bCancel = False And s_tskADD.bInConflict = False Then
                    X = mp_fSnapX(X)
                    s_tskADD.dtEndDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
                    If s_tskADD.dtEndDate = s_tskADD.dtStartDate Then
                        If mp_oControl.AddMode = E_ADDMODE.AT_BOTH Or mp_oControl.AddMode = E_ADDMODE.AT_MILESTONEADD Then
                            mp_oControl.Tasks.Add("", oRow.Key, s_tskADD.dtEndDate, s_tskADD.dtStartDate, "", "DS_TASK", mp_oControl.CurrentLayer)
                        ElseIf mp_oControl.AddMode = E_ADDMODE.AT_DURATION_BOTH Or mp_oControl.AddMode = E_ADDMODE.AT_DURATION_MILESTONEADD Then
                            Dim dtStartDate As AGVBE.DateTime = s_tskADD.dtStartDate
                            Dim dtTLStartDate As AGVBE.DateTime = mp_oControl.CurrentViewObject.TimeLine.StartDate
                            Dim dtTLEndDate As AGVBE.DateTime = mp_oControl.CurrentViewObject.TimeLine.EndDate
                            Dim aTimeBlocks As ArrayList = New ArrayList()
                            mp_oControl.MathLib.mp_GetTimeBlocks(aTimeBlocks, dtTLStartDate, dtTLEndDate)
                            mp_oControl.MathLib.mp_ValidateStartDate(aTimeBlocks, dtStartDate)
                            mp_oControl.Tasks.DAdd(oRow.Key, dtStartDate, mp_oControl.AddDurationInterval, 0, "", "", "DS_TASK", mp_oControl.CurrentLayer)
                        End If
                        mp_oControl.SelectedTaskIndex = mp_oControl.Tasks.Count
                        mp_oControl.ObjectAddedEventArgs.Clear()
                        mp_oControl.ObjectAddedEventArgs.TaskIndex = mp_oControl.Tasks.Count
                        mp_oControl.ObjectAddedEventArgs.EventTarget = E_EVENTTARGET.EVT_MILESTONE
                        mp_oControl.FireObjectAdded()
                    Else
                        If mp_oControl.AddMode = E_ADDMODE.AT_BOTH Or mp_oControl.AddMode = E_ADDMODE.AT_TASKADD Then
                            If s_tskADD.dtEndDate < s_tskADD.dtStartDate Then
                                mp_oControl.Tasks.Add("", oRow.Key, s_tskADD.dtEndDate, s_tskADD.dtStartDate, "", "DS_TASK", mp_oControl.CurrentLayer)
                            Else
                                mp_oControl.Tasks.Add("", oRow.Key, s_tskADD.dtStartDate, s_tskADD.dtEndDate, "", "DS_TASK", mp_oControl.CurrentLayer)
                            End If
                            mp_oControl.SelectedTaskIndex = mp_oControl.Tasks.Count
                            mp_oControl.ObjectAddedEventArgs.Clear()
                            mp_oControl.ObjectAddedEventArgs.TaskIndex = mp_oControl.Tasks.Count
                            mp_oControl.ObjectAddedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                            mp_oControl.FireObjectAdded()
                        ElseIf mp_oControl.AddMode = E_ADDMODE.AT_DURATION_BOTH Or mp_oControl.AddMode = E_ADDMODE.AT_DURATION_TASKADD Then
                            Dim lDuration As Integer = 0
                            Dim dtStartDate As AGVBE.DateTime
                            Dim dtEndDate As AGVBE.DateTime
                            If s_tskADD.dtEndDate > s_tskADD.dtStartDate Then
                                dtStartDate = s_tskADD.dtStartDate
                                dtEndDate = s_tskADD.dtEndDate
                            Else
                                dtStartDate = s_tskADD.dtEndDate
                                dtEndDate = s_tskADD.dtStartDate
                            End If
                            lDuration = mp_oControl.MathLib.CalculateDuration(dtStartDate, dtEndDate, mp_oControl.AddDurationInterval)
                            mp_oControl.Tasks.DAdd(oRow.Key, dtStartDate, mp_oControl.AddDurationInterval, lDuration, "", "", "DS_TASK", mp_oControl.CurrentLayer)
                            mp_oControl.SelectedTaskIndex = mp_oControl.Tasks.Count
                            mp_oControl.ObjectAddedEventArgs.Clear()
                            mp_oControl.ObjectAddedEventArgs.TaskIndex = mp_oControl.Tasks.Count
                            mp_oControl.ObjectAddedEventArgs.EventTarget = E_EVENTTARGET.EVT_TASK
                            mp_oControl.FireObjectAdded()
                        End If
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_PERCENTAGESELECTION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oPercentage As clsPercentage = Nothing
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_perSEL.lPercentageIndex = mp_oControl.MathLib.GetPercentageIndexByPosition(X, Y)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedPercentageIndex = s_perSEL.lPercentageIndex
                oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(mp_oControl.SelectedPercentageIndex), clsPercentage)
                mp_oControl.ObjectSelectedEventArgs.Clear()
                mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_PERCENTAGE
                mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedTaskIndex
                mp_oControl.FireObjectSelected()
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_PERCENTAGESIZING(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oPercentage As clsPercentage = Nothing
        Dim oTask As clsTask = Nothing
        If mp_oControl.AllowEdit = False Then
            Return
        End If
        oPercentage = DirectCast(mp_oControl.Percentages.Item(mp_oControl.SelectedPercentageIndex.ToString()), clsPercentage)
        If oPercentage.AllowSize = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_PERCENTAGE)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_perSZ.bMouseMove = False
                oTask = DirectCast(mp_oControl.Tasks.Item(oPercentage.TaskKey), clsTask)
                mp_oControl.ObjectStateChangedEventArgs.Clear()
                mp_oControl.ObjectStateChangedEventArgs.EventTarget = E_EVENTTARGET.EVT_PERCENTAGE
                mp_oControl.ObjectStateChangedEventArgs.Index = mp_oControl.SelectedPercentageIndex
                mp_oControl.ObjectStateChangedEventArgs.Cancel = False
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    s_perSZ.lTaskIndex = oTask.Index
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                s_perSZ.bMouseMove = True
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    oTask = DirectCast(mp_oControl.Tasks.Item(oPercentage.TaskKey), clsTask)
                    mp_oControl.FireObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        mp_DynamicTimeLineMove(X)
                        s_perSZ.lTaskStart = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate)
                        s_perSZ.lTaskEnd = mp_oControl.MathLib.GetXCoordinateFromDate(oTask.EndDate)
                        If X < s_perSZ.lTaskStart Then
                            X = s_perSZ.lTaskStart
                        End If
                        If X > s_perSZ.lTaskEnd Then
                            X = s_perSZ.lTaskEnd
                        End If
                        Dim fPercent As Single = 0
                        fPercent = mp_oControl.MathLib.PercentageComplete(s_perSZ.lTaskStart, s_perSZ.lTaskEnd, X)
                        fPercent = mp_oControl.MathLib.RoundDouble(fPercent * 100)
                        mp_oControl.ToolTipEventArgs.Clear()
                        mp_oControl.ToolTipEventArgs.PercentageIndex = mp_oControl.SelectedPercentageIndex
                        mp_oControl.ToolTipEventArgs.XStart = s_perSZ.lTaskStart
                        mp_oControl.ToolTipEventArgs.XEnd = s_perSZ.lTaskEnd
                        mp_oControl.ToolTipEventArgs.TaskIndex = oTask.Index
                        mp_oControl.ToolTipEventArgs.RowIndex = mp_oControl.Rows.Item(oTask.RowKey).Index
                        mp_oControl.ToolTipEventArgs.StartDate = oTask.StartDate
                        mp_oControl.ToolTipEventArgs.EndDate = oTask.EndDate
                        mp_oControl.ToolTipEventArgs.X = X
                        mp_oControl.ToolTipEventArgs.Y = Y
                        mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                        mp_DrawMovingReversibleFrame(s_perSZ.lTaskStart, oPercentage.Top, X, oPercentage.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                        s_perSZ.lX = X
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                If s_perSZ.bMouseMove = False Then
                    Return
                End If
                If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.FireEndObjectSize()
                    If mp_oControl.ObjectStateChangedEventArgs.Cancel = False Then
                        s_perSZ.lTaskEnd = s_perSZ.lTaskEnd - s_perSZ.lTaskStart
                        s_perSZ.lX = s_perSZ.lX - s_perSZ.lTaskStart
                        If s_perSZ.lX = 0 Then
                            oPercentage.Percent = 0
                        ElseIf s_perSZ.lX = s_perSZ.lTaskEnd Then
                            oPercentage.Percent = 1
                        Else
                            oPercentage.Percent = System.Convert.ToSingle(s_perSZ.lX / s_perSZ.lTaskEnd)
                        End If
                        mp_oControl.FireCompleteObjectSize()
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_PREDECESSORSELECTION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_preSEL.lPredecessorIndex = mp_oControl.MathLib.GetPredecessorIndexByPosition(X, Y)
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                mp_oControl.SelectedPredecessorIndex = s_preSEL.lPredecessorIndex
                mp_oControl.ObjectSelectedEventArgs.Clear()
                mp_oControl.ObjectSelectedEventArgs.EventTarget = E_EVENTTARGET.EVT_PREDECESSOR
                mp_oControl.ObjectSelectedEventArgs.ObjectIndex = mp_oControl.SelectedPredecessorIndex
                mp_oControl.FireObjectSelected()
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Sub mp_EO_PREDECESSORADDITION(ByVal yMouseKeyBoardEvent As E_MOUSEKEYBOARDEVENTS, ByVal X As Integer, ByVal Y As Integer)
        Dim oTask As clsTask = Nothing
        Dim oPredecessor As clsTask = Nothing
        If mp_oControl.AllowPredecessorAdd = False Then
            Return
        End If
        Select Case yMouseKeyBoardEvent
            Case E_MOUSEKEYBOARDEVENTS.MouseHover
                mp_SetCursor(E_CURSORTYPE.CT_PREDECESSOR)
            Case E_MOUSEKEYBOARDEVENTS.MouseDown
                s_preADD.bCancel = False
                s_preADD.lXStart = X
                s_preADD.lYStart = Y
                s_preADD.lPredecessorIndex = mp_oControl.MathLib.GetTaskIndexByPosition(X, Y)
                oPredecessor = mp_oControl.Tasks.Item(s_preADD.lPredecessorIndex.ToString())
                If (oPredecessor.Key.Length = 0 Or oPredecessor.OutgoingPredecessors = False) Then
                    s_preADD.bCancel = True
                    Return
                End If
                s_preADD.sPredecessorKey = oPredecessor.Key
                If (X <= oPredecessor.Left + ((oPredecessor.Right - oPredecessor.Left) / 2)) Then
                    s_preADD.sPredecessorPosition = "S"
                Else
                    s_preADD.sPredecessorPosition = "E"
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseMove
                If s_preADD.bCancel = False Then
                    mp_oControl.clsG.EraseReversibleFrames()
                    mp_oControl.clsG.EraseReversibleLines()
                    'mp_DynamicRowMove(Y)
                    'mp_DynamicTimeLineMove(X)
                    mp_oControl.clsG.DrawReversibleLine(s_preADD.lXStart, s_preADD.lYStart, X, Y)
                    s_preADD.lTaskIndex = mp_oControl.MathLib.GetTaskIndexByPosition(X, Y)
                    If s_preADD.lTaskIndex > 0 Then
                        oTask = DirectCast(mp_oControl.Tasks.Item(s_preADD.lTaskIndex.ToString()), clsTask)
                        If oTask.IncomingPredecessors = False Then
                            mp_SetCursor(E_CURSORTYPE.CT_NODROP)
                            s_preADD.bCantAccept = True
                        Else
                            s_preADD.bCantAccept = False
                            mp_SetCursor(E_CURSORTYPE.CT_PREDECESSOR)
                            If (X <= oTask.Left + ((oTask.Right - oTask.Left) / 2)) Then
                                s_preADD.sTaskPosition = "S"
                            Else
                                s_preADD.sTaskPosition = "E"
                            End If
                            mp_oControl.ToolTipEventArgs.Clear()
                            mp_oControl.ToolTipEventArgs.PredecessorPosition = s_preADD.sPredecessorPosition
                            mp_oControl.ToolTipEventArgs.TaskPosition = s_preADD.sTaskPosition
                            mp_oControl.clsG.EraseReversibleLines()
                            mp_oControl.FireToolTipOnMouseMove(mp_yOperation)
                            mp_oControl.clsG.DrawReversibleLine(s_preADD.lXStart, s_preADD.lYStart, X, Y)
                            mp_DrawMovingReversibleFrame(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS)
                        End If
                    End If
                End If
            Case E_MOUSEKEYBOARDEVENTS.MouseUp
                Dim sType As String
                Dim mp_yType As E_CONSTRAINTTYPE = 0
                mp_oControl.clsG.EraseReversibleFrames()
                mp_oControl.clsG.EraseReversibleLines()
                If s_preADD.bCancel = False And s_preADD.bCantAccept = False Then
                    If s_preADD.lTaskIndex > 0 Then
                        oTask = DirectCast(mp_oControl.Tasks.Item(s_preADD.lTaskIndex.ToString()), clsTask)
                        If (X <= oTask.Left + ((oTask.Right - oTask.Left) / 2)) Then
                            s_preADD.sTaskPosition = "S"
                        Else
                            s_preADD.sTaskPosition = "E"
                        End If
                        sType = s_preADD.sPredecessorPosition & s_preADD.sTaskPosition
                        If sType = "EE" Then
                            mp_yType = E_CONSTRAINTTYPE.PCT_END_TO_END
                            mp_oControl.Predecessors.Add(oTask.Key, s_preADD.sPredecessorKey, E_CONSTRAINTTYPE.PCT_END_TO_END, "", "DS_PREDECESSOR")
                        ElseIf sType = "SS" Then
                            mp_yType = E_CONSTRAINTTYPE.PCT_START_TO_START
                            mp_oControl.Predecessors.Add(oTask.Key, s_preADD.sPredecessorKey, E_CONSTRAINTTYPE.PCT_START_TO_START, "", "DS_PREDECESSOR")
                        ElseIf sType = "ES" Then
                            mp_yType = E_CONSTRAINTTYPE.PCT_END_TO_START
                            mp_oControl.Predecessors.Add(oTask.Key, s_preADD.sPredecessorKey, E_CONSTRAINTTYPE.PCT_END_TO_START, "", "DS_PREDECESSOR")
                        ElseIf sType = "SE" Then
                            mp_yType = E_CONSTRAINTTYPE.PCT_START_TO_END
                            mp_oControl.Predecessors.Add(oTask.Key, s_preADD.sPredecessorKey, E_CONSTRAINTTYPE.PCT_START_TO_END, "", "DS_PREDECESSOR")
                        End If
                        If mp_oControl.EnforcePredecessors = True Then
                            mp_oControl.CheckPredecessors()
                        End If
                        oPredecessor = mp_oControl.Tasks.Item(s_preADD.sPredecessorKey)
                        mp_oControl.ObjectAddedEventArgs.Clear()
                        mp_oControl.ObjectAddedEventArgs.TaskIndex = oTask.Index
                        mp_oControl.ObjectAddedEventArgs.TaskKey = oTask.Key
                        mp_oControl.ObjectAddedEventArgs.PredecessorTaskIndex = oPredecessor.Index
                        mp_oControl.ObjectAddedEventArgs.PredecessorTaskKey = oPredecessor.Key
                        mp_oControl.ObjectAddedEventArgs.PredecessorObjectIndex = mp_oControl.Predecessors.Count
                        mp_oControl.ObjectAddedEventArgs.PredecessorType = mp_yType
                        mp_oControl.ObjectAddedEventArgs.EventTarget = E_EVENTTARGET.EVT_PREDECESSOR
                        mp_oControl.FireObjectAdded()
                    End If
                End If
                mp_oControl.Redraw()
                mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Case E_MOUSEKEYBOARDEVENTS.MouseClick
            Case E_MOUSEKEYBOARDEVENTS.MouseDblClick
            Case E_MOUSEKEYBOARDEVENTS.MouseWheel
            Case E_MOUSEKEYBOARDEVENTS.KeyDown
            Case E_MOUSEKEYBOARDEVENTS.KeyUp
            Case E_MOUSEKEYBOARDEVENTS.KeyPress
        End Select
    End Sub

    Private Function mp_bOverSplitter(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If mp_oControl.Splitter.Width = 0 Then
            Return False
        End If
        If X >= (mp_oControl.Splitter.Right - mp_oControl.Splitter.Width) And X <= mp_oControl.Splitter.Right And Y < mp_oControl.clsG.Height() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverEmptySpace(ByVal Y As Integer) As Boolean
        If Y > mp_oControl.Rows.TopOffset Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverTimeLine(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom And Y >= mp_oControl.mt_TopMargin Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn = Nothing
        If mp_oControl.SelectedColumnIndex = 0 Or mp_oControl.Columns.Count = 0 Then
            Return False
        End If
        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn = Nothing
        Dim lIndex As Integer
        If Not (X <= mp_oControl.Splitter.Left And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            If oColumn.Visible = True Then
                If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Friend Function mp_bOverSelectedRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        If mp_oControl.SelectedRowIndex = 0 Or mp_oControl.Rows.Count = 0 Then
            Return False
        End If
        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If oRow.MergeCells = True Then
            If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                Return True
            Else
                Return False
            End If
        Else
            If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                If mp_oControl.SelectedCellIndex = mp_oControl.MathLib.GetCellIndexByPosition(X) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End If
    End Function

    Friend Function mp_bOverRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        Dim lIndex As Integer
        If Not (X <= mp_oControl.CurrentViewObject.TimeLine.f_lStart And Y > mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Visible = True Then
                If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverSelectedTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedTask As clsTask = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedTaskIndex = 0 Then
            Return False
        End If
        oSelectedTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If X >= oSelectedTask.Left And X <= oSelectedTask.Right And Y >= oSelectedTask.Top And Y <= oSelectedTask.Bottom And mp_oControl.MathLib.InCurrentLayer(oSelectedTask.LayerIndex) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedPredecessor(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedPredecessor As clsPredecessor = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedPredecessorIndex = 0 Then
            Return False
        End If
        oSelectedPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(mp_oControl.SelectedPredecessorIndex), clsPredecessor)
        If oSelectedPredecessor.HitTest(X, Y) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedPercentage(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedPercentage As clsPercentage = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedPercentageIndex = 0 Then
            Return False
        End If
        oSelectedPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(mp_oControl.SelectedPercentageIndex), clsPercentage)
        If X >= oSelectedPercentage.Left And X <= oSelectedPercentage.RightSel And Y >= oSelectedPercentage.Top And Y <= oSelectedPercentage.Bottom Then
            Return True
        Else
            Return False
        End If
    End Function


    Private Function mp_yTaskArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedTask As clsTask = Nothing
        oSelectedTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If X >= oSelectedTask.Left And X <= oSelectedTask.Right And Y >= oSelectedTask.Top And Y <= oSelectedTask.Bottom And mp_oControl.MathLib.InCurrentLayer(oSelectedTask.LayerIndex) Then
            If X >= oSelectedTask.Left And X <= oSelectedTask.Left + 2 Then
                If oSelectedTask.f_bLeftVisible = True Then
                    Return E_AREA.EA_LEFT
                Else
                    Return E_AREA.EA_CENTER
                End If
            End If
            If X >= oSelectedTask.Right - 2 And X <= oSelectedTask.Right Then
                If oSelectedTask.f_bRightVisible = True Then
                    Return E_AREA.EA_RIGHT
                Else
                    Return E_AREA.EA_CENTER
                End If
            End If
            Return E_AREA.EA_CENTER
        End If
        Return E_AREA.EA_NONE
    End Function

    Friend Function mp_yRowArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedRow As clsRow = Nothing
        oSelectedRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If Y >= oSelectedRow.Bottom And Y <= oSelectedRow.Bottom + 3 Then
            Return E_AREA.EA_BOTTOM
        Else
            Return E_AREA.EA_CENTER
        End If
    End Function

    Private Function mp_yColumnArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedColumn As clsColumn = Nothing
        oSelectedColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If X >= (oSelectedColumn.Right - 3) And X <= oSelectedColumn.Right Then
            Return E_AREA.EA_RIGHT
        Else
            Return E_AREA.EA_CENTER
        End If
    End Function

    Private Sub mp_DynamicColumnMove(ByVal v_X As Integer)
        If v_X < mp_oControl.mt_LeftMargin Then
            If mp_oControl.HorizontalScrollBar.Value > 20 Then
                mp_oControl.HorizontalScrollBar.Value = mp_oControl.HorizontalScrollBar.Value - 20
            Else
                mp_oControl.HorizontalScrollBar.Value = 0
            End If
            mp_oControl.Redraw()
            Return
        End If
        If v_X > mp_oControl.Splitter.Left Then
            If mp_oControl.HorizontalScrollBar.Value < (mp_oControl.HorizontalScrollBar.Max - 20) Then
                mp_oControl.HorizontalScrollBar.Value = mp_oControl.HorizontalScrollBar.Value + 20
            Else
                mp_oControl.HorizontalScrollBar.Value = mp_oControl.HorizontalScrollBar.Max
            End If
            mp_oControl.Redraw()
            Return
        End If
    End Sub

    Friend Sub mp_DynamicRowMove(ByVal v_Y As Integer)
        If v_Y < mp_oControl.CurrentViewObject.TimeLine.Bottom Then
            If mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow > 1 Then
                mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow - 1
                mp_oControl.VerticalScrollBar.Value = mp_oControl.VerticalScrollBar.Value - 1
                mp_oControl.Redraw()
                Return
            End If
        End If
        If v_Y > mp_oControl.CurrentViewObject.ClientArea.Bottom Then
            If mp_oControl.VerticalScrollBar.Value < mp_oControl.VerticalScrollBar.Max Then
                mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow = mp_oControl.CurrentViewObject.ClientArea.FirstVisibleRow + 1
                mp_oControl.VerticalScrollBar.Value = mp_oControl.VerticalScrollBar.Value + 1
                mp_oControl.Redraw()
            End If
        End If
    End Sub

    Private Sub mp_DynamicTimeLineMove(ByVal v_X As Integer)
        If v_X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Enabled = False Then
                mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.f_ScrollInterval, mp_oControl.CurrentViewObject.Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
            Else
                If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value < mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Max Then
                    mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value = mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value + 1
                End If
            End If
            mp_oControl.Redraw()
        End If
        If v_X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Enabled = False Then
                mp_oControl.CurrentViewObject.TimeLine.f_StartDate = mp_oControl.MathLib.DateTimeAdd(mp_oControl.CurrentViewObject.f_ScrollInterval, -mp_oControl.CurrentViewObject.Factor, mp_oControl.CurrentViewObject.TimeLine.StartDate)
            Else
                If mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value > 0 Then
                    mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value = mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Value - 1
                End If
            End If
            mp_oControl.Redraw()
        End If
    End Sub

    Private Function mp_bOverTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oTask As clsTask = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Tasks.Count To 1 Step -1
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            If oTask.Visible = True And mp_oControl.MathLib.InCurrentLayer(oTask.LayerIndex) Then
                If X >= oTask.Left And X <= oTask.Right And Y >= oTask.Top And Y <= oTask.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverPredecessor(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oPredecessor As clsPredecessor = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Predecessors.Count To 1 Step -1
            oPredecessor = (DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor))
            If oPredecessor.Visible = True Then
                If oPredecessor.HitTest(X, Y) = True Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Function mp_bOverPercentage(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oPercentage As clsPercentage = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Percentages.Count To 1 Step -1
            oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            If oPercentage.Visible = True Then
                If X >= oPercentage.Left And X <= oPercentage.RightSel And Y >= oPercentage.Top And Y <= oPercentage.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverClientArea(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y >= mp_oControl.CurrentViewObject.ClientArea.Top Then
            Return True
        Else
            Return False
        End If
    End Function


    Private Function mp_fSnapX(ByVal X As Integer) As Integer
        Dim dtDate As AGVBE.DateTime = New AGVBE.DateTime()
        If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = False Then
            Return X
        End If
        dtDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
        dtDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, dtDate)
        Return mp_oControl.MathLib.GetXCoordinateFromDate(dtDate)
    End Function

    Friend Sub mp_DrawMovingReversibleFrame(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yFocusType As E_FOCUSTYPE)
        mp_oControl.clsG.f_FocusLeft = v_X1
        mp_oControl.clsG.f_FocusTop = v_Y1
        mp_oControl.clsG.f_FocusRight = v_X2
        mp_oControl.clsG.f_FocusBottom = v_Y2
        Select Case v_yFocusType
            Case E_FOCUSTYPE.FCT_NORMAL
            Case E_FOCUSTYPE.FCT_KEEPLEFTRIGHTBOUNDS
                If mp_oControl.clsG.f_FocusLeft < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
                    mp_oControl.clsG.f_FocusLeft = mp_oControl.CurrentViewObject.TimeLine.f_lStart
                End If
                If mp_oControl.clsG.f_FocusRight > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
                    mp_oControl.clsG.f_FocusRight = mp_oControl.CurrentViewObject.TimeLine.f_lEnd
                End If
            Case E_FOCUSTYPE.FCT_ADD
                If mp_oControl.clsG.f_FocusLeft < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
                    mp_oControl.clsG.f_FocusLeft = mp_oControl.CurrentViewObject.TimeLine.f_lStart
                End If
                If mp_oControl.clsG.f_FocusRight < mp_oControl.clsG.f_FocusLeft Then
                    mp_oControl.clsG.f_FocusRight = mp_oControl.clsG.f_FocusLeft
                    mp_oControl.clsG.f_FocusLeft = v_X2
                End If

            Case E_FOCUSTYPE.FCT_VERTICALSPLITTER
                If mp_oControl.clsG.f_FocusLeft >= mp_oControl.Splitter.Right Then
                    mp_oControl.clsG.f_FocusBottom = mp_oControl.CurrentViewObject.ClientArea.Bottom
                Else
                    mp_oControl.clsG.f_FocusBottom = mp_oControl.mt_TableBottom
                End If
        End Select
        mp_oControl.clsG.DrawReversibleFrameEx()
    End Sub

    Public Sub OnMouseLeave()
        '       mp_oToolTip.Visible = False
    End Sub

    Friend Function mp_bCursorEditTextColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn
        oColumn = mp_oControl.Columns.Item(mp_oControl.SelectedColumnIndex.ToString())
        If oColumn.AllowTextEdit = True Then
            If X >= oColumn.mp_lTextLeft And X <= oColumn.mp_lTextRight Then
                If Y >= oColumn.mp_lTextTop And Y <= oColumn.mp_lTextBottom Then
                    mp_SetCursor(E_CURSORTYPE.CT_IBEAM)
                    Return True
                End If
            End If
        End If
        mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        Return False
    End Function

    Friend Function mp_bShowEditTextColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn
        oColumn = mp_oControl.Columns.Item(mp_oControl.SelectedColumnIndex.ToString())
        If oColumn.AllowTextEdit = True Then
            If X >= oColumn.mp_lTextLeft And X <= oColumn.mp_lTextRight Then
                If Y >= oColumn.mp_lTextTop And Y <= oColumn.mp_lTextBottom Then
                    mp_oControl.mp_oTextBox.Initialize(mp_oControl.SelectedColumnIndex, 0, E_TEXTOBJECTTYPE.TOT_COLUMN, X, Y)
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Friend Function mp_bCursorEditTextRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow
        oRow = mp_oControl.Rows.Item(mp_oControl.SelectedRowIndex.ToString())
        If oRow.MergeCells = True Then
            If oRow.AllowTextEdit = True Then
                If X >= oRow.mp_lTextLeft And X <= oRow.mp_lTextRight Then
                    If Y >= oRow.mp_lTextTop And Y <= oRow.mp_lTextBottom Then
                        mp_SetCursor(E_CURSORTYPE.CT_IBEAM)
                        Return True
                    End If
                End If
            End If
        Else
            Dim oCell As clsCell
            Dim lCellIndex As Integer
            Dim oColumn As clsColumn
            For lCellIndex = 1 To mp_oControl.Columns.Count
                oColumn = mp_oControl.Columns.Item(lCellIndex.ToString())
                If oColumn.Visible = True Then
                    oCell = oRow.Cells.Item(lCellIndex.ToString())
                    If oCell.AllowTextEdit = True Then
                        If X >= oCell.mp_lTextLeft And X <= oCell.mp_lTextRight Then
                            If Y >= oCell.mp_lTextTop And Y <= oCell.mp_lTextBottom Then
                                mp_SetCursor(E_CURSORTYPE.CT_IBEAM)
                                Return True
                            End If
                        End If
                    End If
                End If
            Next
        End If
        mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        Return False
    End Function

    Friend Function mp_bShowEditTextRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow
        oRow = mp_oControl.Rows.Item(mp_oControl.SelectedRowIndex.ToString())
        If oRow.MergeCells = True Then
            If oRow.AllowTextEdit = True Then
                If X >= oRow.mp_lTextLeft And X <= oRow.mp_lTextRight Then
                    If Y >= oRow.mp_lTextTop And Y <= oRow.mp_lTextBottom Then
                        mp_oControl.mp_oTextBox.Initialize(mp_oControl.SelectedRowIndex, 0, E_TEXTOBJECTTYPE.TOT_ROW, X, Y)
                        Return True
                    End If
                End If
            End If
        Else
            Dim oCell As clsCell
            Dim lCellIndex As Integer
            Dim oColumn As clsColumn
            For lCellIndex = 1 To mp_oControl.Columns.Count
                oColumn = mp_oControl.Columns.Item(lCellIndex.ToString())
                If oColumn.Visible = True Then
                    oCell = oRow.Cells.Item(lCellIndex.ToString())
                    If oCell.AllowTextEdit = True Then
                        If X >= oCell.mp_lTextLeft And X <= oCell.mp_lTextRight Then
                            If Y >= oCell.mp_lTextTop And Y <= oCell.mp_lTextBottom Then
                                mp_oControl.mp_oTextBox.Initialize(mp_oControl.SelectedRowIndex, lCellIndex, E_TEXTOBJECTTYPE.TOT_CELL, X, Y)
                                Return True
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return False
    End Function

    Friend Function mp_bCursorEditTextTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oTask As clsTask
        If mp_oControl.SelectedTaskIndex <= 0 Then
            mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
            Return False
        End If
        oTask = mp_oControl.Tasks.Item(mp_oControl.SelectedTaskIndex.ToString())
        If oTask.AllowTextEdit = True Then
            If X >= oTask.mp_lTextLeft And X <= oTask.mp_lTextRight Then
                If Y >= oTask.mp_lTextTop And Y <= oTask.mp_lTextBottom Then
                    mp_SetCursor(E_CURSORTYPE.CT_IBEAM)
                    Return True
                End If
            End If
        End If
        mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        Return False
    End Function

    Friend Function mp_bShowEditTextTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oTask As clsTask
        If mp_oControl.SelectedTaskIndex <= 0 Then
            Return False
        End If
        oTask = mp_oControl.Tasks.Item(mp_oControl.SelectedTaskIndex.ToString())
        If oTask.AllowTextEdit = True Then
            If X >= oTask.mp_lTextLeft And X <= oTask.mp_lTextRight Then
                If Y >= oTask.mp_lTextTop And Y <= oTask.mp_lTextBottom Then
                    mp_oControl.mp_oTextBox.Initialize(mp_oControl.SelectedTaskIndex, 0, E_TEXTOBJECTTYPE.TOT_TASK, X, Y)
                    Return True
                End If
            End If
        End If
        Return False
    End Function


    Friend Sub mp_SetCursor(ByVal v_iCursorType As E_CURSORTYPE)
        Select Case v_iCursorType
            Case E_CURSORTYPE.CT_NORMAL
                mp_CustomCursor("")
                mp_oControl.oCanvas.Cursor = System.Windows.Input.Cursors.Arrow
            Case E_CURSORTYPE.CT_SIZETASK
                mp_CustomCursor("")
                mp_oControl.oCanvas.Cursor = System.Windows.Input.Cursors.SizeWE
            Case E_CURSORTYPE.CT_MOVETASK
                mp_CustomCursor("SIZE_ALL.PNG")
            Case E_CURSORTYPE.CT_MOVEMILESTONE
                mp_CustomCursor("SIZE_ALL.PNG")
            Case E_CURSORTYPE.CT_CLIENTAREA
                mp_CustomCursor("CROSS.PNG")
            Case E_CURSORTYPE.CT_MOVESPLITTER
                mp_CustomCursor("")
                mp_oControl.oCanvas.Cursor = System.Windows.Input.Cursors.SizeWE
            Case E_CURSORTYPE.CT_IBEAM
                mp_CustomCursor("")
                mp_oControl.oCanvas.Cursor = System.Windows.Input.Cursors.IBeam
            Case E_CURSORTYPE.CT_ROWHEIGHT
                mp_CustomCursor("HO_SPLIT.PNG")
            Case E_CURSORTYPE.CT_COLUMNWIDTH
                mp_CustomCursor("VE_SPLIT.PNG")
            Case E_CURSORTYPE.CT_MOVEROW
                mp_CustomCursor("DRAGMOVE.PNG")
            Case E_CURSORTYPE.CT_MOVECOLUMN
                mp_CustomCursor("DRAGMOVE.PNG")
            Case E_CURSORTYPE.CT_SCROLLTIMELINE
                mp_CustomCursor("C_WAIT05.PNG")
            Case E_CURSORTYPE.CT_PERCENTAGE
                mp_CustomCursor("PERCENTAGE.PNG")
            Case E_CURSORTYPE.CT_PREDECESSOR
                mp_CustomCursor("PREDECESSOR.PNG")
            Case E_CURSORTYPE.CT_NODROP
                mp_CustomCursor("NO_DROP.PNG")
        End Select
    End Sub

    Private Sub mp_CustomCursor(ByVal sImage As String)
        If sImage.Length = 0 Then
            mp_oControl.mp_oCursor.Source = Nothing
            mp_oControl.mp_sCursor = ""
        Else
            If (sImage <> mp_oControl.mp_sCursor) Or (mp_oControl.mp_sCursor = "") Then
                mp_oControl.mp_sCursor = sImage
                mp_oControl.oCanvas.Cursor = System.Windows.Input.Cursors.None
                Dim oURI As System.Uri = New System.Uri("AGVBE;component/" & sImage, UriKind.Relative)
                Dim oBitmap As New System.Windows.Media.Imaging.BitmapImage()
                Dim oSRI As System.Windows.Resources.StreamResourceInfo = Application.GetResourceStream(oURI)
                oBitmap.SetSource(oSRI.Stream)
                mp_oControl.mp_oCursor.Source = oBitmap
            End If
        End If
    End Sub

End Class
