Option Explicit On
Imports System.Windows.Threading

Public Class clsVScrollBarTemplate

    Private Enum E_BUTTON
        BTN_NONE = 0
        BTN_UP = 1
        BTN_DOWN = 2
        BTN_UPLCHANGE = 3
        BTN_DOWNLCHANGE = 4
        BTN_BUTTON = 5
    End Enum

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lSmallChange As Integer = 1
    Private mp_lLargeChange As Integer
    Private mp_lValue As Integer
    Private mp_lMin As Integer
    Private mp_lMax As Integer
    Private ButtonX1 As Integer
    Private ButtonX2 As Integer
    Private ButtonY1 As Integer
    Private ButtonY2 As Integer
    Private Visible As Boolean
    Private Enabled As Boolean
    Friend Height As Integer
    Friend Width As Integer
    Friend Top As Integer
    Friend Left As Integer
    Private mp_bMouseDown As Boolean
    Private mp_iMouseYPosition As Integer
    Private mp_iButton As E_BUTTON
    Private mp_iTimerInterval As Integer
    Private mp_yState As Integer
    Private mp_lValueBuff As Integer
    Private mp_oTimer As New DispatcherTimer
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Public ArrowButtons As clsButtonState
    Public ThumbButton As clsButtonState

    Friend Event ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs, ByVal Offset As Integer)

    Friend Sub Initialize(ByVal oControl As ActiveGanttVBECtl)
        mp_oControl = oControl
        ArrowButtons = New clsButtonState(oControl, "Arrow")
        ThumbButton = New clsButtonState(oControl, "Thumb")
        mp_lValueBuff = Value
        Width = 17
        mp_iButton = E_BUTTON.BTN_NONE
        mp_iTimerInterval = 100
        mp_lValueBuff = Value
        AddHandler mp_oTimer.Tick, AddressOf mp_oTimer_Tick
        mp_sStyleIndex = "DS_SCROLLBAR"
        mp_oStyle = mp_oControl.Styles.FItem("DS_SCROLLBAR")
    End Sub

    Friend Property Value() As Integer
        Get
            Return mp_lValue
        End Get
        Set(ByVal Value As Integer)
            mp_lValue = Value
            If mp_lValue < mp_lMin Then
                Throw New Exception("Value is less than mp_lMin.")
            End If
        End Set
    End Property

    Friend Property Min() As Integer
        Get
            Return mp_lMin
        End Get
        Set(ByVal Value As Integer)
            mp_lMin = Value
        End Set
    End Property

    Friend Property Max() As Integer
        Get
            Return mp_lMax
        End Get
        Set(ByVal Value As Integer)
            mp_lMax = Value
            If mp_lMax < mp_lMin Then
                Throw New Exception("mp_lMax is less than mp_lMin.")
            End If
        End Set
    End Property

    Friend Property SmallChange() As Integer
        Get
            Return mp_lSmallChange
        End Get
        Set(ByVal Value As Integer)
            mp_lSmallChange = Value
        End Set
    End Property

    Friend Property LargeChange() As Integer
        Get
            Return mp_lLargeChange
        End Get
        Set(ByVal Value As Integer)
            mp_lLargeChange = Value
        End Set
    End Property

    Friend Sub Draw()
        If Visible = False Then Return
        Dim oArrowButtonUpStyle As clsStyle = ArrowButtons.NormalStyle
        Dim oArrowButtonDownStyle As clsStyle = ArrowButtons.NormalStyle
        Dim oThumbButtonStyle As clsStyle = ThumbButton.NormalStyle
        If (mp_lMax - mp_lMin) = 0 Then
            Return
        End If
        mp_oControl.clsG.mp_DrawItem(Left, Top, Left + Width - 1, Top + Height - 1, "", "", False, Nothing, 0, 0, mp_oStyle)
        CalculateV()
        If Enabled = False Then
            oThumbButtonStyle = ThumbButton.DisabledStyle
            oArrowButtonUpStyle = ArrowButtons.DisabledStyle
            oArrowButtonDownStyle = ArrowButtons.DisabledStyle
        ElseIf mp_bMouseDown = True Then
            If mp_iButton = E_BUTTON.BTN_UP Then
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonUpStyle = ArrowButtons.PressedStyle
                oArrowButtonDownStyle = ArrowButtons.NormalStyle
            ElseIf mp_iButton = E_BUTTON.BTN_DOWN Then
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonUpStyle = ArrowButtons.NormalStyle
                oArrowButtonDownStyle = ArrowButtons.PressedStyle
            ElseIf mp_iButton = E_BUTTON.BTN_BUTTON Then
                oThumbButtonStyle = ThumbButton.PressedStyle
                oArrowButtonUpStyle = ArrowButtons.NormalStyle
                oArrowButtonDownStyle = ArrowButtons.NormalStyle
            Else
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonUpStyle = ArrowButtons.NormalStyle
                oArrowButtonDownStyle = ArrowButtons.NormalStyle
            End If
        Else
            oArrowButtonUpStyle = ArrowButtons.NormalStyle
            oArrowButtonDownStyle = ArrowButtons.NormalStyle
            oThumbButtonStyle = ThumbButton.NormalStyle
        End If
        mp_oControl.clsG.mp_DrawItem(Left, Top, Left + Width - 1, Top + 16, "", "", False, Nothing, 0, 0, oArrowButtonUpStyle)
        mp_oControl.clsG.mp_DrawItem(Left, Top + Height - 17, Left + Width - 1, Top + Height - 1, "", "", False, Nothing, 0, 0, oArrowButtonDownStyle)
        mp_oControl.clsG.mp_DrawItem(Left + ButtonX1 - 1, Top + ButtonY1, Left + ButtonX2 - 1, Top + ButtonY2 - 2, "", "", False, Nothing, 0, 0, oThumbButtonStyle)

        If oArrowButtonUpStyle.ScrollBarStyle.DropShadow = True Then
            mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonUpStyle.ScrollBarStyle.DropShadowUpX, Top + oArrowButtonUpStyle.ScrollBarStyle.DropShadowUpY, GRE_ARROWDIRECTION.AWD_UP, oArrowButtonUpStyle.ScrollBarStyle.ArrowSize, oArrowButtonUpStyle.ScrollBarStyle.DropShadowArrowColor)
        End If
        mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonUpStyle.ScrollBarStyle.UpX, Top + oArrowButtonUpStyle.ScrollBarStyle.UpY, GRE_ARROWDIRECTION.AWD_UP, oArrowButtonUpStyle.ScrollBarStyle.ArrowSize, oArrowButtonUpStyle.ScrollBarStyle.ArrowColor)
        If oArrowButtonDownStyle.ScrollBarStyle.DropShadow = True Then
            mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonDownStyle.ScrollBarStyle.DropShadowDownX, Top + Height - 17 + oArrowButtonDownStyle.ScrollBarStyle.DropShadowDownY, GRE_ARROWDIRECTION.AWD_DOWN, oArrowButtonDownStyle.ScrollBarStyle.ArrowSize, oArrowButtonDownStyle.ScrollBarStyle.DropShadowArrowColor)
        End If
        mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonDownStyle.ScrollBarStyle.DownX, Top + Height - 17 + oArrowButtonDownStyle.ScrollBarStyle.DownY, GRE_ARROWDIRECTION.AWD_DOWN, oArrowButtonDownStyle.ScrollBarStyle.ArrowSize, oArrowButtonDownStyle.ScrollBarStyle.ArrowColor)
    End Sub

    Friend Function OverControl(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If X >= Left And X <= Left + Width And Y >= Top And Y <= Top + Height Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ConvertToRelativeCoords(ByRef X As Integer, ByRef Y As Integer)
        X = X - Left
        Y = Y - Top
    End Sub

    Private Function ScrollBarClick(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If OverControl(X, Y) = False Then
            Return False
        End If
        ConvertToRelativeCoords(X, Y)
        CalculateV()
        If Y < 17 Then
            mp_iButton = E_BUTTON.BTN_UP
            SmallChangeUp()
            Return True
        ElseIf Y > 17 And Y < ButtonY1 Then
            mp_iButton = E_BUTTON.BTN_UPLCHANGE
            LargeChangeUp()
            Return True
        ElseIf Y >= ButtonY1 And Y <= ButtonY2 Then
            mp_iButton = E_BUTTON.BTN_BUTTON
            mp_iMouseYPosition = Y
            Return True
        ElseIf Y > ButtonY2 And Y < Height - 17 Then
            mp_iButton = E_BUTTON.BTN_DOWNLCHANGE
            LargeChangeDown()
            Return True
        ElseIf Y > Height - 17 And Y < Height Then
            mp_iButton = E_BUTTON.BTN_DOWN
            SmallChangeDown()
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub SmallChangeUp()
        If (mp_lValue - mp_lSmallChange) >= mp_lMin Then
            mp_lValue = mp_lValue - mp_lSmallChange
        Else
            mp_lValue = mp_lMin
        End If
        mp_ValueChanged()
    End Sub

    Private Sub SmallChangeDown()
        If (mp_lValue + mp_lSmallChange) <= mp_lMax Then
            mp_lValue = mp_lValue + mp_lSmallChange
        Else
            mp_lValue = mp_lMax
        End If
        mp_ValueChanged()
    End Sub

    Private Sub LargeChangeUp()
        If (mp_lValue - mp_lLargeChange) >= mp_lMin Then
            mp_lValue = mp_lValue - mp_lLargeChange
        ElseIf (mp_lValue - mp_lLargeChange) < mp_lMin Then
            mp_lValue = mp_lMin
        End If
        mp_ValueChanged()
    End Sub

    Private Sub LargeChangeDown()
        If (mp_lValue + mp_lLargeChange) <= mp_lMax Then
            mp_lValue = mp_lValue + mp_lLargeChange
        ElseIf (mp_lValue + mp_lLargeChange) > mp_lMax Then
            mp_lValue = mp_lMax
        End If
        mp_ValueChanged()
    End Sub

    Private Sub mp_ValueChanged()
        Dim e As New System.EventArgs()
        If (mp_lValue - mp_lValueBuff) <> 0 Then
            RaiseEvent ValueChanged(Me, e, mp_lValue - mp_lValueBuff)
        End If
        mp_lValueBuff = mp_lValue
    End Sub

    Private Sub CalculateV()
        Dim lHeight As Integer
        Dim lItemLength As Integer = 0
        lHeight = Height - (16 * 2) - 1
        If mp_lLargeChange > 0 Then
            lItemLength = System.Convert.ToInt32(lHeight / (((mp_lMax - mp_lMin) / mp_lLargeChange) + 1))
        End If
        If lItemLength < 8 Then
            lItemLength = 8
        End If
        lItemLength = lItemLength + 1
        lHeight = lHeight - lItemLength
        ButtonX1 = 1
        ButtonX2 = ButtonX1 + Width - 1
        ButtonY1 = System.Convert.ToInt32((((mp_lValue - mp_lMin) / (mp_lMax - mp_lMin)) * lHeight) + 17)
        ButtonY2 = ButtonY1 + lItemLength
    End Sub

    Private Function InvCalculateV(ByVal Y As Integer) As Integer
        Dim lHeight As Integer
        Dim lItemLength As Integer = 0
        Dim iReturnValue As Integer
        lHeight = Height - (16 * 2) - 1
        If mp_lLargeChange > 0 Then
            lItemLength = System.Convert.ToInt32(lHeight / (((mp_lMax - mp_lMin) / mp_lLargeChange) + 1))
        End If
        If lItemLength < 8 Then
            lItemLength = 8
        End If
        lItemLength = lItemLength + 1
        lHeight = lHeight - lItemLength
        iReturnValue = System.Convert.ToInt32((((Y - 17) * (mp_lMax - mp_lMin)) / lHeight) + mp_lMin)
        Return iReturnValue
    End Function

    Friend Property State() As Integer
        Get
            Return mp_yState
        End Get
        Set(ByVal Value As Integer)
            mp_yState = Value
            Select Case mp_yState
                Case E_SCROLLSTATE.SS_CANTDISPLAY
                    mp_yState = E_SCROLLSTATE.SS_HIDDEN
                    Me.Visible = False
                Case E_SCROLLSTATE.SS_NOTNEEDED
                    If mp_oControl.ScrollBarBehaviour = E_SCROLLBEHAVIOUR.SB_DISABLE Then
                        mp_yState = E_SCROLLSTATE.SS_SHOWN
                        Me.Enabled = False
                        Me.Visible = True
                    Else
                        mp_yState = E_SCROLLSTATE.SS_HIDDEN
                        Me.Visible = False
                    End If
                Case E_SCROLLSTATE.SS_NEEDED
                    mp_yState = E_SCROLLSTATE.SS_SHOWN
                    Me.Enabled = True
                    Me.Visible = True
            End Select
        End Set
    End Property

    Public Property TimerInterval() As Integer
        Get
            Return mp_iTimerInterval
        End Get
        Set(ByVal Value As Integer)
            mp_iTimerInterval = Value
        End Set
    End Property

    Private Sub mp_oTimer_Tick()
        Select Case mp_iButton
            Case E_BUTTON.BTN_UP
                SmallChangeUp()
            Case E_BUTTON.BTN_UPLCHANGE
                LargeChangeUp()
            Case E_BUTTON.BTN_DOWNLCHANGE
                LargeChangeDown()
            Case E_BUTTON.BTN_DOWN
                SmallChangeDown()
        End Select
    End Sub

    Friend Sub OnMouseHover(ByVal X As Integer, ByVal Y As Integer)
        mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
    End Sub

    Friend Sub OnMouseDown(ByVal X As Integer, ByVal Y As Integer)
        If Enabled = False Then Return
        If Visible = False Then Return
        mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        mp_bMouseDown = True
        mp_oTimer.Interval = New TimeSpan(0, 0, 0, 0, mp_iTimerInterval)
        mp_oTimer.Start()
        ScrollBarClick(X, Y)
    End Sub

    Friend Sub OnMouseMove(ByVal X As Integer, ByVal Y As Integer)
        If Enabled = False Then Return
        If Visible = False Then Return
        ConvertToRelativeCoords(X, Y)
        mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        If mp_iButton = E_BUTTON.BTN_BUTTON Then
            Dim mp_iProjectedValue As Integer = mp_lValue + (InvCalculateV(Y) - InvCalculateV(mp_iMouseYPosition))
            If mp_iProjectedValue >= mp_lMin And mp_iProjectedValue <= mp_lMax Then
                mp_iMouseYPosition = Y
                mp_lValue = mp_iProjectedValue
                mp_ValueChanged()
            ElseIf mp_iProjectedValue < mp_lMin Then
                mp_iMouseYPosition = Y
                mp_lValue = mp_lMin
                mp_ValueChanged()
            ElseIf mp_iProjectedValue > mp_lMax Then
                mp_iMouseYPosition = Y
                mp_lValue = mp_lMax
                mp_ValueChanged()
            End If
        End If
    End Sub

    Friend Sub OnMouseUp()
        If Enabled = False Then Return
        If Visible = False Then Return
        mp_oControl.MouseKeyboardEvents.mp_SetCursor(E_CURSORTYPE.CT_NORMAL)
        mp_bMouseDown = False
        mp_oTimer.Stop()
        mp_iButton = E_BUTTON.BTN_NONE
        mp_oControl.Redraw()
    End Sub

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_SCROLLBAR" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_SCROLLBAR"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ScrollBar")
        oXML.InitializeWriter()
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteObject(ArrowButtons.GetXML())
        oXML.WriteObject(ThumbButton.GetXML())
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ScrollBar")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        ArrowButtons.SetXML(oXML.ReadObject("ArrowButtonState"))
        ThumbButton.SetXML(oXML.ReadObject("ThumbButtonState"))
    End Sub

End Class
