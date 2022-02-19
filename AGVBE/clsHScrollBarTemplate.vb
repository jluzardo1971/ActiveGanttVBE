Option Explicit On
Imports System.Windows.Threading

Public Class clsHScrollBarTemplate

    Private Enum E_BUTTON
        BTN_NONE = 0
        BTN_LEFT = 1
        BTN_RIGHT = 2
        BTN_LEFTLCHANGE = 3
        BTN_RIGHTLCHANGE = 4
        BTN_BUTTON = 5
    End Enum

    Private mp_oControl As ActiveGanttVBECtl
    Friend mp_lSmallChange As Integer
    Friend mp_lLargeChange As Integer
    Friend mp_lValue As Integer
    Friend mp_lMin As Integer
    Friend mp_lMax As Integer
    Private ButtonX1 As Integer
    Private ButtonX2 As Integer
    Private ButtonY1 As Integer
    Private ButtonY2 As Integer
    Private Visible As Boolean
    Friend mp_bEnabled As Boolean
    Friend Height As Integer
    Friend Width As Integer
    Friend Top As Integer
    Friend Left As Integer
    Private mp_bMouseDown As Boolean
    Private mp_iMouseXPosition As Integer
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
        mp_lSmallChange = 1
        Height = 17
        mp_bMouseDown = False
        mp_iButton = E_BUTTON.BTN_NONE
        mp_iTimerInterval = 100
        mp_lValueBuff = Value
        AddHandler mp_oTimer.Tick, AddressOf mp_oTimer_Tick
        mp_sStyleIndex = "DS_SCROLLBAR"
        mp_oStyle = mp_oControl.Styles.FItem("DS_SCROLLBAR")
    End Sub

    Friend Property Enabled() As Boolean
        Get
            Return mp_bEnabled
        End Get
        Set(ByVal Value As Boolean)
            mp_bEnabled = Value
        End Set
    End Property

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
        Dim oArrowButtonLeftStyle As clsStyle = ArrowButtons.NormalStyle
        Dim oArrowButtonRightStyle As clsStyle = ArrowButtons.NormalStyle
        Dim oThumbButtonStyle As clsStyle = ThumbButton.NormalStyle
        If (mp_lMax - mp_lMin) = 0 Then
            Return
        End If
        mp_oControl.clsG.mp_DrawItem(Left, Top, Left + Width - 1, Top + Height - 1, "", "", False, Nothing, 0, 0, mp_oStyle)
        CalculateH()
        If Enabled = False Then
            oThumbButtonStyle = ThumbButton.DisabledStyle
            oArrowButtonLeftStyle = ArrowButtons.DisabledStyle
            oArrowButtonRightStyle = ArrowButtons.DisabledStyle
        ElseIf mp_bMouseDown = True Then
            If mp_iButton = E_BUTTON.BTN_LEFT Then
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonLeftStyle = ArrowButtons.PressedStyle
                oArrowButtonRightStyle = ArrowButtons.NormalStyle
            ElseIf mp_iButton = E_BUTTON.BTN_RIGHT Then
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonLeftStyle = ArrowButtons.NormalStyle
                oArrowButtonRightStyle = ArrowButtons.PressedStyle
            ElseIf mp_iButton = E_BUTTON.BTN_BUTTON Then
                oThumbButtonStyle = ThumbButton.PressedStyle
                oArrowButtonLeftStyle = ArrowButtons.NormalStyle
                oArrowButtonRightStyle = ArrowButtons.NormalStyle
            Else
                oThumbButtonStyle = ThumbButton.NormalStyle
                oArrowButtonLeftStyle = ArrowButtons.NormalStyle
                oArrowButtonRightStyle = ArrowButtons.NormalStyle
            End If
        Else
            oArrowButtonLeftStyle = ArrowButtons.NormalStyle
            oArrowButtonRightStyle = ArrowButtons.NormalStyle
            oThumbButtonStyle = ThumbButton.NormalStyle
        End If
        mp_oControl.clsG.mp_DrawItem(Left, Top, Left + 16, Top + Height - 1, "", "", False, Nothing, 0, 0, oArrowButtonLeftStyle)
        mp_oControl.clsG.mp_DrawItem(Left + Width - 17, Top, Left + Width - 1, Top + Height - 1, "", "", False, Nothing, 0, 0, oArrowButtonRightStyle)
        mp_oControl.clsG.mp_DrawItem(Left + ButtonX1, Top + ButtonY1 - 1, Left + ButtonX2 - 2, Top + ButtonY2 - 1, "", "", False, Nothing, 0, 0, oThumbButtonStyle)

        If oArrowButtonLeftStyle.ScrollBarStyle.DropShadow = True Then
            mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonLeftStyle.ScrollBarStyle.DropShadowLeftX, Top + oArrowButtonLeftStyle.ScrollBarStyle.DropShadowLeftY, GRE_ARROWDIRECTION.AWD_LEFT, oArrowButtonLeftStyle.ScrollBarStyle.ArrowSize, oArrowButtonLeftStyle.ScrollBarStyle.DropShadowArrowColor)
        End If
        mp_oControl.clsG.mp_DrawArrow(Left + oArrowButtonLeftStyle.ScrollBarStyle.LeftX, Top + oArrowButtonLeftStyle.ScrollBarStyle.LeftY, GRE_ARROWDIRECTION.AWD_LEFT, oArrowButtonLeftStyle.ScrollBarStyle.ArrowSize, oArrowButtonLeftStyle.ScrollBarStyle.ArrowColor)
        If oArrowButtonRightStyle.ScrollBarStyle.DropShadow = True Then
            mp_oControl.clsG.mp_DrawArrow(Left + Width - 17 + oArrowButtonRightStyle.ScrollBarStyle.DropShadowRightX, Top + oArrowButtonRightStyle.ScrollBarStyle.DropShadowRightY, GRE_ARROWDIRECTION.AWD_RIGHT, oArrowButtonRightStyle.ScrollBarStyle.ArrowSize, oArrowButtonRightStyle.ScrollBarStyle.DropShadowArrowColor)
        End If
        mp_oControl.clsG.mp_DrawArrow(Left + Width - 17 + oArrowButtonRightStyle.ScrollBarStyle.RightX, Top + oArrowButtonRightStyle.ScrollBarStyle.RightY, GRE_ARROWDIRECTION.AWD_RIGHT, oArrowButtonRightStyle.ScrollBarStyle.ArrowSize, oArrowButtonRightStyle.ScrollBarStyle.ArrowColor)
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
        CalculateH()
        If X < 17 Then
            mp_iButton = E_BUTTON.BTN_LEFT
            SmallChangeLeft()
            Return True
        ElseIf X > 17 And X < ButtonX1 Then
            mp_iButton = E_BUTTON.BTN_LEFTLCHANGE
            LargeChangeLeft()
            Return True
        ElseIf X >= ButtonX1 And X <= ButtonX2 Then
            mp_iButton = E_BUTTON.BTN_BUTTON
            mp_iMouseXPosition = X
            Return True
        ElseIf X > ButtonY2 And X < Width - 17 Then
            mp_iButton = E_BUTTON.BTN_RIGHTLCHANGE
            LargeChangeRight()
            Return True
        ElseIf X > Width - 17 And X < Width Then
            mp_iButton = E_BUTTON.BTN_RIGHT
            SmallChangeRight()
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub SmallChangeLeft()
        If (mp_lValue - mp_lSmallChange) >= mp_lMin Then
            mp_lValue = mp_lValue - mp_lSmallChange
        Else
            mp_lValue = mp_lMin
        End If
        mp_ValueChanged()
    End Sub

    Private Sub SmallChangeRight()
        If (mp_lValue + mp_lSmallChange) <= mp_lMax Then
            mp_lValue = mp_lValue + mp_lSmallChange
        Else
            mp_lValue = mp_lMax
        End If
        mp_ValueChanged()
    End Sub

    Private Sub LargeChangeLeft()
        If (mp_lValue - mp_lLargeChange) >= mp_lMin Then
            mp_lValue = mp_lValue - mp_lLargeChange
        ElseIf (mp_lValue - mp_lLargeChange) < mp_lMin Then
            mp_lValue = mp_lMin
        End If
        mp_ValueChanged()
    End Sub

    Private Sub LargeChangeRight()
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

    Private Sub CalculateH()
        Dim lWidth As Integer
        Dim lItemLength As Integer = 0
        lWidth = Width - (16 * 2) - 1
        If mp_lLargeChange > 0 Then
            lItemLength = System.Convert.ToInt32(lWidth / (((mp_lMax - mp_lMin) / mp_lLargeChange) + 1))
        End If
        If lItemLength < 8 Then
            lItemLength = 8
        End If
        lItemLength = lItemLength + 1
        lWidth = lWidth - lItemLength
        ButtonX1 = System.Convert.ToInt32((((mp_lValue - mp_lMin) / (mp_lMax - mp_lMin)) * lWidth) + 17)
        ButtonX2 = ButtonX1 + lItemLength
        ButtonY1 = 1
        ButtonY2 = ButtonY1 + Height - 1
    End Sub

    Private Function InvCalculateH(ByVal X As Integer) As Integer
        Dim lWidth As Integer
        Dim lItemLength As Integer = 0
        Dim iReturnValue As Integer
        lWidth = Width - (16 * 2) - 1
        If mp_lLargeChange > 0 Then
            lItemLength = System.Convert.ToInt32(lWidth / (((mp_lMax - mp_lMin) / mp_lLargeChange) + 1))
        End If
        If lItemLength < 8 Then
            lItemLength = 8
        End If
        lItemLength = lItemLength + 1
        lWidth = lWidth - lItemLength
        iReturnValue = System.Convert.ToInt32((((X - 17) * (mp_lMax - mp_lMin)) / lWidth) + mp_lMin)
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

    Friend Sub mp_oTimer_Tick()
        Select Case mp_iButton
            Case E_BUTTON.BTN_LEFT
                SmallChangeLeft()
            Case E_BUTTON.BTN_LEFTLCHANGE
                LargeChangeLeft()
            Case E_BUTTON.BTN_RIGHTLCHANGE
                LargeChangeRight()
            Case E_BUTTON.BTN_RIGHT
                SmallChangeRight()
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
        If mp_iButton = E_BUTTON.BTN_BUTTON Then
            Dim mp_iProjectedValue As Integer = mp_lValue + (InvCalculateH(X) - InvCalculateH(mp_iMouseXPosition))
            If mp_iProjectedValue >= mp_lMin And mp_iProjectedValue <= mp_lMax Then
                mp_iMouseXPosition = X
                mp_lValue = mp_iProjectedValue
                mp_ValueChanged()
            ElseIf mp_iProjectedValue < mp_lMin Then
                mp_iMouseXPosition = X
                mp_lValue = mp_lMin
                mp_ValueChanged()
            ElseIf mp_iProjectedValue > mp_lMax Then
                mp_iMouseXPosition = X
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
