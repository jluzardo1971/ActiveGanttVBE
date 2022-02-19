Option Explicit On

Public Class clsTier

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bVisible As Boolean
    Private mp_lFactor As Integer
    Private mp_lHeight As Integer
    Private mp_yInterval As E_INTERVAL
    Private mp_sTag As String
    Private mp_yTierType As E_TIERTYPE
    Private mp_yTierPosition As E_TIERPOSITION
    Private mp_sTierPosition As String
    Private mp_oTierArea As clsTierArea
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_yBackgroundMode As E_TIERBACKGROUNDMODE

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTierArea As clsTierArea, ByVal yTierPosition As E_TIERPOSITION)
        mp_oControl = Value
        mp_oTierArea = oTierArea
        mp_yTierPosition = yTierPosition
        Select Case mp_yTierPosition
            Case E_TIERPOSITION.SP_UPPER
                mp_sTierPosition = "UpperTier"
            Case E_TIERPOSITION.SP_MIDDLE
                mp_sTierPosition = "MiddleTier"
            Case E_TIERPOSITION.SP_LOWER
                mp_sTierPosition = "LowerTier"
        End Select
        mp_bVisible = True
        mp_yInterval = E_INTERVAL.IL_WEEK
        mp_lFactor = 1
        mp_lHeight = 14
        mp_sTag = ""
        mp_yTierType = E_TIERTYPE.ST_DAYOFWEEK
        mp_sStyleIndex = "DS_TIER"
        mp_oStyle = mp_oControl.Styles.FItem("DS_TIER")
        mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_TIERAPPEARANCE
    End Sub

    
    Public Property Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    
    Public Property Interval() As E_INTERVAL
        Get
            Return mp_yInterval
        End Get
        Set(ByVal Value As E_INTERVAL)
            mp_yInterval = Value
            mp_yTierType = E_TIERTYPE.ST_CUSTOM
        End Set
    End Property

    Public Property Factor() As Integer
        Get
            Return mp_lFactor
        End Get
        Set(ByVal value As Integer)
            mp_lFactor = value
        End Set
    End Property

    
    Public Property TierType() As E_TIERTYPE
        Get
            Return mp_yTierType
        End Get
        Set(ByVal Value As E_TIERTYPE)
            mp_yTierType = Value
            Select Case mp_yTierType
                Case E_TIERTYPE.ST_YEAR
                    mp_yInterval = E_INTERVAL.IL_YEAR
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_QUARTER
                    mp_yInterval = E_INTERVAL.IL_QUARTER
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_MONTH
                    mp_yInterval = E_INTERVAL.IL_MONTH
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_WEEK
                    mp_yInterval = E_INTERVAL.IL_WEEK
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_DAYOFWEEK
                    mp_yInterval = E_INTERVAL.IL_DAY
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_DAY
                    mp_yInterval = E_INTERVAL.IL_DAY
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_DAYOFYEAR
                    mp_yInterval = E_INTERVAL.IL_DAY
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_HOUR
                    mp_yInterval = E_INTERVAL.IL_HOUR
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_MINUTE
                    mp_yInterval = E_INTERVAL.IL_MINUTE
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_SECOND
                    mp_yInterval = E_INTERVAL.IL_SECOND
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_MILLISECOND
                    mp_yInterval = E_INTERVAL.IL_MILLISECOND
                    mp_lFactor = 1
                Case E_TIERTYPE.ST_MICROSECOND
                    mp_yInterval = E_INTERVAL.IL_MICROSECOND
                    mp_lFactor = 1
            End Select
        End Set
    End Property

    
    Public Property Height() As Integer
        Get
            Return mp_lHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lHeight = Value
        End Set
    End Property

    Friend Sub Position()
        Dim dtStart As AGVBE.DateTime
        Dim dtEnd As AGVBE.DateTime
        Dim lTop As Integer
        Dim lBottom As Integer
        Dim lTierHeight As Integer
        If (mp_bVisible = False) Then
            Return
        End If
        lTierHeight = Height
        lTop = mp_oTierArea.TimeLine.TiersTickMarksPosition(mp_sTierPosition)
        lBottom = lTop + lTierHeight
        If (mp_oControl.MathLib.GetXCoordinateFromDate(mp_oControl.MathLib.DateTimeAdd(mp_yInterval, mp_lFactor, mp_oTierArea.TimeLine.StartDate)) - mp_oControl.MathLib.GetXCoordinateFromDate(mp_oTierArea.TimeLine.StartDate) > 5) Then
            dtEnd = mp_oControl.MathLib.RoundDate(mp_yInterval, mp_lFactor, mp_oTierArea.TimeLine.StartDate)
            If (mp_oControl.MathLib.GetXCoordinateFromDate(dtEnd) >= mp_oTierArea.TimeLine.f_lStart) Then
                dtStart = mp_oControl.MathLib.DateTimeAdd(mp_yInterval, -mp_lFactor, dtEnd)
                dtStart = mp_oControl.MathLib.RoundDate(mp_yInterval, mp_lFactor, dtStart)
                Draw(dtStart, dtEnd, lTop, lBottom)
            End If
            Do While (dtEnd < mp_oTierArea.TimeLine.EndDate)
                dtStart = dtEnd
                dtEnd = mp_oControl.MathLib.DateTimeAdd(mp_yInterval, mp_lFactor, dtEnd)
                dtStart = mp_oControl.MathLib.RoundDate(mp_yInterval, mp_lFactor, dtStart)
                dtEnd = mp_oControl.MathLib.RoundDate(mp_yInterval, mp_lFactor, dtEnd)
                Draw(dtStart, dtEnd, lTop, lBottom)
            Loop
        End If
    End Sub

    Private Sub Draw(ByVal dtStart As AGVBE.DateTime, ByVal dtEnd As AGVBE.DateTime, ByVal lTop As Integer, ByVal lBottom As Integer)
        Dim lStart As Integer
        Dim lEnd As Integer
        Dim lStartTrim As Integer
        Dim lEndTrim As Integer

        Dim sText As String
        Dim lTextWidth As Integer

        Dim clrForeColor As Color
        Dim clrBackColor As Color
        Dim clrStartGradientColor As Color
        Dim clrEndGradientColor As Color
        Dim clrHatchBackColor As Color
        Dim clrHatchForeColor As Color

        lStart = mp_oControl.MathLib.GetXCoordinateFromDate(dtStart)
        lEnd = mp_oControl.MathLib.GetXCoordinateFromDate(dtEnd)
        If (lStart < mp_oTierArea.TimeLine.f_lStart) Then
            lStartTrim = mp_oTierArea.TimeLine.f_lStart
        Else
            lStartTrim = lStart
        End If
        If (lEnd > mp_oTierArea.TimeLine.f_lEnd) Then
            lEndTrim = mp_oTierArea.TimeLine.f_lEnd
        Else
            lEndTrim = lEnd
        End If
        sText = ""
        If (mp_yTierType = E_TIERTYPE.ST_CUSTOM) Then
            mp_oControl.CustomTierDrawEventArgs.Clear()
            mp_oControl.CustomTierDrawEventArgs.TierPosition = mp_yTierPosition
            mp_oControl.CustomTierDrawEventArgs.StartDate = dtStart
            mp_oControl.CustomTierDrawEventArgs.EndDate = dtEnd
            mp_oControl.CustomTierDrawEventArgs.Left = lStart
            mp_oControl.CustomTierDrawEventArgs.Top = lTop
            mp_oControl.CustomTierDrawEventArgs.Right = lEnd
            mp_oControl.CustomTierDrawEventArgs.Bottom = lBottom
            mp_oControl.CustomTierDrawEventArgs.LeftTrim = lStartTrim
            mp_oControl.CustomTierDrawEventArgs.RightTrim = lEndTrim
            'mp_oControl.CustomTierDrawEventArgs.Graphics = mp_oControl.clsG.oGraphics
            mp_oControl.CustomTierDrawEventArgs.Text = sText
            mp_oControl.CustomTierDrawEventArgs.StyleIndex = ""
            mp_oControl.CustomTierDrawEventArgs.Interval = Interval
            mp_oControl.CustomTierDrawEventArgs.Factor = Factor
            mp_oControl.FireCustomTierDraw()
            If (mp_oControl.CustomTierDrawEventArgs.StyleIndex <> "") Then
                mp_oControl.clsG.mp_DrawItem(lStart, lTop, lEnd, lBottom, mp_oControl.CustomTierDrawEventArgs.StyleIndex, mp_oControl.CustomTierDrawEventArgs.Text, False, Nothing, lStartTrim, lEndTrim, Nothing)
            End If
        Else
            If mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_TIERAPPEARANCE Then
                Dim oTierAppearance As clsTierAppearance = Nothing
                Dim oTierFormat As clsTierFormat = Nothing
                If (mp_oControl.TierAppearanceScope = E_TIERAPPEARANCESCOPE.TAS_CONTROL) Then
                    oTierAppearance = mp_oControl.TierAppearance
                ElseIf (mp_oControl.TierAppearanceScope = E_TIERAPPEARANCESCOPE.TAS_VIEW) Then
                    oTierAppearance = mp_oTierArea.TierAppearance
                End If
                If (mp_oControl.TierFormatScope = E_TIERFORMATSCOPE.TFS_CONTROL) Then
                    oTierFormat = mp_oControl.TierFormat
                ElseIf (mp_oControl.TierFormatScope = E_TIERFORMATSCOPE.TFS_VIEW) Then
                    oTierFormat = mp_oTierArea.TierFormat
                End If
                If (mp_yInterval = E_INTERVAL.IL_YEAR) Then
                    clrForeColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).ForeColor
                    clrBackColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).BackColor
                    clrStartGradientColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).StartGradientColor
                    clrEndGradientColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).EndGradientColor
                    clrHatchBackColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).HatchBackColor
                    clrHatchForeColor = oTierAppearance.YearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Year())).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.YearIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_QUARTER) Then
                    clrForeColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).ForeColor
                    clrBackColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).BackColor
                    clrStartGradientColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).StartGradientColor
                    clrEndGradientColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).EndGradientColor
                    clrHatchBackColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).HatchBackColor
                    clrHatchForeColor = oTierAppearance.QuarterColors.Item(dtStart.Quarter().ToString()).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.QuarterIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_MONTH) Then
                    clrForeColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).ForeColor
                    clrBackColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).BackColor
                    clrStartGradientColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).StartGradientColor
                    clrEndGradientColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).EndGradientColor
                    clrHatchBackColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).HatchBackColor
                    clrHatchForeColor = oTierAppearance.MonthColors.Item(dtStart.Month().ToString()).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.MonthIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_WEEK) Then
                    clrForeColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).ForeColor
                    clrBackColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).BackColor
                    clrStartGradientColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).StartGradientColor
                    clrEndGradientColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).EndGradientColor
                    clrHatchBackColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).HatchBackColor
                    clrHatchForeColor = oTierAppearance.WeekColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Week())).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.WeekIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_DAY) Then
                    If (mp_yTierType = E_TIERTYPE.ST_DAYOFWEEK) Then
                        clrBackColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).BackColor
                        clrForeColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).ForeColor
                        clrStartGradientColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).StartGradientColor
                        clrEndGradientColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).EndGradientColor
                        clrHatchBackColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).HatchBackColor
                        clrHatchForeColor = oTierAppearance.DayOfWeekColors.Item(dtStart.DayOfWeek().ToString()).HatchForeColor
                        sText = dtStart.ToString(oTierFormat.DayOfWeekIntervalFormat, mp_oControl.Culture)
                    ElseIf (mp_yTierType = E_TIERTYPE.ST_DAY) Then
                        clrBackColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).BackColor
                        clrForeColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).ForeColor
                        clrStartGradientColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).StartGradientColor
                        clrEndGradientColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).EndGradientColor
                        clrHatchBackColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).HatchBackColor
                        clrHatchForeColor = oTierAppearance.DayColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.Day())).HatchForeColor
                        sText = dtStart.ToString(oTierFormat.DayIntervalFormat, mp_oControl.Culture)
                    ElseIf (mp_yTierType = E_TIERTYPE.ST_DAYOFYEAR) Then
                        clrBackColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).BackColor
                        clrForeColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).ForeColor
                        clrStartGradientColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).StartGradientColor
                        clrEndGradientColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).EndGradientColor
                        clrHatchBackColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).HatchBackColor
                        clrHatchForeColor = oTierAppearance.DayOfYearColors.Item(mp_oControl.MathLib.GetTierIndex(dtStart.DayOfYear())).HatchForeColor
                        sText = dtStart.ToString(oTierFormat.DayOfYearIntervalFormat, mp_oControl.Culture)
                    End If
                ElseIf (mp_yInterval = E_INTERVAL.IL_HOUR) Then
                    clrBackColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).BackColor
                    clrForeColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).ForeColor
                    clrStartGradientColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).StartGradientColor
                    clrEndGradientColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).EndGradientColor
                    clrHatchBackColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).HatchBackColor
                    clrHatchForeColor = oTierAppearance.HourColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Hour()))).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.HourIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_MINUTE) Then
                    clrBackColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).BackColor
                    clrForeColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).ForeColor
                    clrStartGradientColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).StartGradientColor
                    clrEndGradientColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).EndGradientColor
                    clrHatchBackColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).HatchBackColor
                    clrHatchForeColor = oTierAppearance.MinuteColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Minute()))).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.MinuteIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_SECOND) Then
                    clrBackColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).BackColor
                    clrForeColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).ForeColor
                    clrStartGradientColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).StartGradientColor
                    clrEndGradientColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).EndGradientColor
                    clrHatchBackColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).HatchBackColor
                    clrHatchForeColor = oTierAppearance.SecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Second()))).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.SecondIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_MILLISECOND) Then
                    clrBackColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).BackColor
                    clrForeColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).ForeColor
                    clrStartGradientColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).StartGradientColor
                    clrEndGradientColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).EndGradientColor
                    clrHatchBackColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).HatchBackColor
                    clrHatchForeColor = oTierAppearance.MillisecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Millisecond()))).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.MillisecondIntervalFormat, mp_oControl.Culture)
                ElseIf (mp_yInterval = E_INTERVAL.IL_MICROSECOND) Then
                    clrBackColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).BackColor
                    clrForeColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).ForeColor
                    clrStartGradientColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).StartGradientColor
                    clrEndGradientColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).EndGradientColor
                    clrHatchBackColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).HatchBackColor
                    clrHatchForeColor = oTierAppearance.MicrosecondColors.Item((mp_oControl.MathLib.GetTierIndex(dtStart.Microsecond()))).HatchForeColor
                    sText = dtStart.ToString(oTierFormat.MicrosecondIntervalFormat, mp_oControl.Culture)
                End If
            End If
            If lEnd > mp_oTierArea.TimeLine.f_lEnd Then
                lEnd = mp_oTierArea.TimeLine.f_lEnd
            End If
            lTextWidth = mp_oControl.mp_lStrWidth(sText, mp_oStyle.Font)
            If (lEnd - lStart) > lTextWidth Then
                mp_oControl.CustomTierDrawEventArgs.Clear()
                mp_oControl.CustomTierDrawEventArgs.TierPosition = mp_yTierPosition
                mp_oControl.CustomTierDrawEventArgs.Text = sText
                mp_oControl.CustomTierDrawEventArgs.StartDate = dtStart
                mp_oControl.FireTierTextDraw()
                If mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_TIERAPPEARANCE Then
                    mp_oControl.clsG.mp_DrawItemEx(lStart, lTop, lEnd, lBottom, sText, False, Nothing, lStartTrim, lEndTrim, mp_oStyle, clrBackColor, clrForeColor, clrStartGradientColor, clrEndGradientColor, clrHatchBackColor, clrHatchForeColor)
                ElseIf mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_STYLE Then
                    mp_oControl.clsG.mp_DrawItem(lStart, lTop, lEnd, lBottom, sText, "", False, Nothing, lStartTrim, lEndTrim, mp_oStyle)
                End If
            Else
                If mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_TIERAPPEARANCE Then
                    mp_oControl.clsG.mp_DrawItemEx(lStart, lTop, lEnd, lBottom, "", False, Nothing, lStartTrim, lEndTrim, mp_oStyle, clrBackColor, clrForeColor, clrStartGradientColor, clrEndGradientColor, clrHatchBackColor, clrHatchForeColor)
                ElseIf mp_yBackgroundMode = E_TIERBACKGROUNDMODE.ETBGM_STYLE Then
                    mp_oControl.clsG.mp_DrawItem(lStart, lTop, lEnd, lBottom, "", "", False, Nothing, lStartTrim, lEndTrim, mp_oStyle)
                End If
            End If
        End If
    End Sub

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_TIER" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_TIER"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Property BackgroundMode() As E_TIERBACKGROUNDMODE
        Get
            Return mp_yBackgroundMode
        End Get
        Set(ByVal Value As E_TIERBACKGROUNDMODE)
            mp_yBackgroundMode = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As clsXML
        oXML = New clsXML(mp_oControl, mp_sTierPosition)
        oXML.InitializeWriter()
        oXML.WriteProperty("yTierPosition", mp_yTierPosition)
        oXML.WriteProperty("sTierPosition", mp_sTierPosition)
        oXML.WriteProperty("Height", mp_lHeight)
        oXML.WriteProperty("Interval", mp_yInterval)
        oXML.WriteProperty("Factor", mp_lFactor)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("TierType", mp_yTierType)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("BackgroundMode", mp_yBackgroundMode)
        Return oXML.GetXML
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As clsXML
        oXML = New clsXML(mp_oControl, mp_sTierPosition)
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("yTierPosition", mp_yTierPosition)
        oXML.ReadProperty("sTierPosition", mp_sTierPosition)
        oXML.ReadProperty("Height", mp_lHeight)
        oXML.ReadProperty("Interval", mp_yInterval)
        oXML.ReadProperty("Factor", mp_lFactor)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("TierType", mp_yTierType)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("BackgroundMode", mp_yBackgroundMode)
    End Sub

End Class

