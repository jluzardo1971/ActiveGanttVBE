Option Explicit On 

Public Class clsTickMarkArea

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lHeight As Integer
    Private mp_lBigTickMarkHeight As Integer
    Private mp_lMediumTickMarkHeight As Integer
    Private mp_lSmallTickMarkHeight As Integer
    Private mp_bVisible As Boolean
    Private mp_lTextOffset As Integer
    Public TickMarks As clsTickMarks
    Private mp_oTimeLine As clsTimeLine
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTimeLine As clsTimeLine, ByVal bInit As Boolean)
        mp_oControl = Value
        mp_oTimeLine = oTimeLine
        mp_sStyleIndex = "DS_TICKMARKAREA"
        mp_oStyle = mp_oControl.Styles.FItem("DS_TICKMARKAREA")
        mp_lHeight = 23
        mp_lBigTickMarkHeight = 12
        mp_lMediumTickMarkHeight = 9
        mp_lSmallTickMarkHeight = 7
        mp_bVisible = True
        mp_lTextOffset = 11
        TickMarks = New clsTickMarks(mp_oControl)
    End Sub

    
    Public Property Height() As Integer
        Get
            Return mp_lHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lHeight = Value
        End Set
    End Property

    
    Public Property BigTickMarkHeight() As Integer
        Get
            Return mp_lBigTickMarkHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lBigTickMarkHeight = Value
        End Set
    End Property

    
    Public Property MediumTickMarkHeight() As Integer
        Get
            Return mp_lMediumTickMarkHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lMediumTickMarkHeight = Value
        End Set
    End Property

    
    Public Property SmallTickMarkHeight() As Integer
        Get
            Return mp_lSmallTickMarkHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lSmallTickMarkHeight = Value
        End Set
    End Property

    
    Public Property Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_TICKMARKAREA" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_TICKMARKAREA"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    
    Public Property TextOffset() As Integer
        Get
            Return mp_lTextOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lTextOffset = Value
        End Set
    End Property


    Friend Sub Draw()
        Dim dtBuff As AGVBE.DateTime = New AGVBE.DateTime()
        Dim oTickMark As clsTickMark = Nothing
        Dim lIndex As Integer
        If Visible = False Then
            Return
        End If
        mp_oControl.clsG.mp_DrawItem(mp_oTimeLine.f_lStart, mp_oTimeLine.Bottom - Height, mp_oTimeLine.f_lEnd, mp_oTimeLine.Bottom, "", "", False, Nothing, mp_oTimeLine.f_lStart, mp_oTimeLine.f_lEnd, mp_oStyle)
        mp_oControl.clsG.ClipRegion(mp_oTimeLine.f_lStart, mp_oTimeLine.Bottom - Height, mp_oTimeLine.f_lEnd, mp_oTimeLine.Bottom, False)
        For lIndex = 1 To TickMarks.Count
            Dim yInterval As E_INTERVAL
            Dim lFactor As Integer
            oTickMark = TickMarks.Item(lIndex.ToString())
            yInterval = oTickMark.Interval
            lFactor = oTickMark.Factor
            If mp_oControl.MathLib.GetXCoordinateFromDate(mp_oControl.MathLib.DateTimeAdd(yInterval, lFactor, mp_oTimeLine.StartDate)) - mp_oControl.MathLib.GetXCoordinateFromDate(mp_oTimeLine.StartDate) < 5 Then
                Exit For
            End If
            dtBuff = mp_oControl.MathLib.RoundDate(yInterval, lFactor, mp_oTimeLine.StartDate)
            If mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff) >= mp_oTimeLine.f_lStart Then
                PaintTickMark(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), oTickMark.TickMarkType)
                If oTickMark.DisplayText = True Then
                    PaintText(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), oTickMark.TextFormat)
                End If
            End If
            Do While dtBuff < mp_oTimeLine.EndDate
                dtBuff = mp_oControl.MathLib.DateTimeAdd(yInterval, lFactor, dtBuff)
                PaintTickMark(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), oTickMark.TickMarkType)
                If oTickMark.DisplayText = True Then
                    PaintText(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), oTickMark.TextFormat)
                End If
            Loop
        Next
        mp_oControl.clsG.ClearClipRegion()
    End Sub

    Private Sub PaintTickMark(ByVal fXCoordinate As Integer, ByVal TickMarkType As E_TICKMARKTYPES)
        Dim lTickMarkHeight As Integer
        Select Case TickMarkType
            Case E_TICKMARKTYPES.TLT_BIG
                lTickMarkHeight = mp_lBigTickMarkHeight
            Case E_TICKMARKTYPES.TLT_MEDIUM
                lTickMarkHeight = mp_lMediumTickMarkHeight
            Case E_TICKMARKTYPES.TLT_SMALL
                lTickMarkHeight = mp_lSmallTickMarkHeight
        End Select
        mp_oControl.clsG.DrawLine(fXCoordinate, mp_oTimeLine.Bottom - lTickMarkHeight, fXCoordinate, mp_oTimeLine.Bottom, GRE_LINETYPE.LT_NORMAL, mp_oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
    End Sub

    Private Sub PaintText(ByVal fXCoordinate As Integer, ByVal sFormat As String)
        Dim sDateBuff As String
        Dim lLeft As Integer
        Dim lTop As Integer
        Dim lRight As Integer
        Dim lBottom As Integer
        Dim lStringWidth As Integer
        Dim lStringHeight As Integer
        sDateBuff = mp_oControl.MathLib.GetDateFromXCoordinate(fXCoordinate).ToString(sFormat)
        lStringWidth = mp_oControl.mp_lStrWidth(sDateBuff, mp_oStyle.Font)
        lStringHeight = mp_oControl.mp_lStrHeight(sDateBuff, mp_oStyle.Font)
        lLeft = System.Convert.ToInt32(fXCoordinate - (lStringWidth / 2) - 10)
        lTop = mp_oTimeLine.Bottom - mp_lTextOffset - lStringHeight
        lRight = System.Convert.ToInt32(fXCoordinate + (lStringWidth / 2) + 10)
        lBottom = lTop + lStringHeight
        mp_oControl.clsG.DrawAlignedText(lLeft, lTop, lRight, lBottom, sDateBuff, GRE_HORIZONTALALIGNMENT.HAL_CENTER, GRE_VERTICALALIGNMENT.VAL_CENTER, mp_oStyle.ForeColor, mp_oStyle.Font)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TickMarkArea")
        oXML.InitializeWriter()
        oXML.WriteProperty("BigTickMarkHeight", mp_lBigTickMarkHeight)
        oXML.WriteProperty("Height", mp_lHeight)
        oXML.WriteProperty("MediumTickMarkHeight", mp_lMediumTickMarkHeight)
        oXML.WriteProperty("SmallTickMarkHeight", mp_lSmallTickMarkHeight)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("TextOffset", mp_lTextOffset)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteObject(TickMarks.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TickMarkArea")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("BigTickMarkHeight", mp_lBigTickMarkHeight)
        oXML.ReadProperty("Height", mp_lHeight)
        oXML.ReadProperty("MediumTickMarkHeight", mp_lMediumTickMarkHeight)
        oXML.ReadProperty("SmallTickMarkHeight", mp_lSmallTickMarkHeight)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("TextOffset", mp_lTextOffset)
        oXML.ReadProperty("Visible", mp_bVisible)
        TickMarks.SetXML(oXML.ReadObject("TickMarks"))
    End Sub

End Class
