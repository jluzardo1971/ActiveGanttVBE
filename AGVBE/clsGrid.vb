Option Explicit On

Public Class clsGrid

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bHorizontalLines As Boolean
    Private mp_bVerticalLines As Boolean
    Private mp_bSnapToGrid As Boolean
    Private mp_bSnapToGridOnSelection As Boolean
    Private mp_clrColor As Color
    Private mp_yInterval As E_INTERVAL
    Private mp_lFactor As Integer
    Private mp_oTimeLine As clsTimeLine

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTimeLine As clsTimeLine)
        mp_oControl = Value
        mp_oTimeLine = oTimeLine
        mp_bHorizontalLines = True
        mp_bVerticalLines = False
        mp_bSnapToGrid = False
        mp_bSnapToGridOnSelection = True
        mp_clrColor = Color.FromArgb(255, 192, 192, 192)
        mp_yInterval = E_INTERVAL.IL_MINUTE
        mp_lFactor = 15
    End Sub

    
    Public Property HorizontalLines() As Boolean
        Get
            Return mp_bHorizontalLines
        End Get
        Set(ByVal Value As Boolean)
            mp_bHorizontalLines = Value
        End Set
    End Property

    
    Public Property VerticalLines() As Boolean
        Get
            Return mp_bVerticalLines
        End Get
        Set(ByVal Value As Boolean)
            mp_bVerticalLines = Value
        End Set
    End Property

    
    Public Property SnapToGrid() As Boolean
        Get
            Return mp_bSnapToGrid
        End Get
        Set(ByVal Value As Boolean)
            mp_bSnapToGrid = Value
        End Set
    End Property

    
    Public Property SnapToGridOnSelection() As Boolean
        Get
            Return mp_bSnapToGridOnSelection
        End Get
        Set(ByVal Value As Boolean)
            mp_bSnapToGridOnSelection = Value
        End Set
    End Property

    
    Public Property Color() As Color
        Get
            Return mp_clrColor
        End Get
        Set(ByVal Value As Color)
            mp_clrColor = Value
        End Set
    End Property
    
    Public Property Interval() As E_INTERVAL
        Get
            Return mp_yInterval
        End Get
        Set(ByVal Value As E_INTERVAL)
            mp_yInterval = Value
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

    Friend Sub Draw()
        Dim dtBuff As AGVBE.DateTime
        If mp_bVerticalLines = False Then
            Return
        End If
        If mp_oControl.MathLib.GetXCoordinateFromDate(mp_oControl.MathLib.DateTimeAdd(mp_yInterval, mp_lFactor, mp_oTimeLine.StartDate)) - mp_oControl.MathLib.GetXCoordinateFromDate(mp_oTimeLine.StartDate) < 5 Then
            Return
        End If
        mp_oControl.clsG.ClipRegion(mp_oTimeLine.f_lStart, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oTimeLine.f_lEnd, mp_oControl.CurrentViewObject.ClientArea.Bottom, True)
        dtBuff = mp_oControl.MathLib.RoundDate(mp_yInterval, mp_lFactor, mp_oTimeLine.StartDate)
        If mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff) >= mp_oTimeLine.f_lStart Then
            mp_PaintVerticalGridLine(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), GRE_LINEDRAWSTYLE.LDS_SOLID)
        End If
        Do While dtBuff < mp_oTimeLine.EndDate
            dtBuff = mp_oControl.MathLib.DateTimeAdd(mp_yInterval, mp_lFactor, dtBuff)
            mp_PaintVerticalGridLine(mp_oControl.MathLib.GetXCoordinateFromDate(dtBuff), GRE_LINEDRAWSTYLE.LDS_SOLID)
        Loop
    End Sub

    Private Sub mp_PaintVerticalGridLine(ByVal fXCoordinate As Integer, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE)
        mp_oControl.clsG.DrawLine(fXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Top, fXCoordinate, mp_oControl.Rows.TopOffset, GRE_LINETYPE.LT_NORMAL, mp_clrColor, v_lDrawStyle)
    End Sub

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Grid")
        oXML.InitializeWriter()
        oXML.WriteProperty("Color", mp_clrColor)
        oXML.WriteProperty("HorizontalLines", mp_bHorizontalLines)
        oXML.WriteProperty("Interval", mp_yInterval)
        oXML.WriteProperty("Factor", mp_lFactor)
        oXML.WriteProperty("SnapToGrid", mp_bSnapToGrid)
        oXML.WriteProperty("SnapToGridOnSelection", mp_bSnapToGridOnSelection)
        oXML.WriteProperty("VerticalLines", mp_bVerticalLines)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Grid")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Color", mp_clrColor)
        oXML.ReadProperty("HorizontalLines", mp_bHorizontalLines)
        oXML.ReadProperty("Interval", mp_yInterval)
        oXML.ReadProperty("Factor", mp_lFactor)
        oXML.ReadProperty("SnapToGrid", mp_bSnapToGrid)
        oXML.ReadProperty("SnapToGridOnSelection", mp_bSnapToGridOnSelection)
        oXML.ReadProperty("VerticalLines", mp_bVerticalLines)
    End Sub

End Class
