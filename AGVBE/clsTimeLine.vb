Option Explicit On

Public Class clsTimeLine

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oView As clsView
    Public TimeLineScrollBar As clsTimeLineScrollBar
    Public TierArea As clsTierArea
    Public TickMarkArea As clsTickMarkArea
    Private mp_sStyleIndex As String
    Private mp_clrForeColor As Color
    Public ProgressLine As clsProgressLine
    Private mp_dtStartDate As AGVBE.DateTime
    Private mp_lEnd As Integer
    Private mp_lStart As Integer
    Private mp_oStyle As clsStyle


    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oView As clsView)
        mp_oControl = Value
        mp_oView = oView
        TimeLineScrollBar = New clsTimeLineScrollBar(mp_oControl)
        TierArea = New clsTierArea(mp_oControl, Me)
        TickMarkArea = New clsTickMarkArea(mp_oControl, Me, True)
        ProgressLine = New clsProgressLine(mp_oControl, Me)
        mp_sStyleIndex = "DS_TIMELINE"
        mp_oStyle = mp_oControl.Styles.FItem("DS_TIMELINE")
        mp_clrForeColor = Colors.Black
        Dim dtTimeNow As AGVBE.DateTime = New AGVBE.DateTime()
        dtTimeNow.SetToCurrentDateTime()
        f_StartDate = mp_oControl.MathLib.DateTimeAdd(E_INTERVAL.IL_HOUR, -3, dtTimeNow)
    End Sub

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_TIMELINE" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_TIMELINE"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    
    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    
    Public Property ForeColor() As Color
        Get
            Return mp_clrForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrForeColor = Value
        End Set
    End Property

    
    Public ReadOnly Property EndDate() As AGVBE.DateTime
        Get
            Calculate()
            Return mp_oControl.MathLib.DateTimeAdd(mp_oView.Interval, (mp_lEnd - mp_lStart) * mp_oView.Factor, mp_dtStartDate)
        End Get
    End Property

    
    Public ReadOnly Property StartDate() As AGVBE.DateTime
        Get
            Return mp_dtStartDate
        End Get
    End Property

    Friend WriteOnly Property f_StartDate() As AGVBE.DateTime
        Set(ByVal Value As AGVBE.DateTime)
            mp_dtStartDate = Value
            Calculate()
        End Set
    End Property

    Friend ReadOnly Property f_lStart() As Integer
        Get
            Return mp_lStart
        End Get
    End Property

    Friend ReadOnly Property f_lEnd() As Integer
        Get
            Return mp_lEnd
        End Get
    End Property

    
    Public Sub Move(ByVal Interval As E_INTERVAL, ByVal Factor As Integer)
        f_StartDate = mp_oControl.MathLib.DateTimeAdd(Interval, Factor, mp_dtStartDate)
    End Sub

    
    Public Sub Position(ByVal TimeLineStartDate As AGVBE.DateTime)
        f_StartDate = TimeLineStartDate
    End Sub

    
    Public ReadOnly Property Height() As Integer
        Get
            Return Bottom - Top
        End Get
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            Return mp_oControl.mt_BorderThickness
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            Dim lReturn As Integer
            Dim lUpperTierHeight As Integer
            Dim lMiddleTierHeight As Integer
            Dim lLowerTierHeight As Integer
            Dim lTickMarkAreaHeight As Integer
            lReturn = 0
            lUpperTierHeight = 0
            lLowerTierHeight = 0
            lTickMarkAreaHeight = 0
            If (TierArea.UpperTier.Visible = True) Then
                lUpperTierHeight = TierArea.UpperTier.Height
            End If
            If (TierArea.MiddleTier.Visible = True) Then
                lMiddleTierHeight = TierArea.MiddleTier.Height
            End If
            If (TierArea.LowerTier.Visible = True) Then
                lLowerTierHeight = TierArea.LowerTier.Height
            End If
            If (TickMarkArea.Visible = True) Then
                lTickMarkAreaHeight = TickMarkArea.Height
            End If
            lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight
            lReturn = lReturn + mp_oControl.mt_BorderThickness
            Return lReturn
        End Get
    End Property

    Friend Sub Calculate()
        If TimeLineScrollBar.Enabled = True Then
            mp_dtStartDate = mp_oControl.MathLib.DateTimeAdd(TimeLineScrollBar.Interval, TimeLineScrollBar.Value * TimeLineScrollBar.Factor, TimeLineScrollBar.StartDate)
        End If
        mp_dtStartDate = mp_oControl.MathLib.RoundDate(mp_oView.Interval, mp_oView.Factor, mp_dtStartDate)
        mp_lStart = mp_oControl.Splitter.Right()
        mp_lEnd = mp_oControl.mt_RightMargin()
        If mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED Then
            If mp_oStyle.ButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
                mp_lStart = mp_oControl.Splitter.Right() + 1
                mp_lEnd = mp_oControl.mt_RightMargin() - 1
            ElseIf mp_oStyle.ButtonStyle = GRE_BUTTONSTYLE.BT_LIGHTWEIGHT Then
                mp_lStart = mp_oControl.Splitter.Right() + 2
                mp_lEnd = mp_oControl.mt_RightMargin() - 2
            End If
        Else
            mp_lStart = mp_oControl.Splitter.Right()
            mp_lEnd = mp_oControl.mt_RightMargin()
        End If
    End Sub

    Friend Function TiersTickMarksPosition(ByVal v_yTier As String) As Integer
        Dim lReturn As Integer
        Dim lUpperTierHeight As Integer
        Dim lMiddleTierHeight As Integer
        Dim lLowerTierHeight As Integer
        Dim lTickMarkAreaHeight As Integer
        lReturn = 0
        lUpperTierHeight = 0
        lLowerTierHeight = 0
        lTickMarkAreaHeight = 0
        If (TierArea.UpperTier.Visible = True) Then
            lUpperTierHeight = TierArea.UpperTier.Height
        End If
        If (TierArea.MiddleTier.Visible = True) Then
            lMiddleTierHeight = TierArea.MiddleTier.Height
        End If
        If (TierArea.LowerTier.Visible = True) Then
            lLowerTierHeight = TierArea.LowerTier.Height
        End If
        If (TickMarkArea.Visible = True) Then
            lTickMarkAreaHeight = TickMarkArea.Height
        End If
        lReturn = lUpperTierHeight + lMiddleTierHeight + lLowerTierHeight + lTickMarkAreaHeight
        lReturn = lReturn + mp_oControl.mt_BorderThickness
        Select Case (v_yTier)
            Case "UpperTier"
                lReturn = lReturn - lUpperTierHeight - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight
            Case "MiddleTier"
                lReturn = lReturn - lMiddleTierHeight - lLowerTierHeight - lTickMarkAreaHeight
            Case "LowerTier"
                lReturn = lReturn - lLowerTierHeight - lTickMarkAreaHeight
            Case "TickMarkArea"
                lReturn = lReturn - lTickMarkAreaHeight
            Case Else
                'MsgBox("TiersTickMarksPosition Error")
        End Select
        TiersTickMarksPosition = lReturn
    End Function

    Friend Sub Draw()
        Dim lBottom As Integer
        Dim lTop As Integer
        Dim lLeft As Integer
        Dim lRight As Integer
        If (Height = 0) Then
            Return
        End If
        lBottom = Bottom
        lTop = Top
        lLeft = mp_oControl.Splitter.Right()
        lRight = mp_oControl.mt_RightMargin()
        mp_oControl.clsG.ClipRegion(lLeft, lTop, lRight, lBottom, True)
        mp_oControl.clsG.mp_DrawItem(lLeft, lTop, lRight, lBottom, "", "", False, Nothing, 0, 0, mp_oStyle)
        If mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_RAISED Then
            If mp_oStyle.ButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
                mp_oControl.clsG.ClipRegion(lLeft + 2, lTop + 2, lRight - 2, lBottom - 2, True)
            ElseIf mp_oStyle.ButtonStyle = GRE_BUTTONSTYLE.BT_LIGHTWEIGHT Then
                mp_oControl.clsG.ClipRegion(lLeft + 1, lTop + 1, lRight - 1, lBottom - 1, True)
            End If
        End If
        mp_oControl.clsG.ClipRegion(lLeft, lTop, lRight, lBottom, True)
        TierArea.UpperTier.Position()
        TierArea.MiddleTier.Position()
        TierArea.LowerTier.Position()
        TickMarkArea.Draw()
    End Sub

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TimeLine")
        oXML.InitializeWriter()
        oXML.WriteProperty("ForeColor", mp_clrForeColor)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("StartDate", mp_dtStartDate)
        oXML.WriteObject(ProgressLine.GetXML())
        oXML.WriteObject(TimeLineScrollBar.GetXML())
        oXML.WriteObject(TickMarkArea.GetXML())
        oXML.WriteObject(TierArea.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TimeLine")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("ForeColor", mp_clrForeColor)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        oXML.ReadProperty("StartDate", mp_dtStartDate)
        StyleIndex = mp_sStyleIndex
        Calculate()
        ProgressLine.SetXML(oXML.ReadObject("ProgressLine"))
        TimeLineScrollBar.SetXML(oXML.ReadObject("TimeLineScrollBar"))
        TickMarkArea.SetXML(oXML.ReadObject("TickMarkArea"))
        TierArea.SetXML(oXML.ReadObject("TierArea"))
    End Sub

End Class
