Option Explicit On

Public Class clsProgressLine

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clrForeColor As Color
    Private mp_dtPosition As AGVBE.DateTime
    Private mp_yLength As E_PROGRESSLINELENGTH
    Private mp_yLineType As E_PROGRESSLINETYPE
    Private mp_oTimeLine As clsTimeLine

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTimeLine As clsTimeLine)
        mp_oControl = Value
        mp_clrForeColor = Colors.Red
        mp_dtPosition = New AGVBE.DateTime()
        mp_dtPosition.SetToCurrentDateTime()
        mp_yLength = E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA
        mp_yLineType = E_PROGRESSLINETYPE.TLMT_SYSTEMTIME
        mp_oTimeLine = oTimeLine
    End Sub

    
    Public Property Position() As AGVBE.DateTime
        Get
            Return mp_dtPosition
        End Get
        Set(ByVal Value As AGVBE.DateTime)
            mp_dtPosition = Value
        End Set
    End Property

    
    Public Property ForeColor() As Color
        Get
            Return mp_clrForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrForeColor = Value
        End Set
    End Property

    
    Public Property Length() As E_PROGRESSLINELENGTH
        Get
            Return mp_yLength
        End Get
        Set(ByVal Value As E_PROGRESSLINELENGTH)
            mp_yLength = Value
        End Set
    End Property

    
    Public Property LineType() As E_PROGRESSLINETYPE
        Get
            Return mp_yLineType
        End Get
        Set(ByVal Value As E_PROGRESSLINETYPE)
            mp_yLineType = Value
        End Set
    End Property

    Friend Sub Draw()
        Dim lXCoordinate As Integer
        Dim yTimeLineMarkerLength As E_PROGRESSLINELENGTH
        Dim dtDate As AGVBE.DateTime = New AGVBE.DateTime()
        If mp_yLineType = E_PROGRESSLINETYPE.TLMT_SYSTEMTIME Then
            dtDate.SetToCurrentDateTime()
        Else
            dtDate = mp_dtPosition
        End If
        If dtDate >= mp_oTimeLine.StartDate And dtDate <= mp_oTimeLine.EndDate Then
            yTimeLineMarkerLength = mp_yLength
            lXCoordinate = mp_oControl.MathLib.GetXCoordinateFromDate(mp_dtPosition)
            If mp_oTimeLine.TickMarkArea.Visible = False And yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_BOTH Then
                yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_CLIENTAREA
            End If
            If mp_oTimeLine.TickMarkArea.Visible = False And yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA Then
                yTimeLineMarkerLength = E_PROGRESSLINELENGTH.TLMA_NONE
            End If
            Select Case yTimeLineMarkerLength
                Case E_PROGRESSLINELENGTH.TLMA_TICKMARKAREA
                    mp_oControl.clsG.DrawLine(lXCoordinate, mp_oTimeLine.TiersTickMarksPosition("TickMarkArea"), lXCoordinate, mp_oTimeLine.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Case E_PROGRESSLINELENGTH.TLMA_CLIENTAREA
                    mp_oControl.clsG.DrawLine(lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Top, lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Case E_PROGRESSLINELENGTH.TLMA_BOTH
                    mp_oControl.clsG.DrawLine(lXCoordinate, mp_oTimeLine.TiersTickMarksPosition("TickMarkArea"), lXCoordinate, mp_oControl.CurrentViewObject.ClientArea.Bottom, GRE_LINETYPE.LT_NORMAL, mp_clrForeColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End Select
        End If
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ProgressLine")
        oXML.InitializeWriter()
        oXML.WriteProperty("ForeColor", mp_clrForeColor)
        oXML.WriteProperty("Length", mp_yLength)
        oXML.WriteProperty("LineType", mp_yLineType)
        oXML.WriteProperty("Position", mp_dtPosition)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ProgressLine")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("ForeColor", mp_clrForeColor)
        oXML.ReadProperty("Length", mp_yLength)
        oXML.ReadProperty("LineType", mp_yLineType)
        oXML.ReadProperty("Position", mp_dtPosition)
    End Sub

End Class
