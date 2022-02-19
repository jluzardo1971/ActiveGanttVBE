Option Explicit On

Public Class clsPredecessorStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lArrowSize As Integer
    Private mp_yLineStyle As GRE_LINEDRAWSTYLE
    Private mp_lLineWidth As Integer
    Private mp_lXOffset As Integer
    Private mp_lYOffset As Integer
    Private mp_clrLineColor As Color

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_lArrowSize = 3
        mp_yLineStyle = GRE_LINEDRAWSTYLE.LDS_SOLID
        mp_lLineWidth = 1
        mp_lXOffset = 10
        mp_lYOffset = 10
        mp_clrLineColor = Colors.Black
    End Sub

    Public Property LineColor() As Color
        Get
            Return mp_clrLineColor
        End Get
        Set(ByVal Value As Color)
            mp_clrLineColor = Value
        End Set
    End Property

    Public Property XOffset() As Integer
        Get
            Return mp_lXOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lXOffset = Value
        End Set
    End Property


    Public Property YOffset() As Integer
        Get
            Return mp_lYOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lYOffset = Value
        End Set
    End Property


    Public Property LineWidth() As Integer
        Get
            Return mp_lLineWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lLineWidth = Value
        End Set
    End Property


    Public Property LineStyle() As GRE_LINEDRAWSTYLE
        Get
            Return mp_yLineStyle
        End Get
        Set(ByVal Value As GRE_LINEDRAWSTYLE)
            mp_yLineStyle = Value
        End Set
    End Property


    Public Property ArrowSize() As Integer
        Get
            Return mp_lArrowSize
        End Get
        Set(ByVal Value As Integer)
            If (Value < 1) Then
                Value = 1
            End If
            mp_lArrowSize = Value
        End Set
    End Property


    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "PredecessorStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("ArrowSize", mp_lArrowSize)
        oXML.WriteProperty("LineColor", mp_clrLineColor)
        oXML.WriteProperty("LineStyle", mp_yLineStyle)
        oXML.WriteProperty("LineWidth", mp_lLineWidth)
        oXML.WriteProperty("XOffset", mp_lXOffset)
        oXML.WriteProperty("YOffset", mp_lYOffset)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "PredecessorStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("ArrowSize", mp_lArrowSize)
        oXML.ReadProperty("LineColor", mp_clrLineColor)
        oXML.ReadProperty("LineStyle", mp_yLineStyle)
        oXML.ReadProperty("LineWidth", mp_lLineWidth)
        oXML.ReadProperty("XOffset", mp_lXOffset)
        oXML.ReadProperty("YOffset", mp_lYOffset)
    End Sub

End Class
