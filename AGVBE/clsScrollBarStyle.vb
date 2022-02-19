Option Explicit On
Imports System.Windows.Media

Public Class clsScrollBarStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clrArrowColor As Color
    Private mp_clrDropShadowArrowColor As Color
    Private mp_bDropShadow As Boolean
    Private mp_lLeftX As Integer
    Private mp_lLeftY As Integer
    Private mp_lUpX As Integer
    Private mp_lUpY As Integer
    Private mp_lRightX As Integer
    Private mp_lRightY As Integer
    Private mp_lDownX As Integer
    Private mp_lDownY As Integer
    Private mp_lDropShadowLeftX As Integer
    Private mp_lDropShadowLeftY As Integer
    Private mp_lDropShadowUpX As Integer
    Private mp_lDropShadowUpY As Integer
    Private mp_lDropShadowRightX As Integer
    Private mp_lDropShadowRightY As Integer
    Private mp_lDropShadowDownX As Integer
    Private mp_lDropShadowDownY As Integer
    Private mp_lArrowSize As Integer

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_clrArrowColor = Colors.Black
        mp_clrDropShadowArrowColor = Color.FromArgb(255, 192, 192, 192)
        mp_bDropShadow = False
        mp_lArrowSize = 3
        mp_lLeftX = 6
        mp_lLeftY = 8
        mp_lUpX = 8
        mp_lUpY = 6
        mp_lRightX = 10
        mp_lRightY = 8
        mp_lDownX = 8
        mp_lDownY = 10
        mp_lDropShadowLeftX = 5
        mp_lDropShadowLeftY = 7
        mp_lDropShadowUpX = 7
        mp_lDropShadowUpY = 5
        mp_lDropShadowRightX = 9
        mp_lDropShadowRightY = 7
        mp_lDropShadowDownX = 7
        mp_lDropShadowDownY = 9
    End Sub

    Public Property ArrowColor() As Color
        Get
            Return mp_clrArrowColor
        End Get
        Set(ByVal Value As Color)
            mp_clrArrowColor = Value
        End Set
    End Property

    Public Property DropShadowArrowColor() As Color
        Get
            Return mp_clrDropShadowArrowColor
        End Get
        Set(ByVal Value As Color)
            mp_clrDropShadowArrowColor = Value
        End Set
    End Property

    Public Property DropShadow() As Boolean
        Get
            Return mp_bDropShadow
        End Get
        Set(ByVal Value As Boolean)
            mp_bDropShadow = Value
        End Set
    End Property

    Public Property ArrowSize() As Integer
        Get
            Return mp_lArrowSize
        End Get
        Set(ByVal value As Integer)
            mp_lArrowSize = value
        End Set
    End Property

    Public Property LeftX() As Integer
        Get
            Return mp_lLeftX
        End Get
        Set(ByVal value As Integer)
            mp_lLeftX = value
        End Set
    End Property

    Public Property LeftY() As Integer
        Get
            Return mp_lLeftY
        End Get
        Set(ByVal value As Integer)
            mp_lLeftY = value
        End Set
    End Property

    Public Property UpX() As Integer
        Get
            Return mp_lUpX
        End Get
        Set(ByVal value As Integer)
            mp_lUpX = value
        End Set
    End Property

    Public Property UpY() As Integer
        Get
            Return mp_lUpY
        End Get
        Set(ByVal value As Integer)
            mp_lUpY = value
        End Set
    End Property

    Public Property RightX() As Integer
        Get
            Return mp_lRightX
        End Get
        Set(ByVal value As Integer)
            mp_lRightX = value
        End Set
    End Property

    Public Property RightY() As Integer
        Get
            Return mp_lRightY
        End Get
        Set(ByVal value As Integer)
            mp_lRightY = value
        End Set
    End Property

    Public Property DownX() As Integer
        Get
            Return mp_lDownX
        End Get
        Set(ByVal value As Integer)
            mp_lDownX = value
        End Set
    End Property

    Public Property DownY() As Integer
        Get
            Return mp_lDownY
        End Get
        Set(ByVal value As Integer)
            mp_lDownY = value
        End Set
    End Property

    '

    Public Property DropShadowLeftX() As Integer
        Get
            Return mp_lDropShadowLeftX
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowLeftX = value
        End Set
    End Property

    Public Property DropShadowLeftY() As Integer
        Get
            Return mp_lDropShadowLeftY
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowLeftY = value
        End Set
    End Property

    Public Property DropShadowUpX() As Integer
        Get
            Return mp_lDropShadowUpX
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowUpX = value
        End Set
    End Property

    Public Property DropShadowUpY() As Integer
        Get
            Return mp_lDropShadowUpY
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowUpY = value
        End Set
    End Property

    Public Property DropShadowRightX() As Integer
        Get
            Return mp_lDropShadowRightX
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowRightX = value
        End Set
    End Property

    Public Property DropShadowRightY() As Integer
        Get
            Return mp_lDropShadowRightY
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowRightY = value
        End Set
    End Property

    Public Property DropShadowDownX() As Integer
        Get
            Return mp_lDropShadowDownX
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowDownX = value
        End Set
    End Property

    Public Property DropShadowDownY() As Integer
        Get
            Return mp_lDropShadowDownY
        End Get
        Set(ByVal value As Integer)
            mp_lDropShadowDownY = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ScrollBarStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("ArrowColor", mp_clrArrowColor)
        oXML.WriteProperty("DropShadowArrowColor", mp_clrDropShadowArrowColor)
        oXML.WriteProperty("DropShadow", mp_bDropShadow)
        oXML.WriteProperty("ArrowSize", mp_lArrowSize)
        oXML.WriteProperty("LeftX", mp_lLeftX)
        oXML.WriteProperty("LeftY", mp_lLeftY)
        oXML.WriteProperty("UpX", mp_lUpX)
        oXML.WriteProperty("UpY", mp_lUpY)
        oXML.WriteProperty("RightX", mp_lRightX)
        oXML.WriteProperty("RightY", mp_lRightY)
        oXML.WriteProperty("DownX", mp_lDownX)
        oXML.WriteProperty("DownY", mp_lDownY)
        oXML.WriteProperty("DropShadowLeftX", mp_lDropShadowLeftX)
        oXML.WriteProperty("DropShadowLeftY", mp_lDropShadowLeftY)
        oXML.WriteProperty("DropShadowUpX", mp_lDropShadowUpX)
        oXML.WriteProperty("DropShadowUpY", mp_lDropShadowUpY)
        oXML.WriteProperty("DropShadowRightX", mp_lDropShadowRightX)
        oXML.WriteProperty("DropShadowRightY", mp_lDropShadowRightY)
        oXML.WriteProperty("DropShadowDownX", mp_lDropShadowDownX)
        oXML.WriteProperty("DropShadowDownY", mp_lDropShadowDownY)

        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ScrollBarStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("ArrowColor", mp_clrArrowColor)
        oXML.ReadProperty("DropShadowArrowColor", mp_clrDropShadowArrowColor)
        oXML.ReadProperty("DropShadow", mp_bDropShadow)
        oXML.ReadProperty("ArrowSize", mp_lArrowSize)
        oXML.ReadProperty("LeftX", mp_lLeftX)
        oXML.ReadProperty("LeftY", mp_lLeftY)
        oXML.ReadProperty("UpX", mp_lUpX)
        oXML.ReadProperty("UpY", mp_lUpY)
        oXML.ReadProperty("RightX", mp_lRightX)
        oXML.ReadProperty("RightY", mp_lRightY)
        oXML.ReadProperty("DownX", mp_lDownX)
        oXML.ReadProperty("DownY", mp_lDownY)
        oXML.ReadProperty("DropShadowLeftX", mp_lDropShadowLeftX)
        oXML.ReadProperty("DropShadowLeftY", mp_lDropShadowLeftY)
        oXML.ReadProperty("DropShadowUpX", mp_lDropShadowUpX)
        oXML.ReadProperty("DropShadowUpY", mp_lDropShadowUpY)
        oXML.ReadProperty("DropShadowRightX", mp_lDropShadowRightX)
        oXML.ReadProperty("DropShadowRightY", mp_lDropShadowRightY)
        oXML.ReadProperty("DropShadowDownX", mp_lDropShadowDownX)
        oXML.ReadProperty("DropShadowDownY", mp_lDropShadowDownY)
    End Sub

End Class

