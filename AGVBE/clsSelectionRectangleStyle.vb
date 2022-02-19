Option Explicit On
Imports System.Windows.Media

Public Class clsSelectionRectangleStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_lOffsetBottom As Integer
    Private mp_lOffsetLeft As Integer
    Private mp_lOffsetRight As Integer
    Private mp_lOffsetTop As Integer
    Private mp_bVisible As Boolean
    Private mp_yMode As E_SELECTIONRECTANGLEMODE
    Private mp_lBorderWidth As Integer
    Private mp_clrColor As Color


    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bVisible = True
        mp_lOffsetBottom = 3
        mp_lOffsetLeft = 3
        mp_lOffsetRight = 3
        mp_lOffsetTop = 3
        mp_yMode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED
        mp_lBorderWidth = 1
        mp_clrColor = Colors.Black
    End Sub

    Public Property OffsetBottom() As Integer
        Get
            Return mp_lOffsetBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetBottom = Value
        End Set
    End Property

    Public Property OffsetLeft() As Integer
        Get
            Return mp_lOffsetLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetLeft = Value
        End Set
    End Property

    Public Property OffsetRight() As Integer
        Get
            Return mp_lOffsetRight
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetRight = Value
        End Set
    End Property

    Public Property OffsetTop() As Integer
        Get
            Return mp_lOffsetTop
        End Get
        Set(ByVal Value As Integer)
            mp_lOffsetTop = Value
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

    Public Property Mode() As E_SELECTIONRECTANGLEMODE
        Get
            Return mp_yMode
        End Get
        Set(ByVal value As E_SELECTIONRECTANGLEMODE)
            mp_yMode = value
        End Set
    End Property

    Public Property BorderWidth() As Integer
        Get
            Return mp_lBorderWidth
        End Get
        Set(ByVal value As Integer)
            mp_lBorderWidth = value
        End Set
    End Property

    Public Property Color() As Color
        Get
            Return mp_clrColor
        End Get
        Set(ByVal value As Color)
            mp_clrColor = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "SelectionRectangleStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.WriteProperty("OffsetLeft", mp_lOffsetLeft)
        oXML.WriteProperty("OffsetRight", mp_lOffsetRight)
        oXML.WriteProperty("OffsetTop", mp_lOffsetTop)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("Mode", mp_yMode)
        oXML.WriteProperty("BorderWidth", mp_lBorderWidth)
        oXML.WriteProperty("Color", mp_clrColor)

        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "SelectionRectangleStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("OffsetBottom", mp_lOffsetBottom)
        oXML.ReadProperty("OffsetLeft", mp_lOffsetLeft)
        oXML.ReadProperty("OffsetRight", mp_lOffsetRight)
        oXML.ReadProperty("OffsetTop", mp_lOffsetTop)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("Mode", mp_yMode)
        oXML.ReadProperty("BorderWidth", mp_lBorderWidth)
        oXML.ReadProperty("Color", mp_clrColor)
    End Sub

End Class

