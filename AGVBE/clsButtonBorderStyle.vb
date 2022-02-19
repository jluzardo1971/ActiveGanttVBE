Option Explicit On
Imports System.Windows.Media

Public Class clsButtonBorderStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clrRaisedExteriorLeftTopColor As Color
    Private mp_clrRaisedInteriorLeftTopColor As Color
    Private mp_clrRaisedExteriorRightBottomColor As Color
    Private mp_clrRaisedInteriorRightBottomColor As Color
    Private mp_clrSunkenExteriorLeftTopColor As Color
    Private mp_clrSunkenInteriorLeftTopColor As Color
    Private mp_clrSunkenExteriorRightBottomColor As Color
    Private mp_clrSunkenInteriorRightBottomColor As Color

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_clrRaisedExteriorLeftTopColor = Color.FromArgb(255, 240, 240, 240)
        mp_clrRaisedInteriorLeftTopColor = Color.FromArgb(255, 192, 192, 192)
        mp_clrRaisedExteriorRightBottomColor = Colors.Gray
        mp_clrRaisedInteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
        mp_clrSunkenExteriorLeftTopColor = Colors.Gray
        mp_clrSunkenInteriorLeftTopColor = Color.FromArgb(255, 64, 64, 64)
        mp_clrSunkenExteriorRightBottomColor = Color.FromArgb(255, 240, 240, 240)
        mp_clrSunkenInteriorRightBottomColor = Color.FromArgb(255, 192, 192, 192)
    End Sub

    Public Property RaisedExteriorLeftTopColor() As Color
        Get
            Return mp_clrRaisedExteriorLeftTopColor
        End Get
        Set(ByVal value As Color)
            mp_clrRaisedExteriorLeftTopColor = value
        End Set
    End Property

    Public Property RaisedInteriorLeftTopColor() As Color
        Get
            Return mp_clrRaisedInteriorLeftTopColor
        End Get
        Set(ByVal value As Color)
            mp_clrRaisedInteriorLeftTopColor = value
        End Set
    End Property

    Public Property RaisedExteriorRightBottomColor() As Color
        Get
            Return mp_clrRaisedExteriorRightBottomColor
        End Get
        Set(ByVal value As Color)
            mp_clrRaisedExteriorRightBottomColor = value
        End Set
    End Property

    Public Property RaisedInteriorRightBottomColor() As Color
        Get
            Return mp_clrRaisedInteriorRightBottomColor
        End Get
        Set(ByVal value As Color)
            mp_clrRaisedInteriorRightBottomColor = value
        End Set
    End Property

    Public Property SunkenExteriorLeftTopColor() As Color
        Get
            Return mp_clrSunkenExteriorLeftTopColor
        End Get
        Set(ByVal value As Color)
            mp_clrSunkenExteriorLeftTopColor = value
        End Set
    End Property

    Public Property SunkenInteriorLeftTopColor() As Color
        Get
            Return mp_clrSunkenInteriorLeftTopColor
        End Get
        Set(ByVal value As Color)
            mp_clrSunkenInteriorLeftTopColor = value
        End Set
    End Property

    Public Property SunkenExteriorRightBottomColor() As Color
        Get
            Return mp_clrSunkenExteriorRightBottomColor
        End Get
        Set(ByVal value As Color)
            mp_clrSunkenExteriorRightBottomColor = value
        End Set
    End Property

    Public Property SunkenInteriorRightBottomColor() As Color
        Get
            Return mp_clrSunkenInteriorRightBottomColor
        End Get
        Set(ByVal value As Color)
            mp_clrSunkenInteriorRightBottomColor = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ButtonBorderStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("RaisedExteriorLeftTopColor", mp_clrRaisedExteriorLeftTopColor)
        oXML.WriteProperty("RaisedInteriorLeftTopColor", mp_clrRaisedInteriorLeftTopColor)
        oXML.WriteProperty("RaisedExteriorRightBottomColor", mp_clrRaisedExteriorRightBottomColor)
        oXML.WriteProperty("RaisedInteriorRightBottomColor", mp_clrRaisedInteriorRightBottomColor)
        oXML.WriteProperty("SunkenExteriorLeftTopColor", mp_clrSunkenExteriorLeftTopColor)
        oXML.WriteProperty("SunkenInteriorLeftTopColor", mp_clrSunkenInteriorLeftTopColor)
        oXML.WriteProperty("SunkenExteriorRightBottomColor", mp_clrSunkenExteriorRightBottomColor)
        oXML.WriteProperty("SunkenInteriorRightBottomColor", mp_clrSunkenInteriorRightBottomColor)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ButtonBorderStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("RaisedExteriorLeftTopColor", mp_clrRaisedExteriorLeftTopColor)
        oXML.ReadProperty("RaisedInteriorLeftTopColor", mp_clrRaisedInteriorLeftTopColor)
        oXML.ReadProperty("RaisedExteriorRightBottomColor", mp_clrRaisedExteriorRightBottomColor)
        oXML.ReadProperty("RaisedInteriorRightBottomColor", mp_clrRaisedInteriorRightBottomColor)
        oXML.ReadProperty("SunkenExteriorLeftTopColor", mp_clrSunkenExteriorLeftTopColor)
        oXML.ReadProperty("SunkenInteriorLeftTopColor", mp_clrSunkenInteriorLeftTopColor)
        oXML.ReadProperty("SunkenExteriorRightBottomColor", mp_clrSunkenExteriorRightBottomColor)
        oXML.ReadProperty("SunkenInteriorRightBottomColor", mp_clrSunkenInteriorRightBottomColor)
    End Sub

End Class