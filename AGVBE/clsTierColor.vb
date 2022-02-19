Option Explicit On

Public Class clsTierColor
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clsTierColors As clsTierColors
    Private mp_clrBackColor As Color
    Private mp_clrForeColor As Color
    Private mp_clrStartGradientColor As Color
    Private mp_clrEndGradientColor As Color
    Private mp_clrHatchBackColor As Color
    Private mp_clrHatchForeColor As Color


    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTierColors As clsTierColors)
        mp_oControl = Value
        mp_clsTierColors = oTierColors
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_clsTierColors.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.TIERCOLORS_SET_KEY)
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

    
    Public Property BackColor() As Color
        Get
            Return mp_clrBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrBackColor = Value
        End Set
    End Property

    Public Property StartGradientColor() As Color
        Get
            Return mp_clrStartGradientColor
        End Get
        Set(ByVal Value As Color)
            mp_clrStartGradientColor = Value
        End Set
    End Property

    Public Property EndGradientColor() As Color
        Get
            Return mp_clrEndGradientColor
        End Get
        Set(ByVal Value As Color)
            mp_clrEndGradientColor = Value
        End Set
    End Property

    Public Property HatchBackColor() As Color
        Get
            Return mp_clrHatchBackColor
        End Get
        Set(ByVal Value As Color)
            mp_clrHatchBackColor = Value
        End Set
    End Property

    Public Property HatchForeColor() As Color
        Get
            Return mp_clrHatchForeColor
        End Get
        Set(ByVal Value As Color)
            mp_clrHatchForeColor = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TierColor")
        oXML.InitializeWriter()
        oXML.WriteProperty("BackColor", mp_clrBackColor)
        oXML.WriteProperty("ForeColor", mp_clrForeColor)
        oXML.WriteProperty("StartGradientColor", mp_clrStartGradientColor)
        oXML.WriteProperty("EndGradientColor", mp_clrEndGradientColor)
        oXML.WriteProperty("HatchBackColor", mp_clrHatchBackColor)
        oXML.WriteProperty("HatchForeColor", mp_clrHatchForeColor)
        oXML.WriteProperty("Key", mp_sKey)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TierColor")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("BackColor", mp_clrBackColor)
        oXML.ReadProperty("ForeColor", mp_clrForeColor)
        oXML.ReadProperty("StartGradientColor", mp_clrStartGradientColor)
        oXML.ReadProperty("EndGradientColor", mp_clrEndGradientColor)
        oXML.ReadProperty("HatchBackColor", mp_clrHatchBackColor)
        oXML.ReadProperty("HatchForeColor", mp_clrHatchForeColor)
        oXML.ReadProperty("Key", mp_sKey)
    End Sub




End Class
