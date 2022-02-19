Option Explicit On

Public Class clsTaskStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clrEndBorderColor As Color
    Private mp_clrEndFillColor As Color
    Private mp_clrStartBorderColor As Color
    Private mp_clrStartFillColor As Color

    Private mp_yEndShapeIndex As GRE_FIGURETYPE
    Private mp_yStartShapeIndex As GRE_FIGURETYPE

    Private mp_oStartImage As Image
    Private mp_oMiddleImage As Image
    Private mp_oEndImage As Image

    Private mp_sStartImageTag As String
    Private mp_sMiddleImageTag As String
    Private mp_sEndImageTag As String

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_clrEndBorderColor = Colors.Black
        mp_clrEndFillColor = Colors.Black
        mp_clrStartBorderColor = Colors.Black
        mp_clrStartFillColor = Colors.Black
        mp_oEndImage = Nothing
        mp_oMiddleImage = Nothing
        mp_oStartImage = Nothing
        mp_yEndShapeIndex = GRE_FIGURETYPE.FT_NONE
        mp_yStartShapeIndex = GRE_FIGURETYPE.FT_NONE
        mp_sStartImageTag = ""
        mp_sMiddleImageTag = ""
        mp_sEndImageTag = ""
    End Sub

    
    Public Property EndBorderColor() As Color
        Get
            Return mp_clrEndBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrEndBorderColor = Value
        End Set
    End Property

    
    Public Property EndFillColor() As Color
        Get
            Return mp_clrEndFillColor
        End Get
        Set(ByVal Value As Color)
            mp_clrEndFillColor = Value
        End Set
    End Property

    
    Public Property StartBorderColor() As Color
        Get
            Return mp_clrStartBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrStartBorderColor = Value
        End Set
    End Property

    
    Public Property StartFillColor() As Color
        Get
            Return mp_clrStartFillColor
        End Get
        Set(ByVal Value As Color)
            mp_clrStartFillColor = Value
        End Set
    End Property

    
    Public Property StartImage() As Image
        Get
            Return mp_oStartImage
        End Get
        Set(ByVal Value As Image)
            mp_oStartImage = Value
        End Set
    End Property

    
    Public Property MiddleImage() As Image
        Get
            Return mp_oMiddleImage
        End Get
        Set(ByVal Value As Image)
            mp_oMiddleImage = Value
        End Set
    End Property

    
    Public Property EndImage() As Image
        Get
            Return mp_oEndImage
        End Get
        Set(ByVal Value As Image)
            mp_oEndImage = Value
        End Set
    End Property

    
    Public Property StartShapeIndex() As GRE_FIGURETYPE
        Get
            Return mp_yStartShapeIndex
        End Get
        Set(ByVal Value As GRE_FIGURETYPE)
            mp_yStartShapeIndex = Value
        End Set
    End Property

    
    Public Property EndShapeIndex() As GRE_FIGURETYPE
        Get
            Return mp_yEndShapeIndex
        End Get
        Set(ByVal Value As GRE_FIGURETYPE)
            mp_yEndShapeIndex = Value
        End Set
    End Property

    Public Property StartImageTag() As String
        Get
            Return mp_sStartImageTag
        End Get
        Set(ByVal value As String)
            mp_sStartImageTag = value
        End Set
    End Property

    Public Property MiddleImageTag() As String
        Get
            Return mp_sMiddleImageTag
        End Get
        Set(ByVal value As String)
            mp_sMiddleImageTag = value
        End Set
    End Property

    Public Property EndImageTag() As String
        Get
            Return mp_sEndImageTag
        End Get
        Set(ByVal value As String)
            mp_sEndImageTag = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TaskStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("EndBorderColor", mp_clrEndBorderColor)
        oXML.WriteProperty("EndFillColor", mp_clrEndFillColor)
        oXML.WriteProperty("EndImage", mp_oEndImage)
        oXML.WriteProperty("EndShapeIndex", mp_yEndShapeIndex)
        oXML.WriteProperty("MiddleImage", mp_oMiddleImage)
        oXML.WriteProperty("StartBorderColor", mp_clrStartBorderColor)
        oXML.WriteProperty("StartFillColor", mp_clrStartFillColor)
        oXML.WriteProperty("StartImage", mp_oStartImage)
        oXML.WriteProperty("StartShapeIndex", mp_yStartShapeIndex)
        oXML.WriteProperty("StartImageTag", mp_sStartImageTag)
        oXML.WriteProperty("MiddleImageTag", mp_sMiddleImageTag)
        oXML.WriteProperty("EndImageTag", mp_sEndImageTag)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TaskStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("EndBorderColor", mp_clrEndBorderColor)
        oXML.ReadProperty("EndFillColor", mp_clrEndFillColor)
        oXML.ReadProperty("EndImage", mp_oEndImage)
        oXML.ReadProperty("EndShapeIndex", mp_yEndShapeIndex)
        oXML.ReadProperty("MiddleImage", mp_oMiddleImage)
        oXML.ReadProperty("StartBorderColor", mp_clrStartBorderColor)
        oXML.ReadProperty("StartFillColor", mp_clrStartFillColor)
        oXML.ReadProperty("StartImage", mp_oStartImage)
        oXML.ReadProperty("StartShapeIndex", mp_yStartShapeIndex)
        oXML.ReadProperty("StartImageTag", mp_sStartImageTag)
        oXML.ReadProperty("MiddleImageTag", mp_sMiddleImageTag)
        oXML.ReadProperty("EndImageTag", mp_sEndImageTag)
    End Sub

End Class
