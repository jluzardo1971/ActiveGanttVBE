Option Explicit On

Public Class clsMilestoneStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_clrBorderColor As Color
    Private mp_clrFillColor As Color
    Private mp_yShapeIndex As GRE_FIGURETYPE
    Private mp_oImage As Image
    Private mp_sImageTag As String

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        '// Parent Control Pointer
        mp_oControl = Value
        '// Object Member Variables
        mp_clrBorderColor = Colors.Black
        mp_clrFillColor = Colors.Black
        mp_yShapeIndex = GRE_FIGURETYPE.FT_NONE
        mp_oImage = Nothing
        mp_sImageTag = ""
    End Sub

    
    Public Property BorderColor() As Color
        Get
            Return mp_clrBorderColor
        End Get
        Set(ByVal Value As Color)
            mp_clrBorderColor = Value
        End Set
    End Property

    
    Public Property FillColor() As Color
        Get
            Return mp_clrFillColor
        End Get
        Set(ByVal Value As Color)
            mp_clrFillColor = Value
        End Set
    End Property

    
    Public Property ShapeIndex() As GRE_FIGURETYPE
        Get
            Return mp_yShapeIndex
        End Get
        Set(ByVal Value As GRE_FIGURETYPE)
            mp_yShapeIndex = Value
        End Set
    End Property

    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal value As String)
            mp_sImageTag = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "MilestoneStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("BorderColor", mp_clrBorderColor)
        oXML.WriteProperty("FillColor", mp_clrFillColor)
        oXML.WriteProperty("ShapeIndex", mp_yShapeIndex)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "MilestoneStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("BorderColor", mp_clrBorderColor)
        oXML.ReadProperty("FillColor", mp_clrFillColor)
        oXML.ReadProperty("ShapeIndex", mp_yShapeIndex)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
    End Sub





End Class
