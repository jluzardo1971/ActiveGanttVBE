Option Explicit On 

Public Class clsTickMark
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bDisplayText As Boolean
    Private mp_clsTickMarks As clsTickMarks
    Private mp_sTextFormat As String
    Private mp_sTag As String
    Private mp_yTickMarkType As E_TICKMARKTYPES
    Private mp_yInterval As E_INTERVAL
    Private mp_lFactor As Integer

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTickMarks As clsTickMarks)
        mp_oControl = Value
        mp_bDisplayText = False
        mp_clsTickMarks = oTickMarks
        mp_sTextFormat = ""
        mp_sTag = ""
        mp_yInterval = E_INTERVAL.IL_SECOND
        mp_lFactor = 1
        mp_yTickMarkType = E_TICKMARKTYPES.TLT_SMALL
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_clsTickMarks.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.TICKMARKS_SET_KEY)
        End Set
    End Property

    
    Public Property DisplayText() As Boolean
        Get
            Return mp_bDisplayText
        End Get
        Set(ByVal Value As Boolean)
            mp_bDisplayText = Value
        End Set
    End Property

    
    Public Property TextFormat() As String
        Get
            Return mp_sTextFormat
        End Get
        Set(ByVal Value As String)
            mp_sTextFormat = Value
        End Set
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
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

    
    Public Property TickMarkType() As E_TICKMARKTYPES
        Get
            Return mp_yTickMarkType
        End Get
        Set(ByVal Value As E_TICKMARKTYPES)
            mp_yTickMarkType = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TickMark")
        oXML.InitializeWriter()
        oXML.WriteProperty("DisplayText", mp_bDisplayText)
        oXML.WriteProperty("Interval", mp_yInterval)
        oXML.WriteProperty("Factor", mp_lFactor)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("TextFormat", mp_sTextFormat)
        oXML.WriteProperty("TickMarkType", mp_yTickMarkType)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TickMark")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("DisplayText", mp_bDisplayText)
        oXML.ReadProperty("Interval", mp_yInterval)
        oXML.ReadProperty("Factor", mp_lFactor)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("TextFormat", mp_sTextFormat)
        oXML.ReadProperty("TickMarkType", mp_yTickMarkType)
    End Sub




End Class
