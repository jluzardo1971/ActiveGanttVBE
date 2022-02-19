Option Explicit On 

Public Class clsLayer
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bVisible As Boolean
    Private mp_sTag As String

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bVisible = True
        mp_sTag = ""
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.Layers.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.LAYERS_SET_KEY)
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

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Layer")
        oXML.InitializeWriter()
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Visible", mp_bVisible)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Layer")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Visible", mp_bVisible)
    End Sub




End Class
