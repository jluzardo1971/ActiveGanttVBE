Option Explicit On 

Public Class clsView
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Public TimeLine As clsTimeLine
    Public ClientArea As clsClientArea
    Private mp_sTag As String
    Private mp_yScrollInterval As E_INTERVAL
    Private mp_yInterval As E_INTERVAL
    Private mp_lFactor As Integer

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_yInterval = E_INTERVAL.IL_SECOND
        mp_lFactor = 1
        mp_yScrollInterval = E_INTERVAL.IL_HOUR
        TimeLine = New clsTimeLine(mp_oControl, Me)
        ClientArea = New clsClientArea(mp_oControl, TimeLine)
        mp_sTag = ""
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.Views.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.VIEWS_SET_KEY)
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

    Friend ReadOnly Property f_ScrollInterval() As E_INTERVAL
        Get
            Return mp_yScrollInterval
        End Get
    End Property

    Public Property Interval() As E_INTERVAL
        Get
            Return mp_yInterval
        End Get
        Set(ByVal Value As E_INTERVAL)
            mp_yInterval = Value
            If mp_yInterval = E_INTERVAL.IL_YEAR Then
                mp_yScrollInterval = E_INTERVAL.IL_YEAR
            Else
                mp_yScrollInterval = DirectCast(mp_yInterval + 1, AGVBE.E_INTERVAL)
            End If
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

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "View")
        oXML.InitializeWriter()
        oXML.WriteProperty("Interval", mp_yInterval)
        oXML.WriteProperty("Factor", mp_lFactor)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteObject(ClientArea.GetXML())
        oXML.WriteObject(TimeLine.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "View")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Interval", mp_yInterval)
        oXML.ReadProperty("Factor", mp_lFactor)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("Tag", mp_sTag)
        ClientArea.SetXML(oXML.ReadObject("ClientArea"))
        TimeLine.SetXML(oXML.ReadObject("TimeLine"))
    End Sub

End Class
