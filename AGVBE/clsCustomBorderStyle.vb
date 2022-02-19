Public Class clsCustomBorderStyle

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bTop As Boolean
    Private mp_bBottom As Boolean
    Private mp_bLeft As Boolean
    Private mp_bRight As Boolean

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bTop = True
        mp_bBottom = True
        mp_bLeft = True
        mp_bRight = True
    End Sub

    
    Public Property Left() As Boolean
        Get
            Return mp_bLeft
        End Get
        Set(ByVal Value As Boolean)
            mp_bLeft = Value
        End Set
    End Property

    
    Public Property Top() As Boolean
        Get
            Return mp_bTop
        End Get
        Set(ByVal Value As Boolean)
            mp_bTop = Value
        End Set
    End Property

    
    Public Property Right() As Boolean
        Get
            Return mp_bRight
        End Get
        Set(ByVal Value As Boolean)
            mp_bRight = Value
        End Set
    End Property

    
    Public Property Bottom() As Boolean
        Get
            Return mp_bBottom
        End Get
        Set(ByVal Value As Boolean)
            mp_bBottom = Value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "CustomBorderStyle")
        oXML.InitializeWriter()
        oXML.WriteProperty("Bottom", mp_bBottom)
        oXML.WriteProperty("Left", mp_bLeft)
        oXML.WriteProperty("Right", mp_bRight)
        oXML.WriteProperty("Top", mp_bTop)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "CustomBorderStyle")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Bottom", mp_bBottom)
        oXML.ReadProperty("Left", mp_bLeft)
        oXML.ReadProperty("Right", mp_bRight)
        oXML.ReadProperty("Top", mp_bTop)
    End Sub

End Class
