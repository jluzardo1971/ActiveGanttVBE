Option Explicit On 

Public Class clsTimeLineScrollBar

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_dtStartDate As AGVBE.DateTime
    Private mp_yInterval As E_INTERVAL
    Private mp_lFactor As Integer
    Private mp_bVisible As Boolean
    Public WithEvents ScrollBar As clsHScrollBarTemplate

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        ScrollBar = New clsHScrollBarTemplate()
        ScrollBar.Initialize(mp_oControl)
        mp_dtStartDate = New AGVBE.DateTime()
        mp_dtStartDate.SetToCurrentDateTime()
        mp_yInterval = E_INTERVAL.IL_MINUTE
        mp_lFactor = 1
        ScrollBar.Enabled = False
        mp_bVisible = False
    End Sub

    
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

    
    Public Property Value() As Integer
        Get
            Return ScrollBar.Value
        End Get
        Set(ByVal lValue As Integer)
            If lValue < 0 Then
                lValue = 0
            End If
            If lValue > ScrollBar.Max Then
                lValue = ScrollBar.Max
            End If
            ScrollBar.Value = lValue
        End Set
    End Property

    
    Public Property Enabled() As Boolean
        Get
            Return ScrollBar.Enabled
        End Get
        Set(ByVal Value As Boolean)
            ScrollBar.Enabled = Value
        End Set
    End Property

    Friend ReadOnly Property mf_Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
    End Property

    
    Public Property Visible() As Boolean
        Get
            If ScrollBar.State <> E_SCROLLSTATE.SS_SHOWN Then
                Return False
            Else
                Return mp_bVisible
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    
    Public Property LargeChange() As Integer
        Get
            Return ScrollBar.LargeChange
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.LargeChange = Value
        End Set
    End Property

    
    Public Property Max() As Integer
        Get
            Return ScrollBar.Max
        End Get
        Set(ByVal Value As Integer)
            If Value < ScrollBar.Min Then
                Return
            End If
            ScrollBar.Max = Value
        End Set
    End Property

    
    Public Property SmallChange() As Integer
        Get
            Return ScrollBar.SmallChange
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.SmallChange = Value
        End Set
    End Property

    
    Public Property StartDate() As AGVBE.DateTime
        Get
            Return mp_dtStartDate
        End Get
        Set(ByVal Value As AGVBE.DateTime)
            mp_dtStartDate = Value
        End Set
    End Property

    Friend Property State() As Integer
        Get
            Return ScrollBar.State
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.State = Value
        End Set
    End Property

    Friend Property Width() As Integer
        Get
            Return ScrollBar.Width
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.Width = Value
        End Set
    End Property

    Friend Property Height() As Integer
        Get
            Return ScrollBar.Height
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.Height = Value
        End Set
    End Property

    Friend Property Left() As Integer
        Get
            Return ScrollBar.Left
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.Left = Value
        End Set
    End Property

    Friend Property Top() As Integer
        Get
            Return ScrollBar.Top
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.Top = Value
        End Set
    End Property

    Friend Sub Position()
        Left = mp_oControl.Splitter.Right
        Top = mp_oControl.clsG.Height - Height - mp_oControl.mt_BorderThickness
        Width = mp_oControl.clsG.Width - mp_oControl.mt_BorderThickness - mp_oControl.Splitter.Right - mp_oControl.VerticalScrollBar.Width
    End Sub

    Private Sub oHScrollBar_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs, ByVal Offset As Integer) Handles ScrollBar.ValueChanged
        mp_oControl.TimeLineScrollBar_ValueChanged(Offset)
    End Sub

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TimeLineScrollBar")
        oXML.InitializeWriter()
        oXML.WriteProperty("Enabled", ScrollBar.mp_bEnabled)
        oXML.WriteProperty("Interval", mp_yInterval)
        oXML.WriteProperty("Factor", mp_lFactor)
        oXML.WriteProperty("LargeChange", ScrollBar.mp_lLargeChange)
        oXML.WriteProperty("Max", ScrollBar.mp_lMax)
        oXML.WriteProperty("SmallChange", ScrollBar.mp_lSmallChange)
        oXML.WriteProperty("StartDate", mp_dtStartDate)
        oXML.WriteProperty("Value", ScrollBar.mp_lValue)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteObject(ScrollBar.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TimeLineScrollBar")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Enabled", ScrollBar.mp_bEnabled)
        oXML.ReadProperty("Interval", mp_yInterval)
        oXML.ReadProperty("Factor", mp_lFactor)
        oXML.ReadProperty("LargeChange", ScrollBar.mp_lLargeChange)
        oXML.ReadProperty("Max", ScrollBar.mp_lMax)
        oXML.ReadProperty("SmallChange", ScrollBar.mp_lSmallChange)
        oXML.ReadProperty("StartDate", mp_dtStartDate)
        oXML.ReadProperty("Value", ScrollBar.mp_lValue)
        oXML.ReadProperty("Visible", mp_bVisible)
        ScrollBar.SetXML(oXML.ReadObject("ScrollBar"))
    End Sub

End Class
