Option Explicit On 

Public Class clsHorizontalScrollBar

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bVisible As Boolean
    Public WithEvents ScrollBar As clsHScrollBarTemplate

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bVisible = True
        ScrollBar = New clsHScrollBarTemplate()
        ScrollBar.Initialize(mp_oControl)
        ScrollBar.LargeChange = 1
        ScrollBar.SmallChange = 1
        ScrollBar.Min = 0
        ScrollBar.Max = 0
        ScrollBar.Value = 0
    End Sub

    
    Public ReadOnly Property Min() As Integer
        Get
            Return 0
        End Get
    End Property

    
    Public ReadOnly Property Max() As Integer
        Get
            Return ScrollBar.Max
        End Get
    End Property

    
    Public Property Value() As Integer
        Get
            Return ScrollBar.Value
        End Get
        Set(ByVal Value1 As Integer)
            ScrollBar.Value = Value1
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

    
    Public Property SmallChange() As Integer
        Get
            Return ScrollBar.SmallChange
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.SmallChange = Value
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

    Friend Property State() As Integer
        Get
            Return ScrollBar.State
        End Get
        Set(ByVal Value As Integer)
            ScrollBar.State = Value
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

    Friend Sub Reset()
        ScrollBar.Min = 0
        ScrollBar.Max = 0
        ScrollBar.Value = 0
    End Sub

    Friend Sub Position()
        Left = mp_oControl.mt_BorderThickness
        Top = mp_oControl.clsG.Height - Height - mp_oControl.mt_BorderThickness
        If mp_oControl.Splitter.Left > 0 Then
            Width = mp_oControl.Splitter.Left - 1
        End If
        ScrollBar.Max = mp_oControl.Columns.Width - mp_oControl.Splitter.Position '//Leave Splitter.Position
    End Sub

    Private Sub ScrollBar_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs, ByVal Offset As Integer) Handles ScrollBar.ValueChanged
        mp_oControl.HorizontalScrollBar_ValueChanged(Offset)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "HorizontalScrollBar")
        oXML.InitializeWriter()
        oXML.WriteObject(ScrollBar.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "HorizontalScrollBar")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        ScrollBar.SetXML(oXML.ReadObject("ScrollBar"))
    End Sub

End Class
