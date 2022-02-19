Option Explicit On 

Public Class clsVerticalScrollBar

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bVisible As Boolean
    Public WithEvents ScrollBar As clsVScrollBarTemplate

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bVisible = True
        ScrollBar = New clsVScrollBarTemplate()
        ScrollBar.Initialize(mp_oControl)
        ScrollBar.LargeChange = 1
        ScrollBar.SmallChange = 1
        ScrollBar.Min = 1
        ScrollBar.Max = 1
        ScrollBar.Value = 1
    End Sub

    
    Public ReadOnly Property Min() As Integer
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return 0
            Else
                Return ScrollBar.Min
            End If
        End Get
    End Property

    
    Public ReadOnly Property Max() As Integer
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return 0
            Else
                ScrollBar.Max = mp_oControl.Rows.Count - mp_oControl.Rows.HiddenRows()
                Return mp_oControl.Rows.Count - mp_oControl.Rows.HiddenRows()
            End If
        End Get
    End Property

    
    Public Property Value() As Integer
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return 0
            Else
                Return ScrollBar.Value
            End If
        End Get
        Set(ByVal Value1 As Integer)
            If mp_oControl.Rows.Count > 0 Then
                If Value1 < 1 Then
                    Value1 = 1
                End If
                If Value1 > mp_oControl.Rows.Count Then
                    Value1 = mp_oControl.Rows.Count
                End If
                ScrollBar.Value = Value1
            End If
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

    Friend Sub Update()
        Dim lHiddenRows As Integer
        lHiddenRows = mp_oControl.Rows.HiddenRows
        If mp_oControl.Rows.Count > 0 Then
            If ScrollBar.Value > (mp_oControl.Rows.Count - lHiddenRows) Then
                ScrollBar.Value = (mp_oControl.Rows.Count - lHiddenRows)
            End If
            ScrollBar.Max = (mp_oControl.Rows.Count - lHiddenRows)
        Else
            Reset()
        End If
    End Sub

    Friend Sub Reset()
        ScrollBar.Min = 1
        ScrollBar.Max = 1
        ScrollBar.Value = 1
    End Sub

    Friend Sub Position()
        Left = mp_oControl.clsG.Width - Width - mp_oControl.mt_BorderThickness
        Top = mp_oControl.mt_TopMargin
        Height = mp_oControl.clsG.Height - (mp_oControl.mt_BorderThickness * 2) - mp_oControl.HorizontalScrollBar.Height
        SmallChange = 1
    End Sub

    Private Sub ScrollBar_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs, ByVal Offset As Integer) Handles ScrollBar.ValueChanged
        mp_oControl.VerticalScrollBar_ValueChanged(Offset)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "VerticalScrollBar")
        oXML.InitializeWriter()
        oXML.WriteObject(ScrollBar.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "VerticalScrollBar")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        ScrollBar.SetXML(oXML.ReadObject("ScrollBar"))
    End Sub

End Class
