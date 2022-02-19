Option Explicit On 

Public Class clsClientArea

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_oTimeLine As clsTimeLine
    Private mp_bDetectConflicts As Boolean
    Private mp_lMilestoneSelectionOffset As Integer
    Private mp_lPredecessorSelectionOffset As Integer
    Private mp_lLastVisibleRow As Integer
    Public Grid As clsGrid
    Private mp_sToolTipFormat As String
    Private mp_bToolTipsVisible As Boolean


    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oTimeLine As clsTimeLine)
        mp_oControl = Value
        mp_oTimeLine = oTimeLine
        mp_bDetectConflicts = True
        mp_lMilestoneSelectionOffset = 5
        mp_lLastVisibleRow = 0
        Grid = New clsGrid(mp_oControl, mp_oTimeLine)
        mp_sToolTipFormat = "ddddd"
        mp_bToolTipsVisible = True
        mp_lPredecessorSelectionOffset = 2
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    
    Public Property DetectConflicts() As Boolean
        Get
            Return mp_bDetectConflicts
        End Get
        Set(ByVal Value As Boolean)
            mp_bDetectConflicts = Value
        End Set
    End Property

    
    Public Property MilestoneSelectionOffset() As Integer
        Get
            Return mp_lMilestoneSelectionOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lMilestoneSelectionOffset = Value
        End Set
    End Property

    
    Public Property FirstVisibleRow() As Integer
        Get
            Return mp_oControl.VerticalScrollBar.Value
        End Get
        Set(ByVal Value As Integer)
            mp_oControl.VerticalScrollBar.Value = Value
        End Set
    End Property

    
    Public ReadOnly Property LastVisibleRow() As Integer
        Get
            Return mp_lLastVisibleRow
        End Get
    End Property

    Friend WriteOnly Property f_LastVisibleRow() As Integer
        Set(ByVal Value As Integer)
            mp_lLastVisibleRow = Value
        End Set
    End Property

    
    Public Property ToolTipFormat() As String
        Get
            Return mp_sToolTipFormat
        End Get
        Set(ByVal Value As String)
            mp_sToolTipFormat = Value
        End Set
    End Property

    
    Public Property ToolTipsVisible() As Boolean
        Get
            Return mp_bToolTipsVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bToolTipsVisible = Value
        End Set
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            If (mp_oTimeLine.Height = 0) Then
                Return mp_oControl.mt_BorderThickness
            Else
                Return mp_oTimeLine.Bottom + 1
            End If
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            If mp_oTimeLine.TimeLineScrollBar.State = 3 Then
                Return mp_oControl.clsG.Height - mp_oControl.mt_BorderThickness - 1 - mp_oTimeLine.TimeLineScrollBar.Height
            Else
                Return mp_oControl.clsG.Height - mp_oControl.mt_BorderThickness - 1
            End If
        End Get
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            Return mp_oTimeLine.f_lStart
        End Get
    End Property

    
    Public ReadOnly Property Right() As Integer
        Get
            Return mp_oTimeLine.f_lEnd
        End Get
    End Property

    
    Public ReadOnly Property Width() As Integer
        Get
            Return Right - Left
        End Get
    End Property

    
    Public ReadOnly Property Height() As Integer
        Get
            Return Bottom - Top
        End Get
    End Property

    Public Property PredecessorSelectionOffset() As Integer
        Get
            Return mp_lPredecessorSelectionOffset
        End Get
        Set(ByVal Value As Integer)
            mp_lPredecessorSelectionOffset = Value
        End Set
    End Property

    Friend Sub Draw()
        Dim lRowIndex As Integer
        Dim oRow As clsRow
        If mp_oControl.Rows.Count = 0 Then
            Return
        End If
        mp_oControl.clsG.ClipRegion(mp_oControl.Splitter.Right, mp_oControl.CurrentViewObject.ClientArea.Top, mp_oControl.mt_RightMargin, mp_oControl.CurrentViewObject.ClientArea.Bottom, True)
        For lRowIndex = mp_oControl.VerticalScrollBar.Value To mp_lLastVisibleRow
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lRowIndex), clsRow)
            mp_oControl.clsG.mp_DrawItem(mp_oControl.Splitter.Right, oRow.Top, mp_oControl.mt_RightMargin, oRow.Bottom, "", "", False, Nothing, 0, 0, oRow.ClientAreaStyle)
        Next lRowIndex
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "ClientArea")
        oXML.InitializeWriter()
        oXML.WriteProperty("DetectConflicts", mp_bDetectConflicts)
        oXML.WriteProperty("MilestoneSelectionOffset", mp_lMilestoneSelectionOffset)
        oXML.WriteProperty("ToolTipFormat", mp_sToolTipFormat)
        oXML.WriteProperty("ToolTipsVisible", mp_bToolTipsVisible)
        oXML.WriteProperty("PredecessorSelectionOffset", mp_lPredecessorSelectionOffset)
        oXML.WriteObject(Grid.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "ClientArea")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("DetectConflicts", mp_bDetectConflicts)
        oXML.ReadProperty("MilestoneSelectionOffset", mp_lMilestoneSelectionOffset)
        oXML.ReadProperty("ToolTipFormat", mp_sToolTipFormat)
        oXML.ReadProperty("ToolTipsVisible", mp_bToolTipsVisible)
        oXML.ReadProperty("PredecessorSelectionOffset", mp_lPredecessorSelectionOffset)
        Grid.SetXML(oXML.ReadObject("Grid"))
    End Sub

End Class
