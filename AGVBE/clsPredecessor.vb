Option Explicit On 

Public Class clsPredecessor
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bVisible As Boolean
    Private mp_clsPredecessors As clsPredecessors
    Private mp_sSuccessorKey As String
    Friend mp_oSuccessorTask As clsTask
    Private mp_sPredecessorKey As String
    Friend mp_oPredecessorTask As clsTask
    Private mp_sStyleIndex As String
    Private mp_sTag As String
    Private mp_yPredecessorType As E_CONSTRAINTTYPE
    Private mp_oStyle As clsStyle
    Private mp_yLagInterval As E_INTERVAL
    Private mp_lLagFactor As Integer
    Private mp_oRectangles As ArrayList
    Friend mp_bWarning As Boolean
    Private mp_sWarningStyleIndex As String
    Private mp_oWarningStyle As clsStyle
    Private mp_sSelectedStyleIndex As String
    Private mp_oSelectedStyle As clsStyle

    Friend Sub New(ByVal Value As ActiveGanttVBECtl, ByVal oPredecessors As clsPredecessors)
        mp_oControl = Value
        mp_bVisible = True
        mp_clsPredecessors = oPredecessors
        mp_sPredecessorKey = ""
        mp_sSuccessorKey = ""
        mp_sStyleIndex = "DS_PREDECESSOR"
        mp_oStyle = mp_oControl.Styles.FItem("DS_PREDECESSOR")
        mp_sTag = ""
        mp_yLagInterval = E_INTERVAL.IL_DAY
        mp_lLagFactor = 0
        mp_oRectangles = New ArrayList()
        mp_bWarning = False
        mp_sWarningStyleIndex = ""
        mp_sSelectedStyleIndex = ""
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_clsPredecessors.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.PREDECESSORS_SET_KEY)
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

    
    Public Property PredecessorKey() As String
        Get
            Return mp_sPredecessorKey
        End Get
        Set(ByVal Value As String)
            mp_sPredecessorKey = Value
            mp_oPredecessorTask = mp_oControl.Tasks.Item(Value)
        End Set
    End Property

    
    Public ReadOnly Property PredecessorTask() As clsTask
        Get
            Return mp_oPredecessorTask
        End Get
    End Property

    
    Public Property PredecessorType() As E_CONSTRAINTTYPE
        Get
            Return mp_yPredecessorType
        End Get
        Set(ByVal Value As E_CONSTRAINTTYPE)
            mp_yPredecessorType = Value
        End Set
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_PREDECESSOR" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            If Trim(Value) = "" Then Value = "DS_PREDECESSOR"
            mp_sStyleIndex = Value
            mp_oStyle = mp_oControl.Styles.FItem(Value)
        End Set
    End Property

    
    Public ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    
    Public Property Tag() As String
        Get
            Return mp_sTag
        End Get
        Set(ByVal Value As String)
            mp_sTag = Value
        End Set
    End Property

    Public Property SuccessorKey() As String
        Get
            Return mp_sSuccessorKey
        End Get
        Set(ByVal Value As String)
            mp_sSuccessorKey = Value
            mp_oSuccessorTask = mp_oControl.Tasks.Item(Value)
        End Set
    End Property

    Public ReadOnly Property SuccessorTask() As clsTask
        Get
            Return mp_oSuccessorTask
        End Get
    End Property

    Public Property LagInterval() As E_INTERVAL
        Get
            Return mp_yLagInterval
        End Get
        Set(ByVal Value As E_INTERVAL)
            mp_yLagInterval = Value
        End Set
    End Property

    Public Property LagFactor() As Integer
        Get
            Return mp_lLagFactor
        End Get
        Set(ByVal value As Integer)
            mp_lLagFactor = value
        End Set
    End Property

    Friend Sub AddRectangle(ByVal oRectangle As S_Rectangle)
        mp_oRectangles.Add(oRectangle)
    End Sub

    Friend Sub ClearRectangles()
        mp_oRectangles.Clear()
    End Sub

    Friend Function HitTest(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim i As Integer
        For i = 0 To mp_oRectangles.Count - 1
            Dim oRectangle As S_Rectangle
            oRectangle = DirectCast(mp_oRectangles.Item(i), AGVBE.S_Rectangle)
            If oRectangle.mp_bInRect(X, Y) = True Then
                Return True
            End If
        Next
        Return False
    End Function

    Public ReadOnly Property Warning() As Boolean
        Get
            Return mp_bWarning
        End Get
    End Property

    Public Property WarningStyleIndex() As String
        Get
            Return mp_sWarningStyleIndex
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            mp_sWarningStyleIndex = Value
            If Value.Length > 0 Then
                mp_oWarningStyle = mp_oControl.Styles.FItem(Value)
            Else
                mp_oWarningStyle = Nothing
            End If
        End Set
    End Property

    Public ReadOnly Property WarningStyle() As clsStyle
        Get
            If mp_sWarningStyleIndex.Length = 0 Then
                Return mp_oStyle
            Else
                Return mp_oWarningStyle
            End If
        End Get
    End Property

    Public Property SelectedStyleIndex() As String
        Get
            Return mp_sSelectedStyleIndex
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            mp_sSelectedStyleIndex = Value
            If Value.Length > 0 Then
                mp_oSelectedStyle = mp_oControl.Styles.FItem(Value)
            Else
                mp_oSelectedStyle = Nothing
            End If
        End Set
    End Property

    Public ReadOnly Property SelectedStyle() As clsStyle
        Get
            If mp_sSelectedStyleIndex.Length = 0 Then
                Return mp_oStyle
            Else
                Return mp_oSelectedStyle
            End If
        End Get
    End Property

    Public Sub Check(ByVal Mode As E_PREDECESSORMODE)
        Dim dtPredecessor As AGVBE.DateTime = New AGVBE.DateTime()
        Dim dtSuccessor As AGVBE.DateTime = New AGVBE.DateTime()
        Dim lDiff As Integer
        Dim lDuration As Integer
        mp_bWarning = False
        Select Case mp_yPredecessorType
            Case E_CONSTRAINTTYPE.PCT_START_TO_START
                dtPredecessor = mp_oPredecessorTask.StartDate
                dtSuccessor = mp_oSuccessorTask.StartDate
                lDiff = mp_oControl.MathLib.DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor)
                If lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_FORCE Then
                    lDuration = mp_oControl.MathLib.DateTimeDiff(mp_oControl.CurrentViewObject.Interval, mp_oSuccessorTask.StartDate, mp_oSuccessorTask.EndDate)
                    mp_oSuccessorTask.StartDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask.StartDate)
                    mp_oSuccessorTask.EndDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, lDuration, mp_oSuccessorTask.StartDate)
                End If
            Case E_CONSTRAINTTYPE.PCT_END_TO_END
                dtPredecessor = mp_oPredecessorTask.EndDate
                dtSuccessor = mp_oSuccessorTask.EndDate
                lDiff = mp_oControl.MathLib.DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor)
                If lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_FORCE Then
                    lDuration = mp_oControl.MathLib.DateTimeDiff(mp_oControl.CurrentViewObject.Interval, mp_oSuccessorTask.StartDate, mp_oSuccessorTask.EndDate)
                    mp_oSuccessorTask.EndDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask.EndDate)
                    mp_oSuccessorTask.StartDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, -lDuration, mp_oSuccessorTask.EndDate)
                End If
            Case E_CONSTRAINTTYPE.PCT_START_TO_END
                dtPredecessor = mp_oPredecessorTask.StartDate
                dtSuccessor = mp_oSuccessorTask.EndDate
                lDiff = mp_oControl.MathLib.DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor)
                If lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_FORCE Then
                    lDuration = mp_oControl.MathLib.DateTimeDiff(mp_oControl.CurrentViewObject.Interval, mp_oSuccessorTask.StartDate, mp_oSuccessorTask.EndDate)
                    mp_oSuccessorTask.EndDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask.StartDate)
                    mp_oSuccessorTask.StartDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, -lDuration, mp_oSuccessorTask.EndDate)
                End If
            Case E_CONSTRAINTTYPE.PCT_END_TO_START
                dtPredecessor = mp_oPredecessorTask.EndDate
                dtSuccessor = mp_oSuccessorTask.StartDate
                lDiff = mp_oControl.MathLib.DateTimeDiff(mp_yLagInterval, dtPredecessor, dtSuccessor)
                If lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_FORCE Then
                    lDuration = mp_oControl.MathLib.DateTimeDiff(mp_oControl.CurrentViewObject.Interval, mp_oSuccessorTask.StartDate, mp_oSuccessorTask.EndDate)
                    mp_oSuccessorTask.StartDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, mp_lLagFactor, mp_oPredecessorTask.EndDate)
                    mp_oSuccessorTask.EndDate = mp_oControl.MathLib.DateTimeAdd(mp_yLagInterval, lDuration, mp_oSuccessorTask.StartDate)
                End If
        End Select
        If lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_CREATEWARNINGFLAG Then
            mp_bWarning = True
            mp_oSuccessorTask.mp_bWarning = True
        ElseIf lDiff <> mp_lLagFactor And Mode = E_PREDECESSORMODE.PM_RAISEEVENT Then
            mp_oControl.PredecessorExceptionEventArgs.Clear()
            mp_oControl.PredecessorExceptionEventArgs.PredecessorIndex = Me.Index
            mp_oControl.PredecessorExceptionEventArgs.PredecessorType = Me.PredecessorType
            mp_oControl.FirePredecessorException()
        End If
    End Sub
    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Predecessor")
        oXML.InitializeWriter()
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("SuccessorKey", mp_sSuccessorKey)
        oXML.WriteProperty("PredecessorKey", mp_sPredecessorKey)
        oXML.WriteProperty("PredecessorType", mp_yPredecessorType)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("LagInterval", mp_yLagInterval)
        oXML.WriteProperty("LagFactor", mp_lLagFactor)
        oXML.WriteProperty("WarningStyleIndex", mp_sWarningStyleIndex)
        oXML.WriteProperty("SelectedStyleIndex", mp_sSelectedStyleIndex)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Predecessor")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("SuccessorKey", mp_sSuccessorKey)
        mp_oSuccessorTask = mp_oControl.Tasks.Item(mp_sSuccessorKey)
        oXML.ReadProperty("PredecessorKey", mp_sPredecessorKey)
        mp_oPredecessorTask = mp_oControl.Tasks.Item(mp_sPredecessorKey)
        oXML.ReadProperty("PredecessorType", mp_yPredecessorType)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("LagInterval", mp_yLagInterval)
        oXML.ReadProperty("LagFactor", mp_lLagFactor)
        oXML.ReadProperty("WarningStyleIndex", mp_sWarningStyleIndex)
        WarningStyleIndex = mp_sWarningStyleIndex
        oXML.ReadProperty("SelectedStyleIndex", mp_sSelectedStyleIndex)
        SelectedStyleIndex = mp_sSelectedStyleIndex
    End Sub

End Class
