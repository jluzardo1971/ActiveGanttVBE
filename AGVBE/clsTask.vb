Option Explicit On

Public Class clsTask
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_bAllowStretchLeft As Boolean
    Private mp_bAllowStretchRight As Boolean
    Private mp_dtEndDate As AGVBE.DateTime
    Private mp_dtStartDate As AGVBE.DateTime
    Private mp_sText As String
    Private mp_sLayerIndex As String
    Private mp_oImage As Image
    Private mp_sRowKey As String
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_sTag As String
    Private mp_yAllowedMovement As E_MOVEMENTTYPE
    Private mp_bVisible As Boolean
    Friend mp_oRow As clsRow
    Private mp_oLayer As clsLayer
    Private mp_bIncomingPredecessors As Boolean
    Private mp_bOutgoingPredecessors As Boolean
    Private mp_sImageTag As String
    Friend mp_lTextLeft As Double
    Friend mp_lTextTop As Double
    Friend mp_lTextRight As Double
    Friend mp_lTextBottom As Double
    Private mp_bAllowTextEdit As Boolean
    Friend mp_bWarning As Boolean
    Private mp_sWarningStyleIndex As String
    Private mp_oWarningStyle As clsStyle
    Private mp_yTaskType As E_TASKTYPE
    Private mp_yDurationInterval As E_INTERVAL
    Private mp_lDurationFactor As Integer


    Public Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_bIncomingPredecessors = True
        mp_bOutgoingPredecessors = True
        mp_bAllowStretchLeft = True
        mp_bAllowStretchRight = True
        mp_dtEndDate = New AGVBE.DateTime()
        mp_dtEndDate.SetToCurrentDateTime()
        mp_dtStartDate = New AGVBE.DateTime()
        mp_dtStartDate.SetToCurrentDateTime()
        mp_sText = ""
        mp_sLayerIndex = "0"
        mp_oLayer = mp_oControl.Layers.FItem("0")
        mp_oImage = Nothing
        mp_sRowKey = ""
        mp_sStyleIndex = "DS_TASK"
        mp_oStyle = mp_oControl.Styles.FItem("DS_TASK")
        mp_sTag = ""
        mp_yAllowedMovement = E_MOVEMENTTYPE.MT_UNRESTRICTED
        mp_bVisible = True
        mp_sImageTag = ""
        mp_bAllowTextEdit = False
        mp_bWarning = False
        mp_sWarningStyleIndex = ""
        mp_yTaskType = E_TASKTYPE.TT_START_END
        mp_yDurationInterval = E_INTERVAL.IL_HOUR
        mp_lDurationFactor = 0
    End Sub

    Public Property AllowTextEdit() As Boolean
        Get
            Return mp_bAllowTextEdit
        End Get
        Set(ByVal value As Boolean)
            mp_bAllowTextEdit = value
        End Set
    End Property

    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.Tasks.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.TASKS_SET_KEY)
        End Set
    End Property

    
    Public Property IncomingPredecessors() As Boolean
        Get
            Return mp_bIncomingPredecessors
        End Get
        Set(ByVal Value As Boolean)
            mp_bIncomingPredecessors = Value
        End Set
    End Property

    
    Public Property OutgoingPredecessors() As Boolean
        Get
            Return mp_bOutgoingPredecessors
        End Get
        Set(ByVal Value As Boolean)
            mp_bOutgoingPredecessors = Value
        End Set
    End Property

    
    Public Property AllowStretchLeft() As Boolean
        Get
            Return mp_bAllowStretchLeft
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowStretchLeft = Value
        End Set
    End Property

    
    Public Property AllowStretchRight() As Boolean
        Get
            Return mp_bAllowStretchRight
        End Get
        Set(ByVal Value As Boolean)
            mp_bAllowStretchRight = Value
        End Set
    End Property

    
    Public Property Text() As String
        Get
            Return mp_sText
        End Get
        Set(ByVal Value As String)
            mp_sText = Value
        End Set
    End Property

    
    Public Property LayerIndex() As String
        Get
            Return mp_sLayerIndex
        End Get
        Set(ByVal Value As String)
            If (Value = "") Then
                Value = "0"
            End If
            mp_sLayerIndex = Value
            mp_oLayer = mp_oControl.Layers.FItem(Value)
        End Set
    End Property

    
    Public ReadOnly Property Layer() As clsLayer
        Get
            Return mp_oLayer
        End Get
    End Property

    
    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    
    Public Property RowKey() As String
        Get
            Return mp_sRowKey
        End Get
        Set(ByVal Value As String)
            If mp_oControl.Rows.oCollection.m_bDoesKeyExist(Value) = False Then
                mp_oControl.mp_ErrorReport(SYS_ERRORS.INVALID_ROW_KEY, "Invalid Row Key (""" & Value & """)", "ActiveGanttVBECtl.clsTask.Let RowKey")
                Exit Property
            End If
            mp_sRowKey = Value
            mp_oRow = mp_oControl.Rows.Item(mp_sRowKey)
        End Set
    End Property

    
    Public ReadOnly Property Row() As clsRow
        Get
            Return mp_oRow
        End Get
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_TASK" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_TASK"
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

    
    Public Property AllowedMovement() As E_MOVEMENTTYPE
        Get
            Return mp_yAllowedMovement
        End Get
        Set(ByVal Value As E_MOVEMENTTYPE)
            mp_yAllowedMovement = Value
        End Set
    End Property

    Public ReadOnly Property LeftTrim() As Integer
        Get
            If Left < mp_oControl.Splitter.Right Then
                Return mp_oControl.Splitter.Right
            Else
                Return Left
            End If
        End Get
    End Property

    
    Public ReadOnly Property RightTrim() As Integer
        Get
            If Right > mp_oControl.mt_RightMargin Then
                Return mp_oControl.mt_RightMargin
            Else
                Return Right
            End If
        End Get
    End Property

    Friend ReadOnly Property f_bLeftVisible() As Boolean
        Get
            If LeftTrim = Left Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Friend ReadOnly Property f_bRightVisible() As Boolean
        Get
            If RightTrim = Right Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    
    Public Property StartDate() As AGVBE.DateTime
        Get
            Return mp_dtStartDate
        End Get
        Set(ByVal Value As AGVBE.DateTime)
            mp_dtStartDate = Value
            If mp_yTaskType = E_TASKTYPE.TT_DURATION And mp_lDurationFactor <> 0 Then
                mp_GetDuration()
            End If
        End Set
    End Property

    Public ReadOnly Property Left() As Integer
        Get
            If mp_dtStartDate = mp_dtEndDate Then
                Return mp_oControl.MathLib.GetXCoordinateFromDate(StartDate) - mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset
            Else
                Return mp_oControl.MathLib.GetXCoordinateFromDate(StartDate)
            End If
        End Get
    End Property

    Public ReadOnly Property Right() As Integer
        Get
            If mp_dtStartDate = mp_dtEndDate Then
                Return mp_oControl.MathLib.GetXCoordinateFromDate(EndDate) + mp_oControl.CurrentViewObject.ClientArea.MilestoneSelectionOffset
            Else
                Return mp_oControl.MathLib.GetXCoordinateFromDate(EndDate)
            End If
        End Get
    End Property

    Public Property EndDate() As AGVBE.DateTime
        Get
            Return mp_dtEndDate
        End Get
        Set(ByVal Value As AGVBE.DateTime)
            If mp_yTaskType = E_TASKTYPE.TT_START_END Then
                mp_dtEndDate = Value
            End If
        End Set
    End Property

    
    Public ReadOnly Property Top() As Integer
        Get
            If (mp_oRow.Height <= -1) Then
                Return mp_oRow.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT Or mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_CELL Then
                Return mp_oRow.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT Then
                Return mp_oRow.Top + mp_oStyle.OffsetTop
            End If
            Return -1
        End Get
    End Property

    
    Public ReadOnly Property Bottom() As Integer
        Get
            If (mp_oRow.Height <= -1) Then
                Return mp_oRow.Top
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_ROWEXTENTSPLACEMENT Or mp_oStyle.Appearance = E_STYLEAPPEARANCE.SA_CELL Then
                Return mp_oRow.Bottom - 1
            End If
            If mp_oStyle.Placement = E_PLACEMENT.PLC_OFFSETPLACEMENT Then
                Return mp_oRow.Top + mp_oStyle.OffsetTop + mp_oStyle.OffsetBottom
            End If
            Return -1
        End Get
    End Property

    Public Property Visible() As Boolean
        Get
            If mp_oLayer.Visible = False Then
                Return False
            End If
            If (mp_oRow.Height <= -1) Then
                Return False
            End If
            If mp_oRow.Visible = False Then
                Return False
            End If
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    Friend ReadOnly Property InsideVisibleTimeLineArea() As Boolean
        Get
            If StartDate > mp_oControl.CurrentViewObject.TimeLine.EndDate Then
                Return False
            End If
            If EndDate < mp_oControl.CurrentViewObject.TimeLine.StartDate Then
                Return False
            End If
            Return True
        End Get
    End Property

    Friend ReadOnly Property ClientAreaVisiblity() As E_CLIENTAREAVISIBILITY
        Get
            If StartDate > mp_oControl.CurrentViewObject.TimeLine.EndDate Then
                Return E_CLIENTAREAVISIBILITY.VS_RIGHTOFVISIBLEAREA
            End If
            If EndDate < mp_oControl.CurrentViewObject.TimeLine.StartDate Then
                Return E_CLIENTAREAVISIBILITY.VS_LEFTOFVISIBLEAREA
            End If
            Return E_CLIENTAREAVISIBILITY.VS_INSIDEVISIBLEAREA
        End Get
    End Property

    
    Public ReadOnly Property Type() As E_OBJECTTYPE
        Get
            If mp_yTaskType = E_TASKTYPE.TT_DURATION And mp_lDurationFactor = 0 Then
                Return E_OBJECTTYPE.OT_MILESTONE
            End If
            If StartDate = EndDate Then
                Return E_OBJECTTYPE.OT_MILESTONE
            Else
                Return E_OBJECTTYPE.OT_TASK
            End If
        End Get
    End Property

    
    Public Function InConflict() As Boolean
        InConflict = mp_oControl.MathLib.DetectConflict(StartDate, EndDate, mp_sRowKey, Index, mp_sLayerIndex)
    End Function

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal Value As String)
            mp_sImageTag = Value
        End Set
    End Property

    Public ReadOnly Property Warning() As Boolean
        Get
            If mp_oControl.Predecessors.Count = 0 Then
                Return False
            Else
                Return mp_bWarning
            End If
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

    Public Property TaskType() As E_TASKTYPE
        Get
            Return mp_yTaskType
        End Get
        Set(ByVal value As E_TASKTYPE)
            mp_yTaskType = value
            If mp_yTaskType = E_TASKTYPE.TT_DURATION Then
                mp_GetDuration()
            End If
        End Set
    End Property

    Public Property DurationInterval() As E_INTERVAL
        Get
            Return mp_yDurationInterval
        End Get
        Set(ByVal value As E_INTERVAL)
            If Not (value = E_INTERVAL.IL_SECOND Or value = E_INTERVAL.IL_MINUTE Or value = E_INTERVAL.IL_HOUR Or value = E_INTERVAL.IL_DAY) Then
                mp_oControl.mp_ErrorReport(SYS_ERRORS.INVALID_DURATION_INTERVAL, "Interval is invalid for a duration", "clsTask.Set_DurationInterval")
                Return
            End If
            mp_yDurationInterval = value
            If mp_yTaskType = E_TASKTYPE.TT_DURATION Then
                mp_GetDuration()
            End If
        End Set
    End Property

    Public Property DurationFactor() As Integer
        Get
            Return mp_lDurationFactor
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then
                value = value * -1
            End If
            mp_lDurationFactor = value
            If mp_yTaskType = E_TASKTYPE.TT_DURATION Then
                mp_GetDuration()
            End If
        End Set
    End Property

    Private Sub mp_GetDuration()
        mp_dtEndDate = mp_oControl.MathLib.GetEndDate(mp_dtStartDate, mp_yDurationInterval, mp_lDurationFactor)
    End Sub

    
    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "Task")
        oXML.InitializeWriter()
        oXML.WriteProperty("AllowedMovement", mp_yAllowedMovement)
        oXML.WriteProperty("AllowStretchLeft", mp_bAllowStretchLeft)
        oXML.WriteProperty("AllowStretchRight", mp_bAllowStretchRight)
        oXML.WriteProperty("EndDate", mp_dtEndDate)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("IncomingPredecessors", mp_bIncomingPredecessors)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("LayerIndex", mp_sLayerIndex)
        oXML.WriteProperty("OutgoingPredecessors", mp_bOutgoingPredecessors)
        oXML.WriteProperty("RowKey", mp_sRowKey)
        oXML.WriteProperty("StartDate", mp_dtStartDate)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("Text", mp_sText)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteProperty("AllowTextEdit", mp_bAllowTextEdit)
        oXML.WriteProperty("WarningStyleIndex", mp_sWarningStyleIndex)
        oXML.WriteProperty("TaskType", mp_yTaskType)
        oXML.WriteProperty("DurationInterval", mp_yDurationInterval)
        oXML.WriteProperty("DurationFactor", mp_lDurationFactor)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "Task")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("AllowedMovement", mp_yAllowedMovement)
        oXML.ReadProperty("AllowStretchLeft", mp_bAllowStretchLeft)
        oXML.ReadProperty("AllowStretchRight", mp_bAllowStretchRight)
        oXML.ReadProperty("EndDate", mp_dtEndDate)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("IncomingPredecessors", mp_bIncomingPredecessors)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("LayerIndex", mp_sLayerIndex)
        mp_oLayer = mp_oControl.Layers.FItem(mp_sLayerIndex)
        oXML.ReadProperty("OutgoingPredecessors", mp_bOutgoingPredecessors)
        oXML.ReadProperty("RowKey", mp_sRowKey)
        mp_oRow = mp_oControl.Rows.Item(mp_sRowKey)
        oXML.ReadProperty("StartDate", mp_dtStartDate)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("Text", mp_sText)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        oXML.ReadProperty("AllowTextEdit", mp_bAllowTextEdit)
        oXML.ReadProperty("WarningStyleIndex", mp_sWarningStyleIndex)
        WarningStyleIndex = mp_sWarningStyleIndex
        oXML.ReadProperty("TaskType", mp_yTaskType)
        oXML.ReadProperty("DurationInterval", mp_yDurationInterval)
        oXML.ReadProperty("DurationFactor", mp_lDurationFactor)
    End Sub




End Class
