Option Explicit On 

Public Class clsTimeBlock
    Inherits clsItemBase

    Private mp_oControl As ActiveGanttVBECtl
    Private mp_sStyleIndex As String
    Private mp_sTag As String
    Private mp_bGenerateConflict As Boolean
    Private mp_bVisible As Boolean
    Private mp_oStyle As clsStyle
    Private mp_yTimeBlockType As E_TIMEBLOCKTYPE
    Private mp_yRecurringType As E_RECURRINGTYPE
    Private mp_yBaseWeekDay As E_WEEKDAY
    Private mp_dtBaseDate As AGVBE.DateTime
    Private mp_yDurationInterval As E_INTERVAL
    Private mp_lDurationFactor As Integer
    Private mp_bNonWorking As Boolean

    Friend Sub New(ByVal Value As ActiveGanttVBECtl)
        mp_oControl = Value
        mp_sStyleIndex = "DS_TIMEBLOCK"
        mp_oStyle = mp_oControl.Styles.FItem("DS_TIMEBLOCK")
        mp_sTag = ""
        mp_bGenerateConflict = False
        mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE
        mp_yRecurringType = E_RECURRINGTYPE.RCT_DAY
        mp_bVisible = True
        mp_yBaseWeekDay = E_WEEKDAY.WD_SUNDAY
        mp_dtBaseDate = New AGVBE.DateTime()
        mp_dtBaseDate.SetToCurrentDateTime()
        mp_yDurationInterval = E_INTERVAL.IL_HOUR
        mp_lDurationFactor = 1
        mp_bNonWorking = False
    End Sub

    
    Public Property Key() As String
        Get
            Return mp_sKey
        End Get
        Set(ByVal Value As String)
            mp_oControl.TimeBlocks.oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.TIMEBLOCKS_SET_KEY)
        End Set
    End Property

    
    Public Property TimeBlockType() As E_TIMEBLOCKTYPE
        Get
            Return mp_yTimeBlockType
        End Get
        Set(ByVal Value As E_TIMEBLOCKTYPE)
            mp_yTimeBlockType = Value
        End Set
    End Property

    
    Public Property RecurringType() As E_RECURRINGTYPE
        Get
            Return mp_yRecurringType
        End Get
        Set(ByVal Value As E_RECURRINGTYPE)
            mp_yRecurringType = Value
        End Set
    End Property
    
    Public ReadOnly Property EndDate() As AGVBE.DateTime
        Get
            If mp_lDurationFactor > 0 Then
                Return mp_oControl.MathLib.DateTimeAdd(mp_yDurationInterval, mp_lDurationFactor, mp_dtBaseDate)
            Else
                Return mp_dtBaseDate
            End If
        End Get
    End Property

    Public ReadOnly Property StartDate() As AGVBE.DateTime
        Get
            If mp_lDurationFactor > 0 Then
                Return mp_dtBaseDate
            Else
                Return mp_oControl.MathLib.DateTimeAdd(mp_yDurationInterval, mp_lDurationFactor, mp_dtBaseDate)
            End If
        End Get
    End Property

    
    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_TIMEBLOCK" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            If Trim(Value) = "" Then Value = "DS_TIMEBLOCK"
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

    
    Public Property GenerateConflict() As Boolean
        Get
            Return mp_bGenerateConflict
        End Get
        Set(ByVal Value As Boolean)
            mp_bGenerateConflict = Value
        End Set
    End Property

    Public ReadOnly Property LeftTrim() As Integer
        Get
            If mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE Then
                If Left < mp_oControl.Splitter.Right Then
                    Return mp_oControl.Splitter.Right
                Else
                    Return Left
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    
    Public ReadOnly Property RightTrim() As Integer
        Get
            If mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE Then
                If Right > mp_oControl.mt_RightMargin Then
                    Return mp_oControl.mt_RightMargin
                Else
                    Return Right
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    
    Public ReadOnly Property Left() As Integer
        Get
            If mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE Then
                Return mp_oControl.MathLib.GetXCoordinateFromDate(StartDate)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Top() As Integer
        Get
            Return mp_oControl.CurrentViewObject.ClientArea.Top
        End Get
    End Property

    Public ReadOnly Property Right() As Integer
        Get
            If mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_SINGLE_OCURRENCE Then
                Return mp_oControl.MathLib.GetXCoordinateFromDate(EndDate)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Bottom() As Integer
        Get
            If mp_oControl.TimeBlockBehaviour = E_TIMEBLOCKBEHAVIOUR.TBB_CONTROLEXTENTS Then
                Return mp_oControl.mt_BottomMargin
            ElseIf mp_oControl.Rows.Count > 0 Then
                Return mp_oControl.Rows.Item(mp_oControl.CurrentViewObject.ClientArea.LastVisibleRow.ToString()).Bottom
            Else
                Return mp_oControl.CurrentViewObject.ClientArea.Top
            End If
        End Get
    End Property

    Friend Property f_Visible() As Boolean
        Get
            Return mp_bVisible
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    Public Property Visible() As Boolean
        Get
            If mp_oControl.Rows.Count = 0 Then
                Return False
            End If
            If mp_yTimeBlockType = E_TIMEBLOCKTYPE.TBT_RECURRING Then
                Return False
            End If
            Dim dtStartDate As AGVBE.DateTime
            Dim dtEndDate As AGVBE.DateTime
            dtStartDate = StartDate
            dtEndDate = EndDate
            If (((dtStartDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And dtStartDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate) Or (dtEndDate >= mp_oControl.CurrentViewObject.TimeLine.StartDate And dtEndDate <= mp_oControl.CurrentViewObject.TimeLine.EndDate)) Or (dtStartDate < mp_oControl.CurrentViewObject.TimeLine.StartDate And dtEndDate > mp_oControl.CurrentViewObject.TimeLine.EndDate)) Then
                Return mp_bVisible
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            mp_bVisible = Value
        End Set
    End Property

    Public Property BaseDate() As AGVBE.DateTime
        Get
            Return mp_dtBaseDate
        End Get
        Set(ByVal value As AGVBE.DateTime)
            mp_dtBaseDate = value
        End Set
    End Property

    Public Property BaseWeekDay() As E_WEEKDAY
        Get
            Return mp_yBaseWeekDay
        End Get
        Set(ByVal value As E_WEEKDAY)
            mp_yBaseWeekDay = value
        End Set
    End Property

    Public Property DurationInterval() As E_INTERVAL
        Get
            Return mp_yDurationInterval
        End Get
        Set(ByVal value As E_INTERVAL)
            mp_yDurationInterval = value
        End Set
    End Property

    Public Property DurationFactor() As Integer
        Get
            Return mp_lDurationFactor
        End Get
        Set(ByVal value As Integer)
            mp_lDurationFactor = value
        End Set
    End Property

    Public Property NonWorking() As Boolean
        Get
            Return mp_bNonWorking
        End Get
        Set(ByVal value As Boolean)
            mp_bNonWorking = value
        End Set
    End Property

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TimeBlock")
        oXML.InitializeWriter()
        oXML.WriteProperty("GenerateConflict", mp_bGenerateConflict)
        oXML.WriteProperty("Key", mp_sKey)
        oXML.WriteProperty("RecurringType", mp_yRecurringType)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Tag", mp_sTag)
        oXML.WriteProperty("TimeBlockType", mp_yTimeBlockType)
        oXML.WriteProperty("Visible", mp_bVisible)
        oXML.WriteProperty("BaseDate", mp_dtBaseDate)
        oXML.WriteProperty("BaseWeekDay", mp_yBaseWeekDay)
        oXML.WriteProperty("DurationInterval", mp_yDurationInterval)
        oXML.WriteProperty("DurationFactor", mp_lDurationFactor)
        oXML.WriteProperty("NonWorking", mp_bNonWorking)
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TimeBlock")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        oXML.ReadProperty("GenerateConflict", mp_bGenerateConflict)
        oXML.ReadProperty("Key", mp_sKey)
        oXML.ReadProperty("RecurringType", mp_yRecurringType)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        StyleIndex = mp_sStyleIndex
        oXML.ReadProperty("Tag", mp_sTag)
        oXML.ReadProperty("TimeBlockType", mp_yTimeBlockType)
        oXML.ReadProperty("Visible", mp_bVisible)
        oXML.ReadProperty("BaseDate", mp_dtBaseDate)
        oXML.ReadProperty("BaseWeekDay", mp_yBaseWeekDay)
        oXML.ReadProperty("DurationInterval", mp_yDurationInterval)
        oXML.ReadProperty("DurationFactor", mp_lDurationFactor)
        oXML.ReadProperty("NonWorking", mp_bNonWorking)
    End Sub

End Class
