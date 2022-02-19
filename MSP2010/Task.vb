Option Explicit On

Public Class Task
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_lID As Integer
	Private mp_sName As String
	Private mp_yType As E_TYPE_4
	Private mp_bIsNull As Boolean
	Private mp_dtCreateDate As System.DateTime
	Private mp_sContact As String
	Private mp_sWBS As String
	Private mp_sWBSLevel As String
	Private mp_sOutlineNumber As String
	Private mp_lOutlineLevel As Integer
	Private mp_lPriority As Integer
	Private mp_dtStart As System.DateTime
	Private mp_dtFinish As System.DateTime
	Private mp_oDuration As Duration
	Private mp_yDurationFormat As E_DURATIONFORMAT
	Private mp_oWork As Duration
	Private mp_dtStop As System.DateTime
	Private mp_dtResume As System.DateTime
	Private mp_bResumeValid As Boolean
	Private mp_bEffortDriven As Boolean
	Private mp_bRecurring As Boolean
	Private mp_bOverAllocated As Boolean
	Private mp_bEstimated As Boolean
	Private mp_bMilestone As Boolean
	Private mp_bSummary As Boolean
	Private mp_bDisplayAsSummary As Boolean
	Private mp_bCritical As Boolean
	Private mp_bIsSubproject As Boolean
	Private mp_bIsSubprojectReadOnly As Boolean
	Private mp_sSubprojectName As String
	Private mp_bExternalTask As Boolean
	Private mp_sExternalTaskProject As String
	Private mp_dtEarlyStart As System.DateTime
	Private mp_dtEarlyFinish As System.DateTime
	Private mp_dtLateStart As System.DateTime
	Private mp_dtLateFinish As System.DateTime
	Private mp_lStartVariance As Integer
	Private mp_lFinishVariance As Integer
	Private mp_fWorkVariance As Single
	Private mp_lFreeSlack As Integer
	Private mp_lStartSlack As Integer
	Private mp_lFinishSlack As Integer
	Private mp_lTotalSlack As Integer
	Private mp_fFixedCost As Single
	Private mp_yFixedCostAccrual As E_FIXEDCOSTACCRUAL
	Private mp_lPercentComplete As Integer
	Private mp_lPercentWorkComplete As Integer
	Private mp_cCost As Decimal
	Private mp_cOvertimeCost As Decimal
	Private mp_oOvertimeWork As Duration
	Private mp_dtActualStart As System.DateTime
	Private mp_dtActualFinish As System.DateTime
	Private mp_oActualDuration As Duration
	Private mp_cActualCost As Decimal
	Private mp_cActualOvertimeCost As Decimal
	Private mp_oActualWork As Duration
	Private mp_oActualOvertimeWork As Duration
	Private mp_oRegularWork As Duration
	Private mp_oRemainingDuration As Duration
	Private mp_cRemainingCost As Decimal
	Private mp_oRemainingWork As Duration
	Private mp_cRemainingOvertimeCost As Decimal
	Private mp_oRemainingOvertimeWork As Duration
	Private mp_fACWP As Single
	Private mp_fCV As Single
	Private mp_yConstraintType As E_CONSTRAINTTYPE
	Private mp_lCalendarUID As Integer
	Private mp_dtConstraintDate As System.DateTime
	Private mp_dtDeadline As System.DateTime
	Private mp_bLevelAssignments As Boolean
	Private mp_bLevelingCanSplit As Boolean
	Private mp_lLevelingDelay As Integer
	Private mp_yLevelingDelayFormat As E_LEVELINGDELAYFORMAT
	Private mp_dtPreLeveledStart As System.DateTime
	Private mp_dtPreLeveledFinish As System.DateTime
	Private mp_sHyperlink As String
	Private mp_sHyperlinkAddress As String
	Private mp_sHyperlinkSubAddress As String
	Private mp_bIgnoreResourceCalendar As Boolean
	Private mp_sNotes As String
	Private mp_bHideBar As Boolean
	Private mp_bRollup As Boolean
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single
	Private mp_lPhysicalPercentComplete As Integer
	Private mp_yEarnedValueMethod As E_EARNEDVALUEMETHOD
	Private mp_oPredecessorLink_C As TaskPredecessorLink_C
	Private mp_oActualWorkProtected As Duration
	Private mp_oActualOvertimeWorkProtected As Duration
	Private mp_oExtendedAttribute_C As TaskExtendedAttribute_C
	Private mp_oBaseline_C As TaskBaseline_C
	Private mp_oOutlineCode_C As TaskOutlineCode_C
	Private mp_bIsPublished As Boolean
	Private mp_sStatusManager As String
	Private mp_dtCommitmentStart As System.DateTime
	Private mp_dtCommitmentFinish As System.DateTime
	Private mp_yCommitmentType As E_COMMITMENTTYPE
	Private mp_bActive As Boolean
	Private mp_bPinned As Boolean
	Private mp_sPinnedStart As String
	Private mp_sPinnedFinish As String
	Private mp_sPinnedDuration As String
	Private mp_oTimephasedData_C As TimephasedData_C

	Public Sub New()
		mp_lUID = 0
		mp_lID = 0
		mp_sName = ""
		mp_yType = E_TYPE_4.T_4_FIXED_UNITS
		mp_bIsNull = False
		mp_dtCreateDate = New System.DateTime(0)
		mp_sContact = ""
		mp_sWBS = ""
		mp_sWBSLevel = ""
		mp_sOutlineNumber = ""
		mp_lOutlineLevel = 0
		mp_lPriority = 0
		mp_dtStart = New System.DateTime(0)
		mp_dtFinish = New System.DateTime(0)
		mp_oDuration = New Duration()
		mp_yDurationFormat = E_DURATIONFORMAT.DF_M
		mp_oWork = New Duration()
		mp_dtStop = New System.DateTime(0)
		mp_dtResume = New System.DateTime(0)
		mp_bResumeValid = False
		mp_bEffortDriven = False
		mp_bRecurring = False
		mp_bOverAllocated = False
		mp_bEstimated = False
		mp_bMilestone = False
		mp_bSummary = False
		mp_bDisplayAsSummary = False
		mp_bCritical = False
		mp_bIsSubproject = False
		mp_bIsSubprojectReadOnly = False
		mp_sSubprojectName = ""
		mp_bExternalTask = False
		mp_sExternalTaskProject = ""
		mp_dtEarlyStart = New System.DateTime(0)
		mp_dtEarlyFinish = New System.DateTime(0)
		mp_dtLateStart = New System.DateTime(0)
		mp_dtLateFinish = New System.DateTime(0)
		mp_lStartVariance = 0
		mp_lFinishVariance = 0
		mp_fWorkVariance = 0
		mp_lFreeSlack = 0
		mp_lStartSlack = 0
		mp_lFinishSlack = 0
		mp_lTotalSlack = 0
		mp_fFixedCost = 0
		mp_yFixedCostAccrual = E_FIXEDCOSTACCRUAL.FCA_START
		mp_lPercentComplete = 0
		mp_lPercentWorkComplete = 0
		mp_cCost = 0
		mp_cOvertimeCost = 0
		mp_oOvertimeWork = New Duration()
		mp_dtActualStart = New System.DateTime(0)
		mp_dtActualFinish = New System.DateTime(0)
		mp_oActualDuration = New Duration()
		mp_cActualCost = 0
		mp_cActualOvertimeCost = 0
		mp_oActualWork = New Duration()
		mp_oActualOvertimeWork = New Duration()
		mp_oRegularWork = New Duration()
		mp_oRemainingDuration = New Duration()
		mp_cRemainingCost = 0
		mp_oRemainingWork = New Duration()
		mp_cRemainingOvertimeCost = 0
		mp_oRemainingOvertimeWork = New Duration()
		mp_fACWP = 0
		mp_fCV = 0
		mp_yConstraintType = E_CONSTRAINTTYPE.CT_AS_SOON_AS_POSSIBLE
		mp_lCalendarUID = 0
		mp_dtConstraintDate = New System.DateTime(0)
		mp_dtDeadline = New System.DateTime(0)
		mp_bLevelAssignments = False
		mp_bLevelingCanSplit = False
		mp_lLevelingDelay = 0
		mp_yLevelingDelayFormat = E_LEVELINGDELAYFORMAT.LDF_M
		mp_dtPreLeveledStart = New System.DateTime(0)
		mp_dtPreLeveledFinish = New System.DateTime(0)
		mp_sHyperlink = ""
		mp_sHyperlinkAddress = ""
		mp_sHyperlinkSubAddress = ""
		mp_bIgnoreResourceCalendar = False
		mp_sNotes = ""
		mp_bHideBar = False
		mp_bRollup = False
		mp_fBCWS = 0
		mp_fBCWP = 0
		mp_lPhysicalPercentComplete = 0
		mp_yEarnedValueMethod = E_EARNEDVALUEMETHOD.EVM_PERCENT_COMPLETE
		mp_oPredecessorLink_C = New TaskPredecessorLink_C()
		mp_oActualWorkProtected = New Duration()
		mp_oActualOvertimeWorkProtected = New Duration()
		mp_oExtendedAttribute_C = New TaskExtendedAttribute_C()
		mp_oBaseline_C = New TaskBaseline_C()
		mp_oOutlineCode_C = New TaskOutlineCode_C()
		mp_bIsPublished = False
		mp_sStatusManager = ""
		mp_dtCommitmentStart = New System.DateTime(0)
		mp_dtCommitmentFinish = New System.DateTime(0)
		mp_yCommitmentType = E_COMMITMENTTYPE.CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE
		mp_bActive = False
		mp_bPinned = False
		mp_sPinnedStart = ""
		mp_sPinnedFinish = ""
		mp_sPinnedDuration = ""
		mp_oTimephasedData_C = New TimephasedData_C()
	End Sub

	Public Property lUID() As Integer
		Get
			Return mp_lUID
		End Get
		Set(ByVal Value As Integer)
			mp_lUID = Value
		End Set
	End Property

	Public Property lID() As Integer
		Get
			Return mp_lID
		End Get
		Set(ByVal Value As Integer)
			mp_lID = Value
		End Set
	End Property

	Public Property sName() As String
		Get
			Return mp_sName
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sName = Value
		End Set
	End Property

	Public Property yType() As E_TYPE_4
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE_4)
			mp_yType = Value
		End Set
	End Property

	Public Property bIsNull() As Boolean
		Get
			Return mp_bIsNull
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsNull = Value
		End Set
	End Property

	Public Property dtCreateDate() As System.DateTime
		Get
			Return mp_dtCreateDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCreateDate = Value
		End Set
	End Property

	Public Property sContact() As String
		Get
			Return mp_sContact
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sContact = Value
		End Set
	End Property

	Public Property sWBS() As String
		Get
			Return mp_sWBS
		End Get
		Set(ByVal Value As String)
			mp_sWBS = Value
		End Set
	End Property

	Public Property sWBSLevel() As String
		Get
			Return mp_sWBSLevel
		End Get
		Set(ByVal Value As String)
			mp_sWBSLevel = Value
		End Set
	End Property

	Public Property sOutlineNumber() As String
		Get
			Return mp_sOutlineNumber
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sOutlineNumber = Value
		End Set
	End Property

	Public Property lOutlineLevel() As Integer
		Get
			Return mp_lOutlineLevel
		End Get
		Set(ByVal Value As Integer)
			mp_lOutlineLevel = Value
		End Set
	End Property

	Public Property lPriority() As Integer
		Get
			Return mp_lPriority
		End Get
		Set(ByVal Value As Integer)
			mp_lPriority = Value
		End Set
	End Property

	Public Property dtStart() As System.DateTime
		Get
			Return mp_dtStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtStart = Value
		End Set
	End Property

	Public Property dtFinish() As System.DateTime
		Get
			Return mp_dtFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtFinish = Value
		End Set
	End Property

	Public ReadOnly Property oDuration() As Duration
		Get
			Return mp_oDuration
		End Get
	End Property

	Public Property yDurationFormat() As E_DURATIONFORMAT
		Get
			Return mp_yDurationFormat
		End Get
		Set(ByVal Value As E_DURATIONFORMAT)
			mp_yDurationFormat = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public Property dtStop() As System.DateTime
		Get
			Return mp_dtStop
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtStop = Value
		End Set
	End Property

	Public Property dtResume() As System.DateTime
		Get
			Return mp_dtResume
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtResume = Value
		End Set
	End Property

	Public Property bResumeValid() As Boolean
		Get
			Return mp_bResumeValid
		End Get
		Set(ByVal Value As Boolean)
			mp_bResumeValid = Value
		End Set
	End Property

	Public Property bEffortDriven() As Boolean
		Get
			Return mp_bEffortDriven
		End Get
		Set(ByVal Value As Boolean)
			mp_bEffortDriven = Value
		End Set
	End Property

	Public Property bRecurring() As Boolean
		Get
			Return mp_bRecurring
		End Get
		Set(ByVal Value As Boolean)
			mp_bRecurring = Value
		End Set
	End Property

	Public Property bOverAllocated() As Boolean
		Get
			Return mp_bOverAllocated
		End Get
		Set(ByVal Value As Boolean)
			mp_bOverAllocated = Value
		End Set
	End Property

	Public Property bEstimated() As Boolean
		Get
			Return mp_bEstimated
		End Get
		Set(ByVal Value As Boolean)
			mp_bEstimated = Value
		End Set
	End Property

	Public Property bMilestone() As Boolean
		Get
			Return mp_bMilestone
		End Get
		Set(ByVal Value As Boolean)
			mp_bMilestone = Value
		End Set
	End Property

	Public Property bSummary() As Boolean
		Get
			Return mp_bSummary
		End Get
		Set(ByVal Value As Boolean)
			mp_bSummary = Value
		End Set
	End Property

	Public Property bDisplayAsSummary() As Boolean
		Get
			Return mp_bDisplayAsSummary
		End Get
		Set(ByVal Value As Boolean)
			mp_bDisplayAsSummary = Value
		End Set
	End Property

	Public Property bCritical() As Boolean
		Get
			Return mp_bCritical
		End Get
		Set(ByVal Value As Boolean)
			mp_bCritical = Value
		End Set
	End Property

	Public Property bIsSubproject() As Boolean
		Get
			Return mp_bIsSubproject
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsSubproject = Value
		End Set
	End Property

	Public Property bIsSubprojectReadOnly() As Boolean
		Get
			Return mp_bIsSubprojectReadOnly
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsSubprojectReadOnly = Value
		End Set
	End Property

	Public Property sSubprojectName() As String
		Get
			Return mp_sSubprojectName
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sSubprojectName = Value
		End Set
	End Property

	Public Property bExternalTask() As Boolean
		Get
			Return mp_bExternalTask
		End Get
		Set(ByVal Value As Boolean)
			mp_bExternalTask = Value
		End Set
	End Property

	Public Property sExternalTaskProject() As String
		Get
			Return mp_sExternalTaskProject
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sExternalTaskProject = Value
		End Set
	End Property

	Public Property dtEarlyStart() As System.DateTime
		Get
			Return mp_dtEarlyStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtEarlyStart = Value
		End Set
	End Property

	Public Property dtEarlyFinish() As System.DateTime
		Get
			Return mp_dtEarlyFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtEarlyFinish = Value
		End Set
	End Property

	Public Property dtLateStart() As System.DateTime
		Get
			Return mp_dtLateStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtLateStart = Value
		End Set
	End Property

	Public Property dtLateFinish() As System.DateTime
		Get
			Return mp_dtLateFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtLateFinish = Value
		End Set
	End Property

	Public Property lStartVariance() As Integer
		Get
			Return mp_lStartVariance
		End Get
		Set(ByVal Value As Integer)
			mp_lStartVariance = Value
		End Set
	End Property

	Public Property lFinishVariance() As Integer
		Get
			Return mp_lFinishVariance
		End Get
		Set(ByVal Value As Integer)
			mp_lFinishVariance = Value
		End Set
	End Property

	Public Property fWorkVariance() As Single
		Get
			Return mp_fWorkVariance
		End Get
		Set(ByVal Value As Single)
			mp_fWorkVariance = Value
		End Set
	End Property

	Public Property lFreeSlack() As Integer
		Get
			Return mp_lFreeSlack
		End Get
		Set(ByVal Value As Integer)
			mp_lFreeSlack = Value
		End Set
	End Property

	Public Property lStartSlack() As Integer
		Get
			Return mp_lStartSlack
		End Get
		Set(ByVal Value As Integer)
			mp_lStartSlack = Value
		End Set
	End Property

	Public Property lFinishSlack() As Integer
		Get
			Return mp_lFinishSlack
		End Get
		Set(ByVal Value As Integer)
			mp_lFinishSlack = Value
		End Set
	End Property

	Public Property lTotalSlack() As Integer
		Get
			Return mp_lTotalSlack
		End Get
		Set(ByVal Value As Integer)
			mp_lTotalSlack = Value
		End Set
	End Property

	Public Property fFixedCost() As Single
		Get
			Return mp_fFixedCost
		End Get
		Set(ByVal Value As Single)
			mp_fFixedCost = Value
		End Set
	End Property

	Public Property yFixedCostAccrual() As E_FIXEDCOSTACCRUAL
		Get
			Return mp_yFixedCostAccrual
		End Get
		Set(ByVal Value As E_FIXEDCOSTACCRUAL)
			mp_yFixedCostAccrual = Value
		End Set
	End Property

	Public Property lPercentComplete() As Integer
		Get
			Return mp_lPercentComplete
		End Get
		Set(ByVal Value As Integer)
			mp_lPercentComplete = Value
		End Set
	End Property

	Public Property lPercentWorkComplete() As Integer
		Get
			Return mp_lPercentWorkComplete
		End Get
		Set(ByVal Value As Integer)
			mp_lPercentWorkComplete = Value
		End Set
	End Property

	Public Property cCost() As Decimal
		Get
			Return mp_cCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cCost = Value
		End Set
	End Property

	Public Property cOvertimeCost() As Decimal
		Get
			Return mp_cOvertimeCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cOvertimeCost = Value
		End Set
	End Property

	Public ReadOnly Property oOvertimeWork() As Duration
		Get
			Return mp_oOvertimeWork
		End Get
	End Property

	Public Property dtActualStart() As System.DateTime
		Get
			Return mp_dtActualStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtActualStart = Value
		End Set
	End Property

	Public Property dtActualFinish() As System.DateTime
		Get
			Return mp_dtActualFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtActualFinish = Value
		End Set
	End Property

	Public ReadOnly Property oActualDuration() As Duration
		Get
			Return mp_oActualDuration
		End Get
	End Property

	Public Property cActualCost() As Decimal
		Get
			Return mp_cActualCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cActualCost = Value
		End Set
	End Property

	Public Property cActualOvertimeCost() As Decimal
		Get
			Return mp_cActualOvertimeCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cActualOvertimeCost = Value
		End Set
	End Property

	Public ReadOnly Property oActualWork() As Duration
		Get
			Return mp_oActualWork
		End Get
	End Property

	Public ReadOnly Property oActualOvertimeWork() As Duration
		Get
			Return mp_oActualOvertimeWork
		End Get
	End Property

	Public ReadOnly Property oRegularWork() As Duration
		Get
			Return mp_oRegularWork
		End Get
	End Property

	Public ReadOnly Property oRemainingDuration() As Duration
		Get
			Return mp_oRemainingDuration
		End Get
	End Property

	Public Property cRemainingCost() As Decimal
		Get
			Return mp_cRemainingCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cRemainingCost = Value
		End Set
	End Property

	Public ReadOnly Property oRemainingWork() As Duration
		Get
			Return mp_oRemainingWork
		End Get
	End Property

	Public Property cRemainingOvertimeCost() As Decimal
		Get
			Return mp_cRemainingOvertimeCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cRemainingOvertimeCost = Value
		End Set
	End Property

	Public ReadOnly Property oRemainingOvertimeWork() As Duration
		Get
			Return mp_oRemainingOvertimeWork
		End Get
	End Property

	Public Property fACWP() As Single
		Get
			Return mp_fACWP
		End Get
		Set(ByVal Value As Single)
			mp_fACWP = Value
		End Set
	End Property

	Public Property fCV() As Single
		Get
			Return mp_fCV
		End Get
		Set(ByVal Value As Single)
			mp_fCV = Value
		End Set
	End Property

	Public Property yConstraintType() As E_CONSTRAINTTYPE
		Get
			Return mp_yConstraintType
		End Get
		Set(ByVal Value As E_CONSTRAINTTYPE)
			mp_yConstraintType = Value
		End Set
	End Property

	Public Property lCalendarUID() As Integer
		Get
			Return mp_lCalendarUID
		End Get
		Set(ByVal Value As Integer)
			mp_lCalendarUID = Value
		End Set
	End Property

	Public Property dtConstraintDate() As System.DateTime
		Get
			Return mp_dtConstraintDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtConstraintDate = Value
		End Set
	End Property

	Public Property dtDeadline() As System.DateTime
		Get
			Return mp_dtDeadline
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtDeadline = Value
		End Set
	End Property

	Public Property bLevelAssignments() As Boolean
		Get
			Return mp_bLevelAssignments
		End Get
		Set(ByVal Value As Boolean)
			mp_bLevelAssignments = Value
		End Set
	End Property

	Public Property bLevelingCanSplit() As Boolean
		Get
			Return mp_bLevelingCanSplit
		End Get
		Set(ByVal Value As Boolean)
			mp_bLevelingCanSplit = Value
		End Set
	End Property

	Public Property lLevelingDelay() As Integer
		Get
			Return mp_lLevelingDelay
		End Get
		Set(ByVal Value As Integer)
			mp_lLevelingDelay = Value
		End Set
	End Property

	Public Property yLevelingDelayFormat() As E_LEVELINGDELAYFORMAT
		Get
			Return mp_yLevelingDelayFormat
		End Get
		Set(ByVal Value As E_LEVELINGDELAYFORMAT)
			mp_yLevelingDelayFormat = Value
		End Set
	End Property

	Public Property dtPreLeveledStart() As System.DateTime
		Get
			Return mp_dtPreLeveledStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtPreLeveledStart = Value
		End Set
	End Property

	Public Property dtPreLeveledFinish() As System.DateTime
		Get
			Return mp_dtPreLeveledFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtPreLeveledFinish = Value
		End Set
	End Property

	Public Property sHyperlink() As String
		Get
			Return mp_sHyperlink
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sHyperlink = Value
		End Set
	End Property

	Public Property sHyperlinkAddress() As String
		Get
			Return mp_sHyperlinkAddress
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sHyperlinkAddress = Value
		End Set
	End Property

	Public Property sHyperlinkSubAddress() As String
		Get
			Return mp_sHyperlinkSubAddress
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sHyperlinkSubAddress = Value
		End Set
	End Property

	Public Property bIgnoreResourceCalendar() As Boolean
		Get
			Return mp_bIgnoreResourceCalendar
		End Get
		Set(ByVal Value As Boolean)
			mp_bIgnoreResourceCalendar = Value
		End Set
	End Property

	Public Property sNotes() As String
		Get
			Return mp_sNotes
		End Get
		Set(ByVal Value As String)
			mp_sNotes = Value
		End Set
	End Property

	Public Property bHideBar() As Boolean
		Get
			Return mp_bHideBar
		End Get
		Set(ByVal Value As Boolean)
			mp_bHideBar = Value
		End Set
	End Property

	Public Property bRollup() As Boolean
		Get
			Return mp_bRollup
		End Get
		Set(ByVal Value As Boolean)
			mp_bRollup = Value
		End Set
	End Property

	Public Property fBCWS() As Single
		Get
			Return mp_fBCWS
		End Get
		Set(ByVal Value As Single)
			mp_fBCWS = Value
		End Set
	End Property

	Public Property fBCWP() As Single
		Get
			Return mp_fBCWP
		End Get
		Set(ByVal Value As Single)
			mp_fBCWP = Value
		End Set
	End Property

	Public Property lPhysicalPercentComplete() As Integer
		Get
			Return mp_lPhysicalPercentComplete
		End Get
		Set(ByVal Value As Integer)
			mp_lPhysicalPercentComplete = Value
		End Set
	End Property

	Public Property yEarnedValueMethod() As E_EARNEDVALUEMETHOD
		Get
			Return mp_yEarnedValueMethod
		End Get
		Set(ByVal Value As E_EARNEDVALUEMETHOD)
			mp_yEarnedValueMethod = Value
		End Set
	End Property

	Public ReadOnly Property oPredecessorLink_C() As TaskPredecessorLink_C
		Get
			Return mp_oPredecessorLink_C
		End Get
	End Property

	Public ReadOnly Property oActualWorkProtected() As Duration
		Get
			Return mp_oActualWorkProtected
		End Get
	End Property

	Public ReadOnly Property oActualOvertimeWorkProtected() As Duration
		Get
			Return mp_oActualOvertimeWorkProtected
		End Get
	End Property

	Public ReadOnly Property oExtendedAttribute_C() As TaskExtendedAttribute_C
		Get
			Return mp_oExtendedAttribute_C
		End Get
	End Property

	Public ReadOnly Property oBaseline_C() As TaskBaseline_C
		Get
			Return mp_oBaseline_C
		End Get
	End Property

	Public ReadOnly Property oOutlineCode_C() As TaskOutlineCode_C
		Get
			Return mp_oOutlineCode_C
		End Get
	End Property

	Public Property bIsPublished() As Boolean
		Get
			Return mp_bIsPublished
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsPublished = Value
		End Set
	End Property

	Public Property sStatusManager() As String
		Get
			Return mp_sStatusManager
		End Get
		Set(ByVal Value As String)
			mp_sStatusManager = Value
		End Set
	End Property

	Public Property dtCommitmentStart() As System.DateTime
		Get
			Return mp_dtCommitmentStart
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCommitmentStart = Value
		End Set
	End Property

	Public Property dtCommitmentFinish() As System.DateTime
		Get
			Return mp_dtCommitmentFinish
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCommitmentFinish = Value
		End Set
	End Property

	Public Property yCommitmentType() As E_COMMITMENTTYPE
		Get
			Return mp_yCommitmentType
		End Get
		Set(ByVal Value As E_COMMITMENTTYPE)
			mp_yCommitmentType = Value
		End Set
	End Property

	Public Property bActive() As Boolean
		Get
			Return mp_bActive
		End Get
		Set(ByVal Value As Boolean)
			mp_bActive = Value
		End Set
	End Property

	Public Property bPinned() As Boolean
		Get
			Return mp_bPinned
		End Get
		Set(ByVal Value As Boolean)
			mp_bPinned = Value
		End Set
	End Property

	Public Property sPinnedStart() As String
		Get
			Return mp_sPinnedStart
		End Get
		Set(ByVal Value As String)
			mp_sPinnedStart = Value
		End Set
	End Property

	Public Property sPinnedFinish() As String
		Get
			Return mp_sPinnedFinish
		End Get
		Set(ByVal Value As String)
			mp_sPinnedFinish = Value
		End Set
	End Property

	Public Property sPinnedDuration() As String
		Get
			Return mp_sPinnedDuration
		End Get
		Set(ByVal Value As String)
			mp_sPinnedDuration = Value
		End Set
	End Property

	Public ReadOnly Property oTimephasedData_C() As TimephasedData_C
		Get
			Return mp_oTimephasedData_C
		End Get
	End Property

	Public Property Key() As String
		Get
			Return mp_sKey
		End Get
		Set(ByVal Value As String)
			mp_oCollection.mp_SetKey(mp_sKey, Value, SYS_ERRORS.MP_SET_KEY)
		End Set
	End Property

	Public Function IsNull() As Boolean
		Dim bReturn As Boolean = True
		If mp_lUID <> 0 Then
			bReturn = False
		End If
		If mp_lID <> 0 Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_yType <> E_TYPE_4.T_4_FIXED_UNITS Then
			bReturn = False
		End If
		If mp_bIsNull <> False Then
			bReturn = False
		End If
		If mp_dtCreateDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_sContact <> "" Then
			bReturn = False
		End If
		If mp_sWBS <> "" Then
			bReturn = False
		End If
		If mp_sWBSLevel <> "" Then
			bReturn = False
		End If
		If mp_sOutlineNumber <> "" Then
			bReturn = False
		End If
		If mp_lOutlineLevel <> 0 Then
			bReturn = False
		End If
		If mp_lPriority <> 0 Then
			bReturn = False
		End If
		If mp_dtStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_oDuration.IsNull() = False Then
			bReturn = False
		End If
		If mp_yDurationFormat <> E_DURATIONFORMAT.DF_M Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_dtStop.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtResume.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bResumeValid <> False Then
			bReturn = False
		End If
		If mp_bEffortDriven <> False Then
			bReturn = False
		End If
		If mp_bRecurring <> False Then
			bReturn = False
		End If
		If mp_bOverAllocated <> False Then
			bReturn = False
		End If
		If mp_bEstimated <> False Then
			bReturn = False
		End If
		If mp_bMilestone <> False Then
			bReturn = False
		End If
		If mp_bSummary <> False Then
			bReturn = False
		End If
		If mp_bDisplayAsSummary <> False Then
			bReturn = False
		End If
		If mp_bCritical <> False Then
			bReturn = False
		End If
		If mp_bIsSubproject <> False Then
			bReturn = False
		End If
		If mp_bIsSubprojectReadOnly <> False Then
			bReturn = False
		End If
		If mp_sSubprojectName <> "" Then
			bReturn = False
		End If
		If mp_bExternalTask <> False Then
			bReturn = False
		End If
		If mp_sExternalTaskProject <> "" Then
			bReturn = False
		End If
		If mp_dtEarlyStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtEarlyFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtLateStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtLateFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_lStartVariance <> 0 Then
			bReturn = False
		End If
		If mp_lFinishVariance <> 0 Then
			bReturn = False
		End If
		If mp_fWorkVariance <> 0 Then
			bReturn = False
		End If
		If mp_lFreeSlack <> 0 Then
			bReturn = False
		End If
		If mp_lStartSlack <> 0 Then
			bReturn = False
		End If
		If mp_lFinishSlack <> 0 Then
			bReturn = False
		End If
		If mp_lTotalSlack <> 0 Then
			bReturn = False
		End If
		If mp_fFixedCost <> 0 Then
			bReturn = False
		End If
		If mp_yFixedCostAccrual <> E_FIXEDCOSTACCRUAL.FCA_START Then
			bReturn = False
		End If
		If mp_lPercentComplete <> 0 Then
			bReturn = False
		End If
		If mp_lPercentWorkComplete <> 0 Then
			bReturn = False
		End If
		If mp_cCost <> 0 Then
			bReturn = False
		End If
		If mp_cOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_dtActualStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtActualFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_oActualDuration.IsNull() = False Then
			bReturn = False
		End If
		If mp_cActualCost <> 0 Then
			bReturn = False
		End If
		If mp_cActualOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oActualWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRegularWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRemainingDuration.IsNull() = False Then
			bReturn = False
		End If
		If mp_cRemainingCost <> 0 Then
			bReturn = False
		End If
		If mp_oRemainingWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_cRemainingOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oRemainingOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_fACWP <> 0 Then
			bReturn = False
		End If
		If mp_fCV <> 0 Then
			bReturn = False
		End If
		If mp_yConstraintType <> E_CONSTRAINTTYPE.CT_AS_SOON_AS_POSSIBLE Then
			bReturn = False
		End If
		If mp_lCalendarUID <> 0 Then
			bReturn = False
		End If
		If mp_dtConstraintDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtDeadline.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bLevelAssignments <> False Then
			bReturn = False
		End If
		If mp_bLevelingCanSplit <> False Then
			bReturn = False
		End If
		If mp_lLevelingDelay <> 0 Then
			bReturn = False
		End If
		If mp_yLevelingDelayFormat <> E_LEVELINGDELAYFORMAT.LDF_M Then
			bReturn = False
		End If
		If mp_dtPreLeveledStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtPreLeveledFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_sHyperlink <> "" Then
			bReturn = False
		End If
		If mp_sHyperlinkAddress <> "" Then
			bReturn = False
		End If
		If mp_sHyperlinkSubAddress <> "" Then
			bReturn = False
		End If
		If mp_bIgnoreResourceCalendar <> False Then
			bReturn = False
		End If
		If mp_sNotes <> "" Then
			bReturn = False
		End If
		If mp_bHideBar <> False Then
			bReturn = False
		End If
		If mp_bRollup <> False Then
			bReturn = False
		End If
		If mp_fBCWS <> 0 Then
			bReturn = False
		End If
		If mp_fBCWP <> 0 Then
			bReturn = False
		End If
		If mp_lPhysicalPercentComplete <> 0 Then
			bReturn = False
		End If
		If mp_yEarnedValueMethod <> E_EARNEDVALUEMETHOD.EVM_PERCENT_COMPLETE Then
			bReturn = False
		End If
		If mp_oPredecessorLink_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualWorkProtected.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualOvertimeWorkProtected.IsNull() = False Then
			bReturn = False
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oBaseline_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oOutlineCode_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_bIsPublished <> False Then
			bReturn = False
		End If
		If mp_sStatusManager <> "" Then
			bReturn = False
		End If
		If mp_dtCommitmentStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtCommitmentFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_yCommitmentType <> E_COMMITMENTTYPE.CT_THE_TASK_HAS_NO_DELIVERABLE_OR_DEPENDENCY_ON_A_DELIVERABLE Then
			bReturn = False
		End If
		If mp_bActive <> False Then
			bReturn = False
		End If
		If mp_bPinned <> False Then
			bReturn = False
		End If
		If mp_sPinnedStart <> "" Then
			bReturn = False
		End If
		If mp_sPinnedFinish <> "" Then
			bReturn = False
		End If
		If mp_sPinnedDuration <> "" Then
			bReturn = False
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Task/>"
		End if
		Dim oXML As New clsXML("Task")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("UID", mp_lUID)
		oXML.WriteProperty("ID", mp_lID)
		If mp_sName <> "" Then
			oXML.WriteProperty("Name", mp_sName)
		End If
		oXML.WriteProperty("Type", mp_yType)
		oXML.WriteProperty("IsNull", mp_bIsNull)
		If mp_dtCreateDate.Ticks <> 0 Then
			oXML.WriteProperty("CreateDate", mp_dtCreateDate)
		End If
		If mp_sContact <> "" Then
			oXML.WriteProperty("Contact", mp_sContact)
		End If
		If mp_sWBS <> "" Then
			oXML.WriteProperty("WBS", mp_sWBS)
		End If
		If mp_sWBSLevel <> "" Then
			oXML.WriteProperty("WBSLevel", mp_sWBSLevel)
		End If
		If mp_sOutlineNumber <> "" Then
			oXML.WriteProperty("OutlineNumber", mp_sOutlineNumber)
		End If
		oXML.WriteProperty("OutlineLevel", mp_lOutlineLevel)
		oXML.WriteProperty("Priority", mp_lPriority)
		If mp_dtStart.Ticks <> 0 Then
			oXML.WriteProperty("Start", mp_dtStart)
		End If
		If mp_dtFinish.Ticks <> 0 Then
			oXML.WriteProperty("Finish", mp_dtFinish)
		End If
		oXML.WriteProperty("Duration", mp_oDuration)
		oXML.WriteProperty("DurationFormat", mp_yDurationFormat)
		oXML.WriteProperty("Work", mp_oWork)
		If mp_dtStop.Ticks <> 0 Then
			oXML.WriteProperty("Stop", mp_dtStop)
		End If
		If mp_dtResume.Ticks <> 0 Then
			oXML.WriteProperty("Resume", mp_dtResume)
		End If
		oXML.WriteProperty("ResumeValid", mp_bResumeValid)
		oXML.WriteProperty("EffortDriven", mp_bEffortDriven)
		oXML.WriteProperty("Recurring", mp_bRecurring)
		oXML.WriteProperty("OverAllocated", mp_bOverAllocated)
		oXML.WriteProperty("Estimated", mp_bEstimated)
		oXML.WriteProperty("Milestone", mp_bMilestone)
		oXML.WriteProperty("Summary", mp_bSummary)
		oXML.WriteProperty("DisplayAsSummary", mp_bDisplayAsSummary)
		oXML.WriteProperty("Critical", mp_bCritical)
		oXML.WriteProperty("IsSubproject", mp_bIsSubproject)
		oXML.WriteProperty("IsSubprojectReadOnly", mp_bIsSubprojectReadOnly)
		If mp_sSubprojectName <> "" Then
			oXML.WriteProperty("SubprojectName", mp_sSubprojectName)
		End If
		oXML.WriteProperty("ExternalTask", mp_bExternalTask)
		If mp_sExternalTaskProject <> "" Then
			oXML.WriteProperty("ExternalTaskProject", mp_sExternalTaskProject)
		End If
		If mp_dtEarlyStart.Ticks <> 0 Then
			oXML.WriteProperty("EarlyStart", mp_dtEarlyStart)
		End If
		If mp_dtEarlyFinish.Ticks <> 0 Then
			oXML.WriteProperty("EarlyFinish", mp_dtEarlyFinish)
		End If
		If mp_dtLateStart.Ticks <> 0 Then
			oXML.WriteProperty("LateStart", mp_dtLateStart)
		End If
		If mp_dtLateFinish.Ticks <> 0 Then
			oXML.WriteProperty("LateFinish", mp_dtLateFinish)
		End If
		oXML.WriteProperty("StartVariance", mp_lStartVariance)
		oXML.WriteProperty("FinishVariance", mp_lFinishVariance)
		oXML.WriteProperty("WorkVariance", mp_fWorkVariance)
		oXML.WriteProperty("FreeSlack", mp_lFreeSlack)
		oXML.WriteProperty("StartSlack", mp_lStartSlack)
		oXML.WriteProperty("FinishSlack", mp_lFinishSlack)
		oXML.WriteProperty("TotalSlack", mp_lTotalSlack)
		oXML.WriteProperty("FixedCost", mp_fFixedCost)
		oXML.WriteProperty("FixedCostAccrual", mp_yFixedCostAccrual)
		oXML.WriteProperty("PercentComplete", mp_lPercentComplete)
		oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.WriteProperty("Cost", mp_cCost)
		oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork)
		If mp_dtActualStart.Ticks <> 0 Then
			oXML.WriteProperty("ActualStart", mp_dtActualStart)
		End If
		If mp_dtActualFinish.Ticks <> 0 Then
			oXML.WriteProperty("ActualFinish", mp_dtActualFinish)
		End If
		oXML.WriteProperty("ActualDuration", mp_oActualDuration)
		oXML.WriteProperty("ActualCost", mp_cActualCost)
		oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.WriteProperty("ActualWork", mp_oActualWork)
		oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		oXML.WriteProperty("RegularWork", mp_oRegularWork)
		oXML.WriteProperty("RemainingDuration", mp_oRemainingDuration)
		oXML.WriteProperty("RemainingCost", mp_cRemainingCost)
		oXML.WriteProperty("RemainingWork", mp_oRemainingWork)
		oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.WriteProperty("ACWP", mp_fACWP)
		oXML.WriteProperty("CV", mp_fCV)
		oXML.WriteProperty("ConstraintType", mp_yConstraintType)
		oXML.WriteProperty("CalendarUID", mp_lCalendarUID)
		If mp_dtConstraintDate.Ticks <> 0 Then
			oXML.WriteProperty("ConstraintDate", mp_dtConstraintDate)
		End If
		If mp_dtDeadline.Ticks <> 0 Then
			oXML.WriteProperty("Deadline", mp_dtDeadline)
		End If
		oXML.WriteProperty("LevelAssignments", mp_bLevelAssignments)
		oXML.WriteProperty("LevelingCanSplit", mp_bLevelingCanSplit)
		oXML.WriteProperty("LevelingDelay", mp_lLevelingDelay)
		oXML.WriteProperty("LevelingDelayFormat", mp_yLevelingDelayFormat)
		If mp_dtPreLeveledStart.Ticks <> 0 Then
			oXML.WriteProperty("PreLeveledStart", mp_dtPreLeveledStart)
		End If
		If mp_dtPreLeveledFinish.Ticks <> 0 Then
			oXML.WriteProperty("PreLeveledFinish", mp_dtPreLeveledFinish)
		End If
		If mp_sHyperlink <> "" Then
			oXML.WriteProperty("Hyperlink", mp_sHyperlink)
		End If
		If mp_sHyperlinkAddress <> "" Then
			oXML.WriteProperty("HyperlinkAddress", mp_sHyperlinkAddress)
		End If
		If mp_sHyperlinkSubAddress <> "" Then
			oXML.WriteProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress)
		End If
		oXML.WriteProperty("IgnoreResourceCalendar", mp_bIgnoreResourceCalendar)
		If mp_sNotes <> "" Then
			oXML.WriteProperty("Notes", mp_sNotes)
		End If
		oXML.WriteProperty("HideBar", mp_bHideBar)
		oXML.WriteProperty("Rollup", mp_bRollup)
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		oXML.WriteProperty("PhysicalPercentComplete", mp_lPhysicalPercentComplete)
		oXML.WriteProperty("EarnedValueMethod", mp_yEarnedValueMethod)
		If mp_oPredecessorLink_C.IsNull() = False Then
			mp_oPredecessorLink_C.WriteObjectProtected(oXML)
		End If
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		If mp_oExtendedAttribute_C.IsNull() = False Then
			mp_oExtendedAttribute_C.WriteObjectProtected(oXML)
		End If
		If mp_oBaseline_C.IsNull() = False Then
			mp_oBaseline_C.WriteObjectProtected(oXML)
		End If
		If mp_oOutlineCode_C.IsNull() = False Then
			mp_oOutlineCode_C.WriteObjectProtected(oXML)
		End If
		oXML.WriteProperty("IsPublished", mp_bIsPublished)
		If mp_sStatusManager <> "" Then
			oXML.WriteProperty("StatusManager", mp_sStatusManager)
		End If
		If mp_dtCommitmentStart.Ticks <> 0 Then
			oXML.WriteProperty("CommitmentStart", mp_dtCommitmentStart)
		End If
		If mp_dtCommitmentFinish.Ticks <> 0 Then
			oXML.WriteProperty("CommitmentFinish", mp_dtCommitmentFinish)
		End If
		oXML.WriteProperty("CommitmentType", mp_yCommitmentType)
		oXML.WriteProperty("Active", mp_bActive)
		oXML.WriteProperty("Pinned", mp_bPinned)
		If mp_sPinnedStart <> "" Then
			oXML.WriteProperty("PinnedStart", mp_sPinnedStart)
		End If
		If mp_sPinnedFinish <> "" Then
			oXML.WriteProperty("PinnedFinish", mp_sPinnedFinish)
		End If
		If mp_sPinnedDuration <> "" Then
			oXML.WriteProperty("PinnedDuration", mp_sPinnedDuration)
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			mp_oTimephasedData_C.WriteObjectProtected(oXML)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Task")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("UID", mp_lUID)
		oXML.ReadProperty("ID", mp_lID)
		oXML.ReadProperty("Name", mp_sName)
		If mp_sName.Length > 512 Then
			mp_sName = mp_sName.Substring(0, 512)
		End If
		oXML.ReadProperty("Type", mp_yType)
		oXML.ReadProperty("IsNull", mp_bIsNull)
		oXML.ReadProperty("CreateDate", mp_dtCreateDate)
		oXML.ReadProperty("Contact", mp_sContact)
		If mp_sContact.Length > 512 Then
			mp_sContact = mp_sContact.Substring(0, 512)
		End If
		oXML.ReadProperty("WBS", mp_sWBS)
		oXML.ReadProperty("WBSLevel", mp_sWBSLevel)
		oXML.ReadProperty("OutlineNumber", mp_sOutlineNumber)
		If mp_sOutlineNumber.Length > 512 Then
			mp_sOutlineNumber = mp_sOutlineNumber.Substring(0, 512)
		End If
		oXML.ReadProperty("OutlineLevel", mp_lOutlineLevel)
		oXML.ReadProperty("Priority", mp_lPriority)
		oXML.ReadProperty("Start", mp_dtStart)
		oXML.ReadProperty("Finish", mp_dtFinish)
		oXML.ReadProperty("Duration", mp_oDuration)
		oXML.ReadProperty("DurationFormat", mp_yDurationFormat)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("Stop", mp_dtStop)
		oXML.ReadProperty("Resume", mp_dtResume)
		oXML.ReadProperty("ResumeValid", mp_bResumeValid)
		oXML.ReadProperty("EffortDriven", mp_bEffortDriven)
		oXML.ReadProperty("Recurring", mp_bRecurring)
		oXML.ReadProperty("OverAllocated", mp_bOverAllocated)
		oXML.ReadProperty("Estimated", mp_bEstimated)
		oXML.ReadProperty("Milestone", mp_bMilestone)
		oXML.ReadProperty("Summary", mp_bSummary)
		oXML.ReadProperty("DisplayAsSummary", mp_bDisplayAsSummary)
		oXML.ReadProperty("Critical", mp_bCritical)
		oXML.ReadProperty("IsSubproject", mp_bIsSubproject)
		oXML.ReadProperty("IsSubprojectReadOnly", mp_bIsSubprojectReadOnly)
		oXML.ReadProperty("SubprojectName", mp_sSubprojectName)
		If mp_sSubprojectName.Length > 512 Then
			mp_sSubprojectName = mp_sSubprojectName.Substring(0, 512)
		End If
		oXML.ReadProperty("ExternalTask", mp_bExternalTask)
		oXML.ReadProperty("ExternalTaskProject", mp_sExternalTaskProject)
		If mp_sExternalTaskProject.Length > 512 Then
			mp_sExternalTaskProject = mp_sExternalTaskProject.Substring(0, 512)
		End If
		oXML.ReadProperty("EarlyStart", mp_dtEarlyStart)
		oXML.ReadProperty("EarlyFinish", mp_dtEarlyFinish)
		oXML.ReadProperty("LateStart", mp_dtLateStart)
		oXML.ReadProperty("LateFinish", mp_dtLateFinish)
		oXML.ReadProperty("StartVariance", mp_lStartVariance)
		oXML.ReadProperty("FinishVariance", mp_lFinishVariance)
		oXML.ReadProperty("WorkVariance", mp_fWorkVariance)
		oXML.ReadProperty("FreeSlack", mp_lFreeSlack)
		oXML.ReadProperty("StartSlack", mp_lStartSlack)
		oXML.ReadProperty("FinishSlack", mp_lFinishSlack)
		oXML.ReadProperty("TotalSlack", mp_lTotalSlack)
		oXML.ReadProperty("FixedCost", mp_fFixedCost)
		oXML.ReadProperty("FixedCostAccrual", mp_yFixedCostAccrual)
		oXML.ReadProperty("PercentComplete", mp_lPercentComplete)
		oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.ReadProperty("Cost", mp_cCost)
		oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork)
		oXML.ReadProperty("ActualStart", mp_dtActualStart)
		oXML.ReadProperty("ActualFinish", mp_dtActualFinish)
		oXML.ReadProperty("ActualDuration", mp_oActualDuration)
		oXML.ReadProperty("ActualCost", mp_cActualCost)
		oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.ReadProperty("ActualWork", mp_oActualWork)
		oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		oXML.ReadProperty("RegularWork", mp_oRegularWork)
		oXML.ReadProperty("RemainingDuration", mp_oRemainingDuration)
		oXML.ReadProperty("RemainingCost", mp_cRemainingCost)
		oXML.ReadProperty("RemainingWork", mp_oRemainingWork)
		oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.ReadProperty("ACWP", mp_fACWP)
		oXML.ReadProperty("CV", mp_fCV)
		oXML.ReadProperty("ConstraintType", mp_yConstraintType)
		oXML.ReadProperty("CalendarUID", mp_lCalendarUID)
		oXML.ReadProperty("ConstraintDate", mp_dtConstraintDate)
		oXML.ReadProperty("Deadline", mp_dtDeadline)
		oXML.ReadProperty("LevelAssignments", mp_bLevelAssignments)
		oXML.ReadProperty("LevelingCanSplit", mp_bLevelingCanSplit)
		oXML.ReadProperty("LevelingDelay", mp_lLevelingDelay)
		oXML.ReadProperty("LevelingDelayFormat", mp_yLevelingDelayFormat)
		oXML.ReadProperty("PreLeveledStart", mp_dtPreLeveledStart)
		oXML.ReadProperty("PreLeveledFinish", mp_dtPreLeveledFinish)
		oXML.ReadProperty("Hyperlink", mp_sHyperlink)
		If mp_sHyperlink.Length > 512 Then
			mp_sHyperlink = mp_sHyperlink.Substring(0, 512)
		End If
		oXML.ReadProperty("HyperlinkAddress", mp_sHyperlinkAddress)
		If mp_sHyperlinkAddress.Length > 512 Then
			mp_sHyperlinkAddress = mp_sHyperlinkAddress.Substring(0, 512)
		End If
		oXML.ReadProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress)
		If mp_sHyperlinkSubAddress.Length > 512 Then
			mp_sHyperlinkSubAddress = mp_sHyperlinkSubAddress.Substring(0, 512)
		End If
		oXML.ReadProperty("IgnoreResourceCalendar", mp_bIgnoreResourceCalendar)
		oXML.ReadProperty("Notes", mp_sNotes)
		oXML.ReadProperty("HideBar", mp_bHideBar)
		oXML.ReadProperty("Rollup", mp_bRollup)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
		oXML.ReadProperty("PhysicalPercentComplete", mp_lPhysicalPercentComplete)
		oXML.ReadProperty("EarnedValueMethod", mp_yEarnedValueMethod)
		mp_oPredecessorLink_C.ReadObjectProtected(oXML)
		oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		mp_oExtendedAttribute_C.ReadObjectProtected(oXML)
		mp_oBaseline_C.ReadObjectProtected(oXML)
		mp_oOutlineCode_C.ReadObjectProtected(oXML)
		oXML.ReadProperty("IsPublished", mp_bIsPublished)
		oXML.ReadProperty("StatusManager", mp_sStatusManager)
		oXML.ReadProperty("CommitmentStart", mp_dtCommitmentStart)
		oXML.ReadProperty("CommitmentFinish", mp_dtCommitmentFinish)
		oXML.ReadProperty("CommitmentType", mp_yCommitmentType)
		oXML.ReadProperty("Active", mp_bActive)
		oXML.ReadProperty("Pinned", mp_bPinned)
		oXML.ReadProperty("PinnedStart", mp_sPinnedStart)
		oXML.ReadProperty("PinnedFinish", mp_sPinnedFinish)
		oXML.ReadProperty("PinnedDuration", mp_sPinnedDuration)
		mp_oTimephasedData_C.ReadObjectProtected(oXML)
	End Sub

End Class
