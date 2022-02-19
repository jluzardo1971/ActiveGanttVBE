Option Explicit On

Public Class MP14


	Private mp_lSaveVersion As Integer
	Private mp_sUID As String
	Private mp_sName As String
	Private mp_sTitle As String
	Private mp_sSubject As String
	Private mp_sCategory As String
	Private mp_sCompany As String
	Private mp_sManager As String
	Private mp_sAuthor As String
	Private mp_dtCreationDate As System.DateTime
	Private mp_lRevision As Integer
	Private mp_dtLastSaved As System.DateTime
	Private mp_bScheduleFromStart As Boolean
	Private mp_dtStartDate As System.DateTime
	Private mp_dtFinishDate As System.DateTime
	Private mp_yFYStartDate As E_FYSTARTDATE
	Private mp_lCriticalSlackLimit As Integer
	Private mp_lCurrencyDigits As Integer
	Private mp_sCurrencySymbol As String
	Private mp_sCurrencyCode As String
	Private mp_yCurrencySymbolPosition As E_CURRENCYSYMBOLPOSITION
	Private mp_lCalendarUID As Integer
	Private mp_oDefaultStartTime As Time
	Private mp_oDefaultFinishTime As Time
	Private mp_lMinutesPerDay As Integer
	Private mp_lMinutesPerWeek As Integer
	Private mp_lDaysPerMonth As Integer
	Private mp_yDefaultTaskType As E_DEFAULTTASKTYPE
	Private mp_yDefaultFixedCostAccrual As E_DEFAULTFIXEDCOSTACCRUAL
	Private mp_fDefaultStandardRate As Single
	Private mp_fDefaultOvertimeRate As Single
	Private mp_yDurationFormat As E_DURATIONFORMAT
	Private mp_yWorkFormat As E_WORKFORMAT
	Private mp_bEditableActualCosts As Boolean
	Private mp_bHonorConstraints As Boolean
	Private mp_yEarnedValueMethod As E_EARNEDVALUEMETHOD
	Private mp_bInsertedProjectsLikeSummary As Boolean
	Private mp_bMultipleCriticalPaths As Boolean
	Private mp_bNewTasksEffortDriven As Boolean
	Private mp_bNewTasksEstimated As Boolean
	Private mp_bSplitsInProgressTasks As Boolean
	Private mp_bSpreadActualCost As Boolean
	Private mp_bSpreadPercentComplete As Boolean
	Private mp_bTaskUpdatesResource As Boolean
	Private mp_bFiscalYearStart As Boolean
	Private mp_yWeekStartDay As E_WEEKSTARTDAY
	Private mp_bMoveCompletedEndsBack As Boolean
	Private mp_bMoveRemainingStartsBack As Boolean
	Private mp_bMoveRemainingStartsForward As Boolean
	Private mp_bMoveCompletedEndsForward As Boolean
	Private mp_yBaselineForEarnedValue As E_BASELINEFOREARNEDVALUE
	Private mp_bAutoAddNewResourcesAndTasks As Boolean
	Private mp_dtStatusDate As System.DateTime
	Private mp_dtCurrentDate As System.DateTime
	Private mp_bMicrosoftProjectServerURL As Boolean
	Private mp_bAutolink As Boolean
	Private mp_yNewTaskStartDate As E_NEWTASKSTARTDATE
	Private mp_yDefaultTaskEVMethod As E_DEFAULTTASKEVMETHOD
	Private mp_bProjectExternallyEdited As Boolean
	Private mp_dtExtendedCreationDate As System.DateTime
	Private mp_bActualsInSync As Boolean
	Private mp_bRemoveFileProperties As Boolean
	Private mp_bAdminProject As Boolean
	Private mp_sBaslineCalendar As String
	Private mp_bNewTasksAreManual As Boolean
	Private mp_bUpdateManuallyScheduledTasksWhenEditingLinks As Boolean
	Private mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled As Boolean
	Private mp_oOutlineCodes As OutlineCodes
	Private mp_oWBSMasks As WBSMasks
	Private mp_oExtendedAttributes As ExtendedAttributes
	Private mp_oCalendars As Calendars
	Private mp_oTasks As Tasks
	Private mp_oResources As Resources
	Private mp_oAssignments As Assignments

	Public Sub New()
		mp_lSaveVersion = 0
		mp_sUID = ""
		mp_sName = ""
		mp_sTitle = ""
		mp_sSubject = ""
		mp_sCategory = ""
		mp_sCompany = ""
		mp_sManager = ""
		mp_sAuthor = ""
		mp_dtCreationDate = New System.DateTime(0)
		mp_lRevision = 0
		mp_dtLastSaved = New System.DateTime(0)
		mp_bScheduleFromStart = True
		mp_dtStartDate = New System.DateTime(0)
		mp_dtFinishDate = New System.DateTime(0)
		mp_yFYStartDate = E_FYSTARTDATE.FYSD_JANUARY
		mp_lCriticalSlackLimit = 0
		mp_lCurrencyDigits = 0
		mp_sCurrencySymbol = ""
		mp_sCurrencyCode = ""
		mp_yCurrencySymbolPosition = E_CURRENCYSYMBOLPOSITION.CSP_BEFORE
		mp_lCalendarUID = 0
		mp_oDefaultStartTime = New Time()
		mp_oDefaultFinishTime = New Time()
		mp_lMinutesPerDay = 0
		mp_lMinutesPerWeek = 0
		mp_lDaysPerMonth = 0
		mp_yDefaultTaskType = E_DEFAULTTASKTYPE.DTT_FIXED_DURATION
		mp_yDefaultFixedCostAccrual = E_DEFAULTFIXEDCOSTACCRUAL.DFCA_START
		mp_fDefaultStandardRate = 0
		mp_fDefaultOvertimeRate = 0
		mp_yDurationFormat = E_DURATIONFORMAT.DF_M
		mp_yWorkFormat = E_WORKFORMAT.WF_M
		mp_bEditableActualCosts = False
		mp_bHonorConstraints = True
		mp_yEarnedValueMethod = E_EARNEDVALUEMETHOD.EVM_PERCENT_COMPLETE
		mp_bInsertedProjectsLikeSummary = True
		mp_bMultipleCriticalPaths = False
		mp_bNewTasksEffortDriven = True
		mp_bNewTasksEstimated = True
		mp_bSplitsInProgressTasks = True
		mp_bSpreadActualCost = True
		mp_bSpreadPercentComplete = False
		mp_bTaskUpdatesResource = False
		mp_bFiscalYearStart = False
		mp_yWeekStartDay = E_WEEKSTARTDAY.WSD_SUNDAY
		mp_bMoveCompletedEndsBack = False
		mp_bMoveRemainingStartsBack = False
		mp_bMoveRemainingStartsForward = False
		mp_bMoveCompletedEndsForward = False
		mp_yBaselineForEarnedValue = E_BASELINEFOREARNEDVALUE.BFEV_BASELINE
		mp_bAutoAddNewResourcesAndTasks = True
		mp_dtStatusDate = New System.DateTime(0)
		mp_dtCurrentDate = New System.DateTime(0)
		mp_bMicrosoftProjectServerURL = False
		mp_bAutolink = False
		mp_yNewTaskStartDate = E_NEWTASKSTARTDATE.NTSD_PROJECT_START_DATE
		mp_yDefaultTaskEVMethod = E_DEFAULTTASKEVMETHOD.DTEVM_PERCENT_COMPLETE
		mp_bProjectExternallyEdited = False
		mp_dtExtendedCreationDate = New System.DateTime(0)
		mp_bActualsInSync = False
		mp_bRemoveFileProperties = False
		mp_bAdminProject = False
		mp_sBaslineCalendar = ""
		mp_bNewTasksAreManual = False
		mp_bUpdateManuallyScheduledTasksWhenEditingLinks = False
		mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled = False
		mp_oOutlineCodes = New OutlineCodes()
		mp_oWBSMasks = New WBSMasks()
		mp_oExtendedAttributes = New ExtendedAttributes()
		mp_oCalendars = New Calendars()
		mp_oTasks = New Tasks()
		mp_oResources = New Resources()
		mp_oAssignments = New Assignments()
	End Sub

	Public Property lSaveVersion() As Integer
		Get
			Return mp_lSaveVersion
		End Get
		Set(ByVal Value As Integer)
			mp_lSaveVersion = Value
		End Set
	End Property

	Public Property sUID() As String
		Get
			Return mp_sUID
		End Get
		Set(ByVal Value As String)
			If Value.Length > 16 Then
				Value = Value.Substring(0, 16)
			End If
			mp_sUID = Value
		End Set
	End Property

	Public Property sName() As String
		Get
			Return mp_sName
		End Get
		Set(ByVal Value As String)
			If Value.Length > 255 Then
				Value = Value.Substring(0, 255)
			End If
			mp_sName = Value
		End Set
	End Property

	Public Property sTitle() As String
		Get
			Return mp_sTitle
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sTitle = Value
		End Set
	End Property

	Public Property sSubject() As String
		Get
			Return mp_sSubject
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sSubject = Value
		End Set
	End Property

	Public Property sCategory() As String
		Get
			Return mp_sCategory
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sCategory = Value
		End Set
	End Property

	Public Property sCompany() As String
		Get
			Return mp_sCompany
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sCompany = Value
		End Set
	End Property

	Public Property sManager() As String
		Get
			Return mp_sManager
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sManager = Value
		End Set
	End Property

	Public Property sAuthor() As String
		Get
			Return mp_sAuthor
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sAuthor = Value
		End Set
	End Property

	Public Property dtCreationDate() As System.DateTime
		Get
			Return mp_dtCreationDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCreationDate = Value
		End Set
	End Property

	Public Property lRevision() As Integer
		Get
			Return mp_lRevision
		End Get
		Set(ByVal Value As Integer)
			mp_lRevision = Value
		End Set
	End Property

	Public Property dtLastSaved() As System.DateTime
		Get
			Return mp_dtLastSaved
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtLastSaved = Value
		End Set
	End Property

	Public Property bScheduleFromStart() As Boolean
		Get
			Return mp_bScheduleFromStart
		End Get
		Set(ByVal Value As Boolean)
			mp_bScheduleFromStart = Value
		End Set
	End Property

	Public Property dtStartDate() As System.DateTime
		Get
			Return mp_dtStartDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtStartDate = Value
		End Set
	End Property

	Public Property dtFinishDate() As System.DateTime
		Get
			Return mp_dtFinishDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtFinishDate = Value
		End Set
	End Property

	Public Property yFYStartDate() As E_FYSTARTDATE
		Get
			Return mp_yFYStartDate
		End Get
		Set(ByVal Value As E_FYSTARTDATE)
			mp_yFYStartDate = Value
		End Set
	End Property

	Public Property lCriticalSlackLimit() As Integer
		Get
			Return mp_lCriticalSlackLimit
		End Get
		Set(ByVal Value As Integer)
			mp_lCriticalSlackLimit = Value
		End Set
	End Property

	Public Property lCurrencyDigits() As Integer
		Get
			Return mp_lCurrencyDigits
		End Get
		Set(ByVal Value As Integer)
			mp_lCurrencyDigits = Value
		End Set
	End Property

	Public Property sCurrencySymbol() As String
		Get
			Return mp_sCurrencySymbol
		End Get
		Set(ByVal Value As String)
			If Value.Length > 20 Then
				Value = Value.Substring(0, 20)
			End If
			mp_sCurrencySymbol = Value
		End Set
	End Property

	Public Property sCurrencyCode() As String
		Get
			Return mp_sCurrencyCode
		End Get
		Set(ByVal Value As String)
			If Value.Length > 3 Then
				Value = Value.Substring(0, 3)
			End If
			mp_sCurrencyCode = Value
		End Set
	End Property

	Public Property yCurrencySymbolPosition() As E_CURRENCYSYMBOLPOSITION
		Get
			Return mp_yCurrencySymbolPosition
		End Get
		Set(ByVal Value As E_CURRENCYSYMBOLPOSITION)
			mp_yCurrencySymbolPosition = Value
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

	Public ReadOnly Property oDefaultStartTime() As Time
		Get
			Return mp_oDefaultStartTime
		End Get
	End Property

	Public ReadOnly Property oDefaultFinishTime() As Time
		Get
			Return mp_oDefaultFinishTime
		End Get
	End Property

	Public Property lMinutesPerDay() As Integer
		Get
			Return mp_lMinutesPerDay
		End Get
		Set(ByVal Value As Integer)
			mp_lMinutesPerDay = Value
		End Set
	End Property

	Public Property lMinutesPerWeek() As Integer
		Get
			Return mp_lMinutesPerWeek
		End Get
		Set(ByVal Value As Integer)
			mp_lMinutesPerWeek = Value
		End Set
	End Property

	Public Property lDaysPerMonth() As Integer
		Get
			Return mp_lDaysPerMonth
		End Get
		Set(ByVal Value As Integer)
			mp_lDaysPerMonth = Value
		End Set
	End Property

	Public Property yDefaultTaskType() As E_DEFAULTTASKTYPE
		Get
			Return mp_yDefaultTaskType
		End Get
		Set(ByVal Value As E_DEFAULTTASKTYPE)
			mp_yDefaultTaskType = Value
		End Set
	End Property

	Public Property yDefaultFixedCostAccrual() As E_DEFAULTFIXEDCOSTACCRUAL
		Get
			Return mp_yDefaultFixedCostAccrual
		End Get
		Set(ByVal Value As E_DEFAULTFIXEDCOSTACCRUAL)
			mp_yDefaultFixedCostAccrual = Value
		End Set
	End Property

	Public Property fDefaultStandardRate() As Single
		Get
			Return mp_fDefaultStandardRate
		End Get
		Set(ByVal Value As Single)
			mp_fDefaultStandardRate = Value
		End Set
	End Property

	Public Property fDefaultOvertimeRate() As Single
		Get
			Return mp_fDefaultOvertimeRate
		End Get
		Set(ByVal Value As Single)
			mp_fDefaultOvertimeRate = Value
		End Set
	End Property

	Public Property yDurationFormat() As E_DURATIONFORMAT
		Get
			Return mp_yDurationFormat
		End Get
		Set(ByVal Value As E_DURATIONFORMAT)
			mp_yDurationFormat = Value
		End Set
	End Property

	Public Property yWorkFormat() As E_WORKFORMAT
		Get
			Return mp_yWorkFormat
		End Get
		Set(ByVal Value As E_WORKFORMAT)
			mp_yWorkFormat = Value
		End Set
	End Property

	Public Property bEditableActualCosts() As Boolean
		Get
			Return mp_bEditableActualCosts
		End Get
		Set(ByVal Value As Boolean)
			mp_bEditableActualCosts = Value
		End Set
	End Property

	Public Property bHonorConstraints() As Boolean
		Get
			Return mp_bHonorConstraints
		End Get
		Set(ByVal Value As Boolean)
			mp_bHonorConstraints = Value
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

	Public Property bInsertedProjectsLikeSummary() As Boolean
		Get
			Return mp_bInsertedProjectsLikeSummary
		End Get
		Set(ByVal Value As Boolean)
			mp_bInsertedProjectsLikeSummary = Value
		End Set
	End Property

	Public Property bMultipleCriticalPaths() As Boolean
		Get
			Return mp_bMultipleCriticalPaths
		End Get
		Set(ByVal Value As Boolean)
			mp_bMultipleCriticalPaths = Value
		End Set
	End Property

	Public Property bNewTasksEffortDriven() As Boolean
		Get
			Return mp_bNewTasksEffortDriven
		End Get
		Set(ByVal Value As Boolean)
			mp_bNewTasksEffortDriven = Value
		End Set
	End Property

	Public Property bNewTasksEstimated() As Boolean
		Get
			Return mp_bNewTasksEstimated
		End Get
		Set(ByVal Value As Boolean)
			mp_bNewTasksEstimated = Value
		End Set
	End Property

	Public Property bSplitsInProgressTasks() As Boolean
		Get
			Return mp_bSplitsInProgressTasks
		End Get
		Set(ByVal Value As Boolean)
			mp_bSplitsInProgressTasks = Value
		End Set
	End Property

	Public Property bSpreadActualCost() As Boolean
		Get
			Return mp_bSpreadActualCost
		End Get
		Set(ByVal Value As Boolean)
			mp_bSpreadActualCost = Value
		End Set
	End Property

	Public Property bSpreadPercentComplete() As Boolean
		Get
			Return mp_bSpreadPercentComplete
		End Get
		Set(ByVal Value As Boolean)
			mp_bSpreadPercentComplete = Value
		End Set
	End Property

	Public Property bTaskUpdatesResource() As Boolean
		Get
			Return mp_bTaskUpdatesResource
		End Get
		Set(ByVal Value As Boolean)
			mp_bTaskUpdatesResource = Value
		End Set
	End Property

	Public Property bFiscalYearStart() As Boolean
		Get
			Return mp_bFiscalYearStart
		End Get
		Set(ByVal Value As Boolean)
			mp_bFiscalYearStart = Value
		End Set
	End Property

	Public Property yWeekStartDay() As E_WEEKSTARTDAY
		Get
			Return mp_yWeekStartDay
		End Get
		Set(ByVal Value As E_WEEKSTARTDAY)
			mp_yWeekStartDay = Value
		End Set
	End Property

	Public Property bMoveCompletedEndsBack() As Boolean
		Get
			Return mp_bMoveCompletedEndsBack
		End Get
		Set(ByVal Value As Boolean)
			mp_bMoveCompletedEndsBack = Value
		End Set
	End Property

	Public Property bMoveRemainingStartsBack() As Boolean
		Get
			Return mp_bMoveRemainingStartsBack
		End Get
		Set(ByVal Value As Boolean)
			mp_bMoveRemainingStartsBack = Value
		End Set
	End Property

	Public Property bMoveRemainingStartsForward() As Boolean
		Get
			Return mp_bMoveRemainingStartsForward
		End Get
		Set(ByVal Value As Boolean)
			mp_bMoveRemainingStartsForward = Value
		End Set
	End Property

	Public Property bMoveCompletedEndsForward() As Boolean
		Get
			Return mp_bMoveCompletedEndsForward
		End Get
		Set(ByVal Value As Boolean)
			mp_bMoveCompletedEndsForward = Value
		End Set
	End Property

	Public Property yBaselineForEarnedValue() As E_BASELINEFOREARNEDVALUE
		Get
			Return mp_yBaselineForEarnedValue
		End Get
		Set(ByVal Value As E_BASELINEFOREARNEDVALUE)
			mp_yBaselineForEarnedValue = Value
		End Set
	End Property

	Public Property bAutoAddNewResourcesAndTasks() As Boolean
		Get
			Return mp_bAutoAddNewResourcesAndTasks
		End Get
		Set(ByVal Value As Boolean)
			mp_bAutoAddNewResourcesAndTasks = Value
		End Set
	End Property

	Public Property dtStatusDate() As System.DateTime
		Get
			Return mp_dtStatusDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtStatusDate = Value
		End Set
	End Property

	Public Property dtCurrentDate() As System.DateTime
		Get
			Return mp_dtCurrentDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCurrentDate = Value
		End Set
	End Property

	Public Property bMicrosoftProjectServerURL() As Boolean
		Get
			Return mp_bMicrosoftProjectServerURL
		End Get
		Set(ByVal Value As Boolean)
			mp_bMicrosoftProjectServerURL = Value
		End Set
	End Property

	Public Property bAutolink() As Boolean
		Get
			Return mp_bAutolink
		End Get
		Set(ByVal Value As Boolean)
			mp_bAutolink = Value
		End Set
	End Property

	Public Property yNewTaskStartDate() As E_NEWTASKSTARTDATE
		Get
			Return mp_yNewTaskStartDate
		End Get
		Set(ByVal Value As E_NEWTASKSTARTDATE)
			mp_yNewTaskStartDate = Value
		End Set
	End Property

	Public Property yDefaultTaskEVMethod() As E_DEFAULTTASKEVMETHOD
		Get
			Return mp_yDefaultTaskEVMethod
		End Get
		Set(ByVal Value As E_DEFAULTTASKEVMETHOD)
			mp_yDefaultTaskEVMethod = Value
		End Set
	End Property

	Public Property bProjectExternallyEdited() As Boolean
		Get
			Return mp_bProjectExternallyEdited
		End Get
		Set(ByVal Value As Boolean)
			mp_bProjectExternallyEdited = Value
		End Set
	End Property

	Public Property dtExtendedCreationDate() As System.DateTime
		Get
			Return mp_dtExtendedCreationDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtExtendedCreationDate = Value
		End Set
	End Property

	Public Property bActualsInSync() As Boolean
		Get
			Return mp_bActualsInSync
		End Get
		Set(ByVal Value As Boolean)
			mp_bActualsInSync = Value
		End Set
	End Property

	Public Property bRemoveFileProperties() As Boolean
		Get
			Return mp_bRemoveFileProperties
		End Get
		Set(ByVal Value As Boolean)
			mp_bRemoveFileProperties = Value
		End Set
	End Property

	Public Property bAdminProject() As Boolean
		Get
			Return mp_bAdminProject
		End Get
		Set(ByVal Value As Boolean)
			mp_bAdminProject = Value
		End Set
	End Property

	Public Property sBaslineCalendar() As String
		Get
			Return mp_sBaslineCalendar
		End Get
		Set(ByVal Value As String)
			mp_sBaslineCalendar = Value
		End Set
	End Property

	Public Property bNewTasksAreManual() As Boolean
		Get
			Return mp_bNewTasksAreManual
		End Get
		Set(ByVal Value As Boolean)
			mp_bNewTasksAreManual = Value
		End Set
	End Property

	Public Property bUpdateManuallyScheduledTasksWhenEditingLinks() As Boolean
		Get
			Return mp_bUpdateManuallyScheduledTasksWhenEditingLinks
		End Get
		Set(ByVal Value As Boolean)
			mp_bUpdateManuallyScheduledTasksWhenEditingLinks = Value
		End Set
	End Property

	Public Property bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled() As Boolean
		Get
			Return mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled
		End Get
		Set(ByVal Value As Boolean)
			mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled = Value
		End Set
	End Property

	Public ReadOnly Property oOutlineCodes() As OutlineCodes
		Get
			Return mp_oOutlineCodes
		End Get
	End Property

	Public ReadOnly Property oWBSMasks() As WBSMasks
		Get
			Return mp_oWBSMasks
		End Get
	End Property

	Public ReadOnly Property oExtendedAttributes() As ExtendedAttributes
		Get
			Return mp_oExtendedAttributes
		End Get
	End Property

	Public ReadOnly Property oCalendars() As Calendars
		Get
			Return mp_oCalendars
		End Get
	End Property

	Public ReadOnly Property oTasks() As Tasks
		Get
			Return mp_oTasks
		End Get
	End Property

	Public ReadOnly Property oResources() As Resources
		Get
			Return mp_oResources
		End Get
	End Property

	Public ReadOnly Property oAssignments() As Assignments
		Get
			Return mp_oAssignments
		End Get
	End Property

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Project")
		oXML.SetXML(sXML)
		mp_ReadXML(oXML)
	End Sub

	Public Function GetXML() As String
		Dim oXML As New clsXML("Project")
		mp_WriteXML(oXML)
		Return oXML.GetXML
	End Function

	Public Function IsNull() As Boolean
		Dim bReturn As Boolean = True
		If mp_lSaveVersion <> 0 Then
			bReturn = False
		End If
		If mp_sUID <> "" Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_sTitle <> "" Then
			bReturn = False
		End If
		If mp_sSubject <> "" Then
			bReturn = False
		End If
		If mp_sCategory <> "" Then
			bReturn = False
		End If
		If mp_sCompany <> "" Then
			bReturn = False
		End If
		If mp_sManager <> "" Then
			bReturn = False
		End If
		If mp_sAuthor <> "" Then
			bReturn = False
		End If
		If mp_dtCreationDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_lRevision <> 0 Then
			bReturn = False
		End If
		If mp_dtLastSaved.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bScheduleFromStart <> true Then
			bReturn = False
		End If
		If mp_dtStartDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtFinishDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_yFYStartDate <> E_FYSTARTDATE.FYSD_JANUARY Then
			bReturn = False
		End If
		If mp_lCriticalSlackLimit <> 0 Then
			bReturn = False
		End If
		If mp_lCurrencyDigits <> 0 Then
			bReturn = False
		End If
		If mp_sCurrencySymbol <> "" Then
			bReturn = False
		End If
		If mp_sCurrencyCode <> "" Then
			bReturn = False
		End If
		If mp_yCurrencySymbolPosition <> E_CURRENCYSYMBOLPOSITION.CSP_BEFORE Then
			bReturn = False
		End If
		If mp_lCalendarUID <> 0 Then
			bReturn = False
		End If
		If mp_oDefaultStartTime.IsNull() = False Then
			bReturn = False
		End If
		If mp_oDefaultFinishTime.IsNull() = False Then
			bReturn = False
		End If
		If mp_lMinutesPerDay <> 0 Then
			bReturn = False
		End If
		If mp_lMinutesPerWeek <> 0 Then
			bReturn = False
		End If
		If mp_lDaysPerMonth <> 0 Then
			bReturn = False
		End If
		If mp_yDefaultTaskType <> E_DEFAULTTASKTYPE.DTT_FIXED_DURATION Then
			bReturn = False
		End If
		If mp_yDefaultFixedCostAccrual <> E_DEFAULTFIXEDCOSTACCRUAL.DFCA_START Then
			bReturn = False
		End If
		If mp_fDefaultStandardRate <> 0 Then
			bReturn = False
		End If
		If mp_fDefaultOvertimeRate <> 0 Then
			bReturn = False
		End If
		If mp_yDurationFormat <> E_DURATIONFORMAT.DF_M Then
			bReturn = False
		End If
		If mp_yWorkFormat <> E_WORKFORMAT.WF_M Then
			bReturn = False
		End If
		If mp_bEditableActualCosts <> false Then
			bReturn = False
		End If
		If mp_bHonorConstraints <> true Then
			bReturn = False
		End If
		If mp_yEarnedValueMethod <> E_EARNEDVALUEMETHOD.EVM_PERCENT_COMPLETE Then
			bReturn = False
		End If
		If mp_bInsertedProjectsLikeSummary <> true Then
			bReturn = False
		End If
		If mp_bMultipleCriticalPaths <> false Then
			bReturn = False
		End If
		If mp_bNewTasksEffortDriven <> true Then
			bReturn = False
		End If
		If mp_bNewTasksEstimated <> true Then
			bReturn = False
		End If
		If mp_bSplitsInProgressTasks <> true Then
			bReturn = False
		End If
		If mp_bSpreadActualCost <> true Then
			bReturn = False
		End If
		If mp_bSpreadPercentComplete <> false Then
			bReturn = False
		End If
		If mp_bTaskUpdatesResource <> False Then
			bReturn = False
		End If
		If mp_bFiscalYearStart <> False Then
			bReturn = False
		End If
		If mp_yWeekStartDay <> E_WEEKSTARTDAY.WSD_SUNDAY Then
			bReturn = False
		End If
		If mp_bMoveCompletedEndsBack <> false Then
			bReturn = False
		End If
		If mp_bMoveRemainingStartsBack <> false Then
			bReturn = False
		End If
		If mp_bMoveRemainingStartsForward <> false Then
			bReturn = False
		End If
		If mp_bMoveCompletedEndsForward <> false Then
			bReturn = False
		End If
		If mp_yBaselineForEarnedValue <> E_BASELINEFOREARNEDVALUE.BFEV_BASELINE Then
			bReturn = False
		End If
		If mp_bAutoAddNewResourcesAndTasks <> true Then
			bReturn = False
		End If
		If mp_dtStatusDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtCurrentDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bMicrosoftProjectServerURL <> False Then
			bReturn = False
		End If
		If mp_bAutolink <> False Then
			bReturn = False
		End If
		If mp_yNewTaskStartDate <> E_NEWTASKSTARTDATE.NTSD_PROJECT_START_DATE Then
			bReturn = False
		End If
		If mp_yDefaultTaskEVMethod <> E_DEFAULTTASKEVMETHOD.DTEVM_PERCENT_COMPLETE Then
			bReturn = False
		End If
		If mp_bProjectExternallyEdited <> False Then
			bReturn = False
		End If
		If mp_dtExtendedCreationDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bActualsInSync <> False Then
			bReturn = False
		End If
		If mp_bRemoveFileProperties <> False Then
			bReturn = False
		End If
		If mp_bAdminProject <> False Then
			bReturn = False
		End If
		If mp_sBaslineCalendar <> "" Then
			bReturn = False
		End If
		If mp_bNewTasksAreManual <> False Then
			bReturn = False
		End If
		If mp_bUpdateManuallyScheduledTasksWhenEditingLinks <> False Then
			bReturn = False
		End If
		If mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled <> False Then
			bReturn = False
		End If
		If mp_oOutlineCodes.IsNull() = False Then
			bReturn = False
		End If
		If mp_oWBSMasks.IsNull() = False Then
			bReturn = False
		End If
		If mp_oExtendedAttributes.IsNull() = False Then
			bReturn = False
		End If
		If mp_oCalendars.IsNull() = False Then
			bReturn = False
		End If
		If mp_oTasks.IsNull() = False Then
			bReturn = False
		End If
		If mp_oResources.IsNull() = False Then
			bReturn = False
		End If
		If mp_oAssignments.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Private Sub mp_WriteXML(ByRef oXML As clsXML)
		oXML.InitializeWriter()
		oXML.AddAttribute("xmlns", "http://schemas.microsoft.com/project")
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("SaveVersion", mp_lSaveVersion)
		If mp_sUID <> "" Then
			oXML.WriteProperty("UID", mp_sUID)
		End If
		If mp_sName <> "" Then
			oXML.WriteProperty("Name", mp_sName)
		End If
		If mp_sTitle <> "" Then
			oXML.WriteProperty("Title", mp_sTitle)
		End If
		If mp_sSubject <> "" Then
			oXML.WriteProperty("Subject", mp_sSubject)
		End If
		If mp_sCategory <> "" Then
			oXML.WriteProperty("Category", mp_sCategory)
		End If
		If mp_sCompany <> "" Then
			oXML.WriteProperty("Company", mp_sCompany)
		End If
		If mp_sManager <> "" Then
			oXML.WriteProperty("Manager", mp_sManager)
		End If
		If mp_sAuthor <> "" Then
			oXML.WriteProperty("Author", mp_sAuthor)
		End If
		If mp_dtCreationDate.Ticks <> 0 Then
			oXML.WriteProperty("CreationDate", mp_dtCreationDate)
		End If
		oXML.WriteProperty("Revision", mp_lRevision)
		If mp_dtLastSaved.Ticks <> 0 Then
			oXML.WriteProperty("LastSaved", mp_dtLastSaved)
		End If
		oXML.WriteProperty("ScheduleFromStart", mp_bScheduleFromStart)
		If mp_dtStartDate.Ticks <> 0 Then
			oXML.WriteProperty("StartDate", mp_dtStartDate)
		End If
		If mp_dtFinishDate.Ticks <> 0 Then
			oXML.WriteProperty("FinishDate", mp_dtFinishDate)
		End If
		oXML.WriteProperty("FYStartDate", mp_yFYStartDate)
		oXML.WriteProperty("CriticalSlackLimit", mp_lCriticalSlackLimit)
		oXML.WriteProperty("CurrencyDigits", mp_lCurrencyDigits)
		If mp_sCurrencySymbol <> "" Then
			oXML.WriteProperty("CurrencySymbol", mp_sCurrencySymbol)
		End If
		oXML.WriteProperty("CurrencyCode", mp_sCurrencyCode)
		oXML.WriteProperty("CurrencySymbolPosition", mp_yCurrencySymbolPosition)
		oXML.WriteProperty("CalendarUID", mp_lCalendarUID)
		If mp_oDefaultStartTime.IsNull() = False Then
			oXML.WriteProperty("DefaultStartTime", mp_oDefaultStartTime)
		End If
		If mp_oDefaultFinishTime.IsNull() = False Then
			oXML.WriteProperty("DefaultFinishTime", mp_oDefaultFinishTime)
		End If
		oXML.WriteProperty("MinutesPerDay", mp_lMinutesPerDay)
		oXML.WriteProperty("MinutesPerWeek", mp_lMinutesPerWeek)
		oXML.WriteProperty("DaysPerMonth", mp_lDaysPerMonth)
		oXML.WriteProperty("DefaultTaskType", mp_yDefaultTaskType)
		oXML.WriteProperty("DefaultFixedCostAccrual", mp_yDefaultFixedCostAccrual)
		oXML.WriteProperty("DefaultStandardRate", mp_fDefaultStandardRate)
		oXML.WriteProperty("DefaultOvertimeRate", mp_fDefaultOvertimeRate)
		oXML.WriteProperty("DurationFormat", mp_yDurationFormat)
		oXML.WriteProperty("WorkFormat", mp_yWorkFormat)
		oXML.WriteProperty("EditableActualCosts", mp_bEditableActualCosts)
		oXML.WriteProperty("HonorConstraints", mp_bHonorConstraints)
		oXML.WriteProperty("EarnedValueMethod", mp_yEarnedValueMethod)
		oXML.WriteProperty("InsertedProjectsLikeSummary", mp_bInsertedProjectsLikeSummary)
		oXML.WriteProperty("MultipleCriticalPaths", mp_bMultipleCriticalPaths)
		oXML.WriteProperty("NewTasksEffortDriven", mp_bNewTasksEffortDriven)
		oXML.WriteProperty("NewTasksEstimated", mp_bNewTasksEstimated)
		oXML.WriteProperty("SplitsInProgressTasks", mp_bSplitsInProgressTasks)
		oXML.WriteProperty("SpreadActualCost", mp_bSpreadActualCost)
		oXML.WriteProperty("SpreadPercentComplete", mp_bSpreadPercentComplete)
		oXML.WriteProperty("TaskUpdatesResource", mp_bTaskUpdatesResource)
		oXML.WriteProperty("FiscalYearStart", mp_bFiscalYearStart)
		oXML.WriteProperty("WeekStartDay", mp_yWeekStartDay)
		oXML.WriteProperty("MoveCompletedEndsBack", mp_bMoveCompletedEndsBack)
		oXML.WriteProperty("MoveRemainingStartsBack", mp_bMoveRemainingStartsBack)
		oXML.WriteProperty("MoveRemainingStartsForward", mp_bMoveRemainingStartsForward)
		oXML.WriteProperty("MoveCompletedEndsForward", mp_bMoveCompletedEndsForward)
		oXML.WriteProperty("BaselineForEarnedValue", mp_yBaselineForEarnedValue)
		oXML.WriteProperty("AutoAddNewResourcesAndTasks", mp_bAutoAddNewResourcesAndTasks)
		If mp_dtStatusDate.Ticks <> 0 Then
			oXML.WriteProperty("StatusDate", mp_dtStatusDate)
		End If
		If mp_dtCurrentDate.Ticks <> 0 Then
			oXML.WriteProperty("CurrentDate", mp_dtCurrentDate)
		End If
		oXML.WriteProperty("MicrosoftProjectServerURL", mp_bMicrosoftProjectServerURL)
		oXML.WriteProperty("Autolink", mp_bAutolink)
		oXML.WriteProperty("NewTaskStartDate", mp_yNewTaskStartDate)
		oXML.WriteProperty("DefaultTaskEVMethod", mp_yDefaultTaskEVMethod)
		oXML.WriteProperty("ProjectExternallyEdited", mp_bProjectExternallyEdited)
		If mp_dtExtendedCreationDate.Ticks <> 0 Then
			oXML.WriteProperty("ExtendedCreationDate", mp_dtExtendedCreationDate)
		End If
		oXML.WriteProperty("ActualsInSync", mp_bActualsInSync)
		oXML.WriteProperty("RemoveFileProperties", mp_bRemoveFileProperties)
		oXML.WriteProperty("AdminProject", mp_bAdminProject)
		If mp_sBaslineCalendar <> "" Then
			oXML.WriteProperty("BaslineCalendar", mp_sBaslineCalendar)
		End If
		oXML.WriteProperty("NewTasksAreManual", mp_bNewTasksAreManual)
		oXML.WriteProperty("UpdateManuallyScheduledTasksWhenEditingLinks", mp_bUpdateManuallyScheduledTasksWhenEditingLinks)
		oXML.WriteProperty("KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled)
		oXML.WriteObject(mp_oOutlineCodes.GetXML())
		oXML.WriteObject(mp_oWBSMasks.GetXML())
		oXML.WriteObject(mp_oExtendedAttributes.GetXML())
		oXML.WriteObject(mp_oCalendars.GetXML())
		oXML.WriteObject(mp_oTasks.GetXML())
		oXML.WriteObject(mp_oResources.GetXML())
		oXML.WriteObject(mp_oAssignments.GetXML())
	End Sub

	Private Sub mp_ReadXML(ByRef oXML As clsXML)
		oXML.SupportOptional = True
		oXML.InitializeReader()
		oXML.ReadProperty("SaveVersion", mp_lSaveVersion)
		oXML.ReadProperty("UID", mp_sUID)
		If mp_sUID.Length > 16 Then
			mp_sUID = mp_sUID.Substring(0, 16)
		End If
		oXML.ReadProperty("Name", mp_sName)
		If mp_sName.Length > 255 Then
			mp_sName = mp_sName.Substring(0, 255)
		End If
		oXML.ReadProperty("Title", mp_sTitle)
		If mp_sTitle.Length > 512 Then
			mp_sTitle = mp_sTitle.Substring(0, 512)
		End If
		oXML.ReadProperty("Subject", mp_sSubject)
		If mp_sSubject.Length > 512 Then
			mp_sSubject = mp_sSubject.Substring(0, 512)
		End If
		oXML.ReadProperty("Category", mp_sCategory)
		If mp_sCategory.Length > 512 Then
			mp_sCategory = mp_sCategory.Substring(0, 512)
		End If
		oXML.ReadProperty("Company", mp_sCompany)
		If mp_sCompany.Length > 512 Then
			mp_sCompany = mp_sCompany.Substring(0, 512)
		End If
		oXML.ReadProperty("Manager", mp_sManager)
		If mp_sManager.Length > 512 Then
			mp_sManager = mp_sManager.Substring(0, 512)
		End If
		oXML.ReadProperty("Author", mp_sAuthor)
		If mp_sAuthor.Length > 512 Then
			mp_sAuthor = mp_sAuthor.Substring(0, 512)
		End If
		oXML.ReadProperty("CreationDate", mp_dtCreationDate)
		oXML.ReadProperty("Revision", mp_lRevision)
		oXML.ReadProperty("LastSaved", mp_dtLastSaved)
		oXML.ReadProperty("ScheduleFromStart", mp_bScheduleFromStart)
		oXML.ReadProperty("StartDate", mp_dtStartDate)
		oXML.ReadProperty("FinishDate", mp_dtFinishDate)
		oXML.ReadProperty("FYStartDate", mp_yFYStartDate)
		oXML.ReadProperty("CriticalSlackLimit", mp_lCriticalSlackLimit)
		oXML.ReadProperty("CurrencyDigits", mp_lCurrencyDigits)
		oXML.ReadProperty("CurrencySymbol", mp_sCurrencySymbol)
		If mp_sCurrencySymbol.Length > 20 Then
			mp_sCurrencySymbol = mp_sCurrencySymbol.Substring(0, 20)
		End If
		oXML.ReadProperty("CurrencyCode", mp_sCurrencyCode)
		If mp_sCurrencyCode.Length > 3 Then
			mp_sCurrencyCode = mp_sCurrencyCode.Substring(0, 3)
		End If
		oXML.ReadProperty("CurrencySymbolPosition", mp_yCurrencySymbolPosition)
		oXML.ReadProperty("CalendarUID", mp_lCalendarUID)
		oXML.ReadProperty("DefaultStartTime", mp_oDefaultStartTime)
		oXML.ReadProperty("DefaultFinishTime", mp_oDefaultFinishTime)
		oXML.ReadProperty("MinutesPerDay", mp_lMinutesPerDay)
		oXML.ReadProperty("MinutesPerWeek", mp_lMinutesPerWeek)
		oXML.ReadProperty("DaysPerMonth", mp_lDaysPerMonth)
		oXML.ReadProperty("DefaultTaskType", mp_yDefaultTaskType)
		oXML.ReadProperty("DefaultFixedCostAccrual", mp_yDefaultFixedCostAccrual)
		oXML.ReadProperty("DefaultStandardRate", mp_fDefaultStandardRate)
		oXML.ReadProperty("DefaultOvertimeRate", mp_fDefaultOvertimeRate)
		oXML.ReadProperty("DurationFormat", mp_yDurationFormat)
		oXML.ReadProperty("WorkFormat", mp_yWorkFormat)
		oXML.ReadProperty("EditableActualCosts", mp_bEditableActualCosts)
		oXML.ReadProperty("HonorConstraints", mp_bHonorConstraints)
		oXML.ReadProperty("EarnedValueMethod", mp_yEarnedValueMethod)
		oXML.ReadProperty("InsertedProjectsLikeSummary", mp_bInsertedProjectsLikeSummary)
		oXML.ReadProperty("MultipleCriticalPaths", mp_bMultipleCriticalPaths)
		oXML.ReadProperty("NewTasksEffortDriven", mp_bNewTasksEffortDriven)
		oXML.ReadProperty("NewTasksEstimated", mp_bNewTasksEstimated)
		oXML.ReadProperty("SplitsInProgressTasks", mp_bSplitsInProgressTasks)
		oXML.ReadProperty("SpreadActualCost", mp_bSpreadActualCost)
		oXML.ReadProperty("SpreadPercentComplete", mp_bSpreadPercentComplete)
		oXML.ReadProperty("TaskUpdatesResource", mp_bTaskUpdatesResource)
		oXML.ReadProperty("FiscalYearStart", mp_bFiscalYearStart)
		oXML.ReadProperty("WeekStartDay", mp_yWeekStartDay)
		oXML.ReadProperty("MoveCompletedEndsBack", mp_bMoveCompletedEndsBack)
		oXML.ReadProperty("MoveRemainingStartsBack", mp_bMoveRemainingStartsBack)
		oXML.ReadProperty("MoveRemainingStartsForward", mp_bMoveRemainingStartsForward)
		oXML.ReadProperty("MoveCompletedEndsForward", mp_bMoveCompletedEndsForward)
		oXML.ReadProperty("BaselineForEarnedValue", mp_yBaselineForEarnedValue)
		oXML.ReadProperty("AutoAddNewResourcesAndTasks", mp_bAutoAddNewResourcesAndTasks)
		oXML.ReadProperty("StatusDate", mp_dtStatusDate)
		oXML.ReadProperty("CurrentDate", mp_dtCurrentDate)
		oXML.ReadProperty("MicrosoftProjectServerURL", mp_bMicrosoftProjectServerURL)
		oXML.ReadProperty("Autolink", mp_bAutolink)
		oXML.ReadProperty("NewTaskStartDate", mp_yNewTaskStartDate)
		oXML.ReadProperty("DefaultTaskEVMethod", mp_yDefaultTaskEVMethod)
		oXML.ReadProperty("ProjectExternallyEdited", mp_bProjectExternallyEdited)
		oXML.ReadProperty("ExtendedCreationDate", mp_dtExtendedCreationDate)
		oXML.ReadProperty("ActualsInSync", mp_bActualsInSync)
		oXML.ReadProperty("RemoveFileProperties", mp_bRemoveFileProperties)
		oXML.ReadProperty("AdminProject", mp_bAdminProject)
		oXML.ReadProperty("BaslineCalendar", mp_sBaslineCalendar)
		oXML.ReadProperty("NewTasksAreManual", mp_bNewTasksAreManual)
		oXML.ReadProperty("UpdateManuallyScheduledTasksWhenEditingLinks", mp_bUpdateManuallyScheduledTasksWhenEditingLinks)
		oXML.ReadProperty("KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", mp_bKeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled)
		mp_oOutlineCodes.SetXML(oXML.ReadObject("OutlineCodes"))
		mp_oWBSMasks.SetXML(oXML.ReadObject("WBSMasks"))
		mp_oExtendedAttributes.SetXML(oXML.ReadObject("ExtendedAttributes"))
		mp_oCalendars.SetXML(oXML.ReadObject("Calendars"))
		mp_oTasks.SetXML(oXML.ReadObject("Tasks"))
		mp_oResources.SetXML(oXML.ReadObject("Resources"))
		mp_oAssignments.SetXML(oXML.ReadObject("Assignments"))
	End Sub

End Class
