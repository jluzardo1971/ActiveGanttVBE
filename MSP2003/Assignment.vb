Option Explicit On

Public Class Assignment
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_lTaskUID As Integer
	Private mp_lResourceUID As Integer
	Private mp_lPercentWorkComplete As Integer
	Private mp_cActualCost As Decimal
	Private mp_dtActualFinish As System.DateTime
	Private mp_cActualOvertimeCost As Decimal
	Private mp_oActualOvertimeWork As Duration
	Private mp_dtActualStart As System.DateTime
	Private mp_oActualWork As Duration
	Private mp_fACWP As Single
	Private mp_bConfirmed As Boolean
	Private mp_cCost As Decimal
	Private mp_yCostRateTable As E_COSTRATETABLE
	Private mp_fCostVariance As Single
	Private mp_fCV As Single
	Private mp_lDelay As Integer
	Private mp_dtFinish As System.DateTime
	Private mp_lFinishVariance As Integer
	Private mp_sHyperlink As String
	Private mp_sHyperlinkAddress As String
	Private mp_sHyperlinkSubAddress As String
	Private mp_fWorkVariance As Single
	Private mp_bHasFixedRateUnits As Boolean
	Private mp_bFixedMaterial As Boolean
	Private mp_lLevelingDelay As Integer
	Private mp_yLevelingDelayFormat As E_LEVELINGDELAYFORMAT
	Private mp_bLinkedFields As Boolean
	Private mp_bMilestone As Boolean
	Private mp_sNotes As String
	Private mp_bOverallocated As Boolean
	Private mp_cOvertimeCost As Decimal
	Private mp_oOvertimeWork As Duration
	Private mp_oRegularWork As Duration
	Private mp_cRemainingCost As Decimal
	Private mp_cRemainingOvertimeCost As Decimal
	Private mp_oRemainingOvertimeWork As Duration
	Private mp_oRemainingWork As Duration
	Private mp_bResponsePending As Boolean
	Private mp_dtStart As System.DateTime
	Private mp_dtStop As System.DateTime
	Private mp_dtResume As System.DateTime
	Private mp_lStartVariance As Integer
	Private mp_fUnits As Single
	Private mp_bUpdateNeeded As Boolean
	Private mp_fVAC As Single
	Private mp_oWork As Duration
	Private mp_yWorkContour As E_WORKCONTOUR
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single
	Private mp_yBookingType As E_BOOKINGTYPE
	Private mp_oActualWorkProtected As Duration
	Private mp_oActualOvertimeWorkProtected As Duration
	Private mp_dtCreationDate As System.DateTime
	Private mp_oExtendedAttribute_C As AssignmentExtendedAttribute_C
	Private mp_oBaseline_C As AssignmentBaseline_C
	Private mp_oTimephasedData_C As TimephasedData_C

	Public Sub New()
		mp_lUID = 0
		mp_lTaskUID = 0
		mp_lResourceUID = 0
		mp_lPercentWorkComplete = 0
		mp_cActualCost = 0
		mp_dtActualFinish = New System.DateTime(0)
		mp_cActualOvertimeCost = 0
		mp_oActualOvertimeWork = New Duration()
		mp_dtActualStart = New System.DateTime(0)
		mp_oActualWork = New Duration()
		mp_fACWP = 0
		mp_bConfirmed = False
		mp_cCost = 0
		mp_yCostRateTable = E_COSTRATETABLE.CRT_COST_RATE_TABLE_0
		mp_fCostVariance = 0
		mp_fCV = 0
		mp_lDelay = 0
		mp_dtFinish = New System.DateTime(0)
		mp_lFinishVariance = 0
		mp_sHyperlink = ""
		mp_sHyperlinkAddress = ""
		mp_sHyperlinkSubAddress = ""
		mp_fWorkVariance = 0
		mp_bHasFixedRateUnits = False
		mp_bFixedMaterial = False
		mp_lLevelingDelay = 0
		mp_yLevelingDelayFormat = E_LEVELINGDELAYFORMAT.LDF_M
		mp_bLinkedFields = False
		mp_bMilestone = False
		mp_sNotes = ""
		mp_bOverallocated = False
		mp_cOvertimeCost = 0
		mp_oOvertimeWork = New Duration()
		mp_oRegularWork = New Duration()
		mp_cRemainingCost = 0
		mp_cRemainingOvertimeCost = 0
		mp_oRemainingOvertimeWork = New Duration()
		mp_oRemainingWork = New Duration()
		mp_bResponsePending = False
		mp_dtStart = New System.DateTime(0)
		mp_dtStop = New System.DateTime(0)
		mp_dtResume = New System.DateTime(0)
		mp_lStartVariance = 0
		mp_fUnits = 0
		mp_bUpdateNeeded = False
		mp_fVAC = 0
		mp_oWork = New Duration()
		mp_yWorkContour = E_WORKCONTOUR.WC_FLAT
		mp_fBCWS = 0
		mp_fBCWP = 0
		mp_yBookingType = E_BOOKINGTYPE.BT_COMMITED
		mp_oActualWorkProtected = New Duration()
		mp_oActualOvertimeWorkProtected = New Duration()
		mp_dtCreationDate = New System.DateTime(0)
		mp_oExtendedAttribute_C = New AssignmentExtendedAttribute_C()
		mp_oBaseline_C = New AssignmentBaseline_C()
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

	Public Property lTaskUID() As Integer
		Get
			Return mp_lTaskUID
		End Get
		Set(ByVal Value As Integer)
			mp_lTaskUID = Value
		End Set
	End Property

	Public Property lResourceUID() As Integer
		Get
			Return mp_lResourceUID
		End Get
		Set(ByVal Value As Integer)
			mp_lResourceUID = Value
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

	Public Property cActualCost() As Decimal
		Get
			Return mp_cActualCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cActualCost = Value
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

	Public Property cActualOvertimeCost() As Decimal
		Get
			Return mp_cActualOvertimeCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cActualOvertimeCost = Value
		End Set
	End Property

	Public ReadOnly Property oActualOvertimeWork() As Duration
		Get
			Return mp_oActualOvertimeWork
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

	Public ReadOnly Property oActualWork() As Duration
		Get
			Return mp_oActualWork
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

	Public Property bConfirmed() As Boolean
		Get
			Return mp_bConfirmed
		End Get
		Set(ByVal Value As Boolean)
			mp_bConfirmed = Value
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

	Public Property yCostRateTable() As E_COSTRATETABLE
		Get
			Return mp_yCostRateTable
		End Get
		Set(ByVal Value As E_COSTRATETABLE)
			mp_yCostRateTable = Value
		End Set
	End Property

	Public Property fCostVariance() As Single
		Get
			Return mp_fCostVariance
		End Get
		Set(ByVal Value As Single)
			mp_fCostVariance = Value
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

	Public Property lDelay() As Integer
		Get
			Return mp_lDelay
		End Get
		Set(ByVal Value As Integer)
			mp_lDelay = Value
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

	Public Property lFinishVariance() As Integer
		Get
			Return mp_lFinishVariance
		End Get
		Set(ByVal Value As Integer)
			mp_lFinishVariance = Value
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

	Public Property fWorkVariance() As Single
		Get
			Return mp_fWorkVariance
		End Get
		Set(ByVal Value As Single)
			mp_fWorkVariance = Value
		End Set
	End Property

	Public Property bHasFixedRateUnits() As Boolean
		Get
			Return mp_bHasFixedRateUnits
		End Get
		Set(ByVal Value As Boolean)
			mp_bHasFixedRateUnits = Value
		End Set
	End Property

	Public Property bFixedMaterial() As Boolean
		Get
			Return mp_bFixedMaterial
		End Get
		Set(ByVal Value As Boolean)
			mp_bFixedMaterial = Value
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

	Public Property bLinkedFields() As Boolean
		Get
			Return mp_bLinkedFields
		End Get
		Set(ByVal Value As Boolean)
			mp_bLinkedFields = Value
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

	Public Property sNotes() As String
		Get
			Return mp_sNotes
		End Get
		Set(ByVal Value As String)
			mp_sNotes = Value
		End Set
	End Property

	Public Property bOverallocated() As Boolean
		Get
			Return mp_bOverallocated
		End Get
		Set(ByVal Value As Boolean)
			mp_bOverallocated = Value
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

	Public ReadOnly Property oRegularWork() As Duration
		Get
			Return mp_oRegularWork
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

	Public ReadOnly Property oRemainingWork() As Duration
		Get
			Return mp_oRemainingWork
		End Get
	End Property

	Public Property bResponsePending() As Boolean
		Get
			Return mp_bResponsePending
		End Get
		Set(ByVal Value As Boolean)
			mp_bResponsePending = Value
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

	Public Property lStartVariance() As Integer
		Get
			Return mp_lStartVariance
		End Get
		Set(ByVal Value As Integer)
			mp_lStartVariance = Value
		End Set
	End Property

	Public Property fUnits() As Single
		Get
			Return mp_fUnits
		End Get
		Set(ByVal Value As Single)
			mp_fUnits = Value
		End Set
	End Property

	Public Property bUpdateNeeded() As Boolean
		Get
			Return mp_bUpdateNeeded
		End Get
		Set(ByVal Value As Boolean)
			mp_bUpdateNeeded = Value
		End Set
	End Property

	Public Property fVAC() As Single
		Get
			Return mp_fVAC
		End Get
		Set(ByVal Value As Single)
			mp_fVAC = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public Property yWorkContour() As E_WORKCONTOUR
		Get
			Return mp_yWorkContour
		End Get
		Set(ByVal Value As E_WORKCONTOUR)
			mp_yWorkContour = Value
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

	Public Property yBookingType() As E_BOOKINGTYPE
		Get
			Return mp_yBookingType
		End Get
		Set(ByVal Value As E_BOOKINGTYPE)
			mp_yBookingType = Value
		End Set
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

	Public Property dtCreationDate() As System.DateTime
		Get
			Return mp_dtCreationDate
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtCreationDate = Value
		End Set
	End Property

	Public ReadOnly Property oExtendedAttribute_C() As AssignmentExtendedAttribute_C
		Get
			Return mp_oExtendedAttribute_C
		End Get
	End Property

	Public ReadOnly Property oBaseline_C() As AssignmentBaseline_C
		Get
			Return mp_oBaseline_C
		End Get
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
		If mp_lTaskUID <> 0 Then
			bReturn = False
		End If
		If mp_lResourceUID <> 0 Then
			bReturn = False
		End If
		If mp_lPercentWorkComplete <> 0 Then
			bReturn = False
		End If
		If mp_cActualCost <> 0 Then
			bReturn = False
		End If
		If mp_dtActualFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_cActualOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oActualOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_dtActualStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_oActualWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_fACWP <> 0 Then
			bReturn = False
		End If
		If mp_bConfirmed <> False Then
			bReturn = False
		End If
		If mp_cCost <> 0 Then
			bReturn = False
		End If
		If mp_yCostRateTable <> E_COSTRATETABLE.CRT_COST_RATE_TABLE_0 Then
			bReturn = False
		End If
		If mp_fCostVariance <> 0 Then
			bReturn = False
		End If
		If mp_fCV <> 0 Then
			bReturn = False
		End If
		If mp_lDelay <> 0 Then
			bReturn = False
		End If
		If mp_dtFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_lFinishVariance <> 0 Then
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
		If mp_fWorkVariance <> 0 Then
			bReturn = False
		End If
		If mp_bHasFixedRateUnits <> False Then
			bReturn = False
		End If
		If mp_bFixedMaterial <> False Then
			bReturn = False
		End If
		If mp_lLevelingDelay <> 0 Then
			bReturn = False
		End If
		If mp_yLevelingDelayFormat <> E_LEVELINGDELAYFORMAT.LDF_M Then
			bReturn = False
		End If
		If mp_bLinkedFields <> False Then
			bReturn = False
		End If
		If mp_bMilestone <> False Then
			bReturn = False
		End If
		If mp_sNotes <> "" Then
			bReturn = False
		End If
		If mp_bOverallocated <> False Then
			bReturn = False
		End If
		If mp_cOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRegularWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_cRemainingCost <> 0 Then
			bReturn = False
		End If
		If mp_cRemainingOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_oRemainingOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRemainingWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_bResponsePending <> False Then
			bReturn = False
		End If
		If mp_dtStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtStop.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtResume.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_lStartVariance <> 0 Then
			bReturn = False
		End If
		If mp_fUnits <> 0 Then
			bReturn = False
		End If
		If mp_bUpdateNeeded <> False Then
			bReturn = False
		End If
		If mp_fVAC <> 0 Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_yWorkContour <> E_WORKCONTOUR.WC_FLAT Then
			bReturn = False
		End If
		If mp_fBCWS <> 0 Then
			bReturn = False
		End If
		If mp_fBCWP <> 0 Then
			bReturn = False
		End If
		If mp_yBookingType <> E_BOOKINGTYPE.BT_COMMITED Then
			bReturn = False
		End If
		If mp_oActualWorkProtected.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualOvertimeWorkProtected.IsNull() = False Then
			bReturn = False
		End If
		If mp_dtCreationDate.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oBaseline_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Assignment/>"
		End if
		Dim oXML As New clsXML("Assignment")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("UID", mp_lUID)
		oXML.WriteProperty("TaskUID", mp_lTaskUID)
		oXML.WriteProperty("ResourceUID", mp_lResourceUID)
		oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.WriteProperty("ActualCost", mp_cActualCost)
		If mp_dtActualFinish.Ticks <> 0 Then
			oXML.WriteProperty("ActualFinish", mp_dtActualFinish)
		End If
		oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		If mp_dtActualStart.Ticks <> 0 Then
			oXML.WriteProperty("ActualStart", mp_dtActualStart)
		End If
		oXML.WriteProperty("ActualWork", mp_oActualWork)
		oXML.WriteProperty("ACWP", mp_fACWP)
		oXML.WriteProperty("Confirmed", mp_bConfirmed)
		oXML.WriteProperty("Cost", mp_cCost)
		oXML.WriteProperty("CostRateTable", mp_yCostRateTable)
		oXML.WriteProperty("CostVariance", mp_fCostVariance)
		oXML.WriteProperty("CV", mp_fCV)
		oXML.WriteProperty("Delay", mp_lDelay)
		If mp_dtFinish.Ticks <> 0 Then
			oXML.WriteProperty("Finish", mp_dtFinish)
		End If
		oXML.WriteProperty("FinishVariance", mp_lFinishVariance)
		If mp_sHyperlink <> "" Then
			oXML.WriteProperty("Hyperlink", mp_sHyperlink)
		End If
		If mp_sHyperlinkAddress <> "" Then
			oXML.WriteProperty("HyperlinkAddress", mp_sHyperlinkAddress)
		End If
		If mp_sHyperlinkSubAddress <> "" Then
			oXML.WriteProperty("HyperlinkSubAddress", mp_sHyperlinkSubAddress)
		End If
		oXML.WriteProperty("WorkVariance", mp_fWorkVariance)
		oXML.WriteProperty("HasFixedRateUnits", mp_bHasFixedRateUnits)
		oXML.WriteProperty("FixedMaterial", mp_bFixedMaterial)
		oXML.WriteProperty("LevelingDelay", mp_lLevelingDelay)
		oXML.WriteProperty("LevelingDelayFormat", mp_yLevelingDelayFormat)
		oXML.WriteProperty("LinkedFields", mp_bLinkedFields)
		oXML.WriteProperty("Milestone", mp_bMilestone)
		If mp_sNotes <> "" Then
			oXML.WriteProperty("Notes", mp_sNotes)
		End If
		oXML.WriteProperty("Overallocated", mp_bOverallocated)
		oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork)
		oXML.WriteProperty("RegularWork", mp_oRegularWork)
		oXML.WriteProperty("RemainingCost", mp_cRemainingCost)
		oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.WriteProperty("RemainingWork", mp_oRemainingWork)
		oXML.WriteProperty("ResponsePending", mp_bResponsePending)
		If mp_dtStart.Ticks <> 0 Then
			oXML.WriteProperty("Start", mp_dtStart)
		End If
		If mp_dtStop.Ticks <> 0 Then
			oXML.WriteProperty("Stop", mp_dtStop)
		End If
		If mp_dtResume.Ticks <> 0 Then
			oXML.WriteProperty("Resume", mp_dtResume)
		End If
		oXML.WriteProperty("StartVariance", mp_lStartVariance)
		oXML.WriteProperty("Units", mp_fUnits)
		oXML.WriteProperty("UpdateNeeded", mp_bUpdateNeeded)
		oXML.WriteProperty("VAC", mp_fVAC)
		oXML.WriteProperty("Work", mp_oWork)
		oXML.WriteProperty("WorkContour", mp_yWorkContour)
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		oXML.WriteProperty("BookingType", mp_yBookingType)
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		If mp_dtCreationDate.Ticks <> 0 Then
			oXML.WriteProperty("CreationDate", mp_dtCreationDate)
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			mp_oExtendedAttribute_C.WriteObjectProtected(oXML)
		End If
		If mp_oBaseline_C.IsNull() = False Then
			mp_oBaseline_C.WriteObjectProtected(oXML)
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			mp_oTimephasedData_C.WriteObjectProtected(oXML)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Assignment")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("UID", mp_lUID)
		oXML.ReadProperty("TaskUID", mp_lTaskUID)
		oXML.ReadProperty("ResourceUID", mp_lResourceUID)
		oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.ReadProperty("ActualCost", mp_cActualCost)
		oXML.ReadProperty("ActualFinish", mp_dtActualFinish)
		oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		oXML.ReadProperty("ActualStart", mp_dtActualStart)
		oXML.ReadProperty("ActualWork", mp_oActualWork)
		oXML.ReadProperty("ACWP", mp_fACWP)
		oXML.ReadProperty("Confirmed", mp_bConfirmed)
		oXML.ReadProperty("Cost", mp_cCost)
		oXML.ReadProperty("CostRateTable", mp_yCostRateTable)
		oXML.ReadProperty("CostVariance", mp_fCostVariance)
		oXML.ReadProperty("CV", mp_fCV)
		oXML.ReadProperty("Delay", mp_lDelay)
		oXML.ReadProperty("Finish", mp_dtFinish)
		oXML.ReadProperty("FinishVariance", mp_lFinishVariance)
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
		oXML.ReadProperty("WorkVariance", mp_fWorkVariance)
		oXML.ReadProperty("HasFixedRateUnits", mp_bHasFixedRateUnits)
		oXML.ReadProperty("FixedMaterial", mp_bFixedMaterial)
		oXML.ReadProperty("LevelingDelay", mp_lLevelingDelay)
		oXML.ReadProperty("LevelingDelayFormat", mp_yLevelingDelayFormat)
		oXML.ReadProperty("LinkedFields", mp_bLinkedFields)
		oXML.ReadProperty("Milestone", mp_bMilestone)
		oXML.ReadProperty("Notes", mp_sNotes)
		oXML.ReadProperty("Overallocated", mp_bOverallocated)
		oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork)
		oXML.ReadProperty("RegularWork", mp_oRegularWork)
		oXML.ReadProperty("RemainingCost", mp_cRemainingCost)
		oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.ReadProperty("RemainingWork", mp_oRemainingWork)
		oXML.ReadProperty("ResponsePending", mp_bResponsePending)
		oXML.ReadProperty("Start", mp_dtStart)
		oXML.ReadProperty("Stop", mp_dtStop)
		oXML.ReadProperty("Resume", mp_dtResume)
		oXML.ReadProperty("StartVariance", mp_lStartVariance)
		oXML.ReadProperty("Units", mp_fUnits)
		oXML.ReadProperty("UpdateNeeded", mp_bUpdateNeeded)
		oXML.ReadProperty("VAC", mp_fVAC)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("WorkContour", mp_yWorkContour)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
		oXML.ReadProperty("BookingType", mp_yBookingType)
		oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		oXML.ReadProperty("CreationDate", mp_dtCreationDate)
		mp_oExtendedAttribute_C.ReadObjectProtected(oXML)
		mp_oBaseline_C.ReadObjectProtected(oXML)
		mp_oTimephasedData_C.ReadObjectProtected(oXML)
	End Sub

End Class
