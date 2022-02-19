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
	Private mp_fPeakUnits As Single
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
	Private mp_bSummary As Boolean
	Private mp_fSV As Single
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
	Private mp_sAssnOwner As String
	Private mp_sAssnOwnerGuid As String
	Private mp_cBudgetCost As Decimal
	Private mp_oBudgetWork As Duration
	Private mp_oExtendedAttribute_C As AssignmentExtendedAttribute_C
	Private mp_oBaseline_C As AssignmentBaseline_C
	Private mp_sf404000 As String
	Private mp_sf404001 As String
	Private mp_sf404002 As String
	Private mp_sf404003 As String
	Private mp_sf404004 As String
	Private mp_sf404005 As String
	Private mp_sf404006 As String
	Private mp_sf404007 As String
	Private mp_sf404008 As String
	Private mp_sf404009 As String
	Private mp_sf40400a As String
	Private mp_sf40400b As String
	Private mp_sf40400c As String
	Private mp_sf40400d As String
	Private mp_sf40400e As String
	Private mp_sf40400f As String
	Private mp_sf404010 As String
	Private mp_sf404011 As String
	Private mp_sf404012 As String
	Private mp_sf404013 As String
	Private mp_sf404014 As String
	Private mp_sf404015 As String
	Private mp_sf404016 As String
	Private mp_sf404017 As String
	Private mp_sf404018 As String
	Private mp_sf404019 As String
	Private mp_sf40401a As String
	Private mp_sf40401b As String
	Private mp_sf40401c As String
	Private mp_sf40401d As String
	Private mp_sf40401e As String
	Private mp_sf40401f As String
	Private mp_sf404020 As String
	Private mp_sf404021 As String
	Private mp_sf404022 As String
	Private mp_sf404023 As String
	Private mp_sf404024 As String
	Private mp_sf404025 As String
	Private mp_sf404026 As String
	Private mp_sf404027 As String
	Private mp_sf404028 As String
	Private mp_sf404029 As String
	Private mp_sf40402a As String
	Private mp_sf40402b As String
	Private mp_sf40402c As String
	Private mp_sf40402d As String
	Private mp_sf40402e As String
	Private mp_sf40402f As String
	Private mp_sf404030 As String
	Private mp_sf404031 As String
	Private mp_sf404032 As String
	Private mp_sf404033 As String
	Private mp_sf404034 As String
	Private mp_sf404035 As String
	Private mp_sf404036 As String
	Private mp_sf404037 As String
	Private mp_sf404038 As String
	Private mp_sf404039 As String
	Private mp_sf40403a As String
	Private mp_sf40403b As String
	Private mp_sf40403c As String
	Private mp_sf40403d As String
	Private mp_sf40403e As String
	Private mp_sf40403f As String
	Private mp_sf404040 As String
	Private mp_sf404041 As String
	Private mp_sf404042 As String
	Private mp_sf404043 As String
	Private mp_sf404044 As String
	Private mp_sf404045 As String
	Private mp_sf404046 As String
	Private mp_sf404047 As String
	Private mp_sf404048 As String
	Private mp_sf404049 As String
	Private mp_sf40404a As String
	Private mp_sf40404b As String
	Private mp_sf40404c As String
	Private mp_sf40404d As String
	Private mp_sf40404e As String
	Private mp_sf40404f As String
	Private mp_sf404050 As String
	Private mp_sf404051 As String
	Private mp_sf404052 As String
	Private mp_sf404053 As String
	Private mp_sf404054 As String
	Private mp_sf404055 As String
	Private mp_sf404056 As String
	Private mp_sf404057 As String
	Private mp_sf404058 As String
	Private mp_sf404059 As String
	Private mp_sf40405a As String
	Private mp_sf40405b As String
	Private mp_sf40405c As String
	Private mp_sf40405d As String
	Private mp_sf40405e As String
	Private mp_sf40405f As String
	Private mp_sf404060 As String
	Private mp_sf404061 As String
	Private mp_sf404062 As String
	Private mp_sf404063 As String
	Private mp_sf404064 As String
	Private mp_sf404065 As String
	Private mp_sf404066 As String
	Private mp_sf404067 As String
	Private mp_sf404068 As String
	Private mp_sf404069 As String
	Private mp_sf40406a As String
	Private mp_sf40406b As String
	Private mp_sf40406c As String
	Private mp_sf40406d As String
	Private mp_sf40406e As String
	Private mp_sf40406f As String
	Private mp_sf404070 As String
	Private mp_sf404071 As String
	Private mp_sf404072 As String
	Private mp_sf404073 As String
	Private mp_sf404074 As String
	Private mp_sf404075 As String
	Private mp_sf404076 As String
	Private mp_sf404077 As String
	Private mp_sf404078 As String
	Private mp_sf404079 As String
	Private mp_sf40407a As String
	Private mp_sf40407b As String
	Private mp_sf40407c As String
	Private mp_sf40407d As String
	Private mp_sf40407e As String
	Private mp_sf40407f As String
	Private mp_sf404080 As String
	Private mp_sf404081 As String
	Private mp_sf404082 As String
	Private mp_sf404083 As String
	Private mp_sf404084 As String
	Private mp_sf404085 As String
	Private mp_sf404086 As String
	Private mp_sf404087 As String
	Private mp_sf404088 As String
	Private mp_sf404089 As String
	Private mp_sf40408a As String
	Private mp_sf40408b As String
	Private mp_sf40408c As String
	Private mp_sf40408d As String
	Private mp_sf40408e As String
	Private mp_sf40408f As String
	Private mp_sf404090 As String
	Private mp_sf404091 As String
	Private mp_sf404092 As String
	Private mp_sf404093 As String
	Private mp_sf404094 As String
	Private mp_sf404095 As String
	Private mp_sf404096 As String
	Private mp_sf404097 As String
	Private mp_sf404098 As String
	Private mp_sf404099 As String
	Private mp_sf40409a As String
	Private mp_sf40409b As String
	Private mp_sf40409c As String
	Private mp_sf40409d As String
	Private mp_sf40409e As String
	Private mp_sf40409f As String
	Private mp_sf4040a0 As String
	Private mp_sf4040a1 As String
	Private mp_sf4040a2 As String
	Private mp_sf4040a3 As String
	Private mp_sf4040a4 As String
	Private mp_sf4040a5 As String
	Private mp_sf4040a6 As String
	Private mp_sf4040a7 As String
	Private mp_sf4040a8 As String
	Private mp_sf4040a9 As String
	Private mp_sf4040aa As String
	Private mp_sf4040ab As String
	Private mp_sf4040ac As String
	Private mp_sf4040ad As String
	Private mp_sf4040ae As String
	Private mp_sf4040af As String
	Private mp_sf4040b0 As String
	Private mp_sf4040b1 As String
	Private mp_sf4040b2 As String
	Private mp_sf4040b3 As String
	Private mp_sf4040b4 As String
	Private mp_sf4040b5 As String
	Private mp_sf4040b6 As String
	Private mp_sf4040b7 As String
	Private mp_sf4040b8 As String
	Private mp_sf4040b9 As String
	Private mp_sf4040ba As String
	Private mp_sf4040bb As String
	Private mp_sf4040bc As String
	Private mp_sf4040bd As String
	Private mp_sf4040be As String
	Private mp_sf4040bf As String
	Private mp_sf4040c0 As String
	Private mp_sf4040c1 As String
	Private mp_sf4040c2 As String
	Private mp_sf4040c3 As String
	Private mp_sf4040c4 As String
	Private mp_sf4040c5 As String
	Private mp_sf4040c6 As String
	Private mp_sf4040c7 As String
	Private mp_sf4040c8 As String
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
		mp_fPeakUnits = 0
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
		mp_bSummary = False
		mp_fSV = 0
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
		mp_sAssnOwner = ""
		mp_sAssnOwnerGuid = ""
		mp_cBudgetCost = 0
		mp_oBudgetWork = New Duration()
		mp_oExtendedAttribute_C = New AssignmentExtendedAttribute_C()
		mp_oBaseline_C = New AssignmentBaseline_C()
		mp_sf404000 = ""
		mp_sf404001 = ""
		mp_sf404002 = ""
		mp_sf404003 = ""
		mp_sf404004 = ""
		mp_sf404005 = ""
		mp_sf404006 = ""
		mp_sf404007 = ""
		mp_sf404008 = ""
		mp_sf404009 = ""
		mp_sf40400a = ""
		mp_sf40400b = ""
		mp_sf40400c = ""
		mp_sf40400d = ""
		mp_sf40400e = ""
		mp_sf40400f = ""
		mp_sf404010 = ""
		mp_sf404011 = ""
		mp_sf404012 = ""
		mp_sf404013 = ""
		mp_sf404014 = ""
		mp_sf404015 = ""
		mp_sf404016 = ""
		mp_sf404017 = ""
		mp_sf404018 = ""
		mp_sf404019 = ""
		mp_sf40401a = ""
		mp_sf40401b = ""
		mp_sf40401c = ""
		mp_sf40401d = ""
		mp_sf40401e = ""
		mp_sf40401f = ""
		mp_sf404020 = ""
		mp_sf404021 = ""
		mp_sf404022 = ""
		mp_sf404023 = ""
		mp_sf404024 = ""
		mp_sf404025 = ""
		mp_sf404026 = ""
		mp_sf404027 = ""
		mp_sf404028 = ""
		mp_sf404029 = ""
		mp_sf40402a = ""
		mp_sf40402b = ""
		mp_sf40402c = ""
		mp_sf40402d = ""
		mp_sf40402e = ""
		mp_sf40402f = ""
		mp_sf404030 = ""
		mp_sf404031 = ""
		mp_sf404032 = ""
		mp_sf404033 = ""
		mp_sf404034 = ""
		mp_sf404035 = ""
		mp_sf404036 = ""
		mp_sf404037 = ""
		mp_sf404038 = ""
		mp_sf404039 = ""
		mp_sf40403a = ""
		mp_sf40403b = ""
		mp_sf40403c = ""
		mp_sf40403d = ""
		mp_sf40403e = ""
		mp_sf40403f = ""
		mp_sf404040 = ""
		mp_sf404041 = ""
		mp_sf404042 = ""
		mp_sf404043 = ""
		mp_sf404044 = ""
		mp_sf404045 = ""
		mp_sf404046 = ""
		mp_sf404047 = ""
		mp_sf404048 = ""
		mp_sf404049 = ""
		mp_sf40404a = ""
		mp_sf40404b = ""
		mp_sf40404c = ""
		mp_sf40404d = ""
		mp_sf40404e = ""
		mp_sf40404f = ""
		mp_sf404050 = ""
		mp_sf404051 = ""
		mp_sf404052 = ""
		mp_sf404053 = ""
		mp_sf404054 = ""
		mp_sf404055 = ""
		mp_sf404056 = ""
		mp_sf404057 = ""
		mp_sf404058 = ""
		mp_sf404059 = ""
		mp_sf40405a = ""
		mp_sf40405b = ""
		mp_sf40405c = ""
		mp_sf40405d = ""
		mp_sf40405e = ""
		mp_sf40405f = ""
		mp_sf404060 = ""
		mp_sf404061 = ""
		mp_sf404062 = ""
		mp_sf404063 = ""
		mp_sf404064 = ""
		mp_sf404065 = ""
		mp_sf404066 = ""
		mp_sf404067 = ""
		mp_sf404068 = ""
		mp_sf404069 = ""
		mp_sf40406a = ""
		mp_sf40406b = ""
		mp_sf40406c = ""
		mp_sf40406d = ""
		mp_sf40406e = ""
		mp_sf40406f = ""
		mp_sf404070 = ""
		mp_sf404071 = ""
		mp_sf404072 = ""
		mp_sf404073 = ""
		mp_sf404074 = ""
		mp_sf404075 = ""
		mp_sf404076 = ""
		mp_sf404077 = ""
		mp_sf404078 = ""
		mp_sf404079 = ""
		mp_sf40407a = ""
		mp_sf40407b = ""
		mp_sf40407c = ""
		mp_sf40407d = ""
		mp_sf40407e = ""
		mp_sf40407f = ""
		mp_sf404080 = ""
		mp_sf404081 = ""
		mp_sf404082 = ""
		mp_sf404083 = ""
		mp_sf404084 = ""
		mp_sf404085 = ""
		mp_sf404086 = ""
		mp_sf404087 = ""
		mp_sf404088 = ""
		mp_sf404089 = ""
		mp_sf40408a = ""
		mp_sf40408b = ""
		mp_sf40408c = ""
		mp_sf40408d = ""
		mp_sf40408e = ""
		mp_sf40408f = ""
		mp_sf404090 = ""
		mp_sf404091 = ""
		mp_sf404092 = ""
		mp_sf404093 = ""
		mp_sf404094 = ""
		mp_sf404095 = ""
		mp_sf404096 = ""
		mp_sf404097 = ""
		mp_sf404098 = ""
		mp_sf404099 = ""
		mp_sf40409a = ""
		mp_sf40409b = ""
		mp_sf40409c = ""
		mp_sf40409d = ""
		mp_sf40409e = ""
		mp_sf40409f = ""
		mp_sf4040a0 = ""
		mp_sf4040a1 = ""
		mp_sf4040a2 = ""
		mp_sf4040a3 = ""
		mp_sf4040a4 = ""
		mp_sf4040a5 = ""
		mp_sf4040a6 = ""
		mp_sf4040a7 = ""
		mp_sf4040a8 = ""
		mp_sf4040a9 = ""
		mp_sf4040aa = ""
		mp_sf4040ab = ""
		mp_sf4040ac = ""
		mp_sf4040ad = ""
		mp_sf4040ae = ""
		mp_sf4040af = ""
		mp_sf4040b0 = ""
		mp_sf4040b1 = ""
		mp_sf4040b2 = ""
		mp_sf4040b3 = ""
		mp_sf4040b4 = ""
		mp_sf4040b5 = ""
		mp_sf4040b6 = ""
		mp_sf4040b7 = ""
		mp_sf4040b8 = ""
		mp_sf4040b9 = ""
		mp_sf4040ba = ""
		mp_sf4040bb = ""
		mp_sf4040bc = ""
		mp_sf4040bd = ""
		mp_sf4040be = ""
		mp_sf4040bf = ""
		mp_sf4040c0 = ""
		mp_sf4040c1 = ""
		mp_sf4040c2 = ""
		mp_sf4040c3 = ""
		mp_sf4040c4 = ""
		mp_sf4040c5 = ""
		mp_sf4040c6 = ""
		mp_sf4040c7 = ""
		mp_sf4040c8 = ""
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

	Public Property fPeakUnits() As Single
		Get
			Return mp_fPeakUnits
		End Get
		Set(ByVal Value As Single)
			mp_fPeakUnits = Value
		End Set
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

	Public Property bSummary() As Boolean
		Get
			Return mp_bSummary
		End Get
		Set(ByVal Value As Boolean)
			mp_bSummary = Value
		End Set
	End Property

	Public Property fSV() As Single
		Get
			Return mp_fSV
		End Get
		Set(ByVal Value As Single)
			mp_fSV = Value
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

	Public Property sAssnOwner() As String
		Get
			Return mp_sAssnOwner
		End Get
		Set(ByVal Value As String)
			mp_sAssnOwner = Value
		End Set
	End Property

	Public Property sAssnOwnerGuid() As String
		Get
			Return mp_sAssnOwnerGuid
		End Get
		Set(ByVal Value As String)
			mp_sAssnOwnerGuid = Value
		End Set
	End Property

	Public Property cBudgetCost() As Decimal
		Get
			Return mp_cBudgetCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cBudgetCost = Value
		End Set
	End Property

	Public ReadOnly Property oBudgetWork() As Duration
		Get
			Return mp_oBudgetWork
		End Get
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

	Public Property sf404000() As String
		Get
			Return mp_sf404000
		End Get
		Set(ByVal Value As String)
			mp_sf404000 = Value
		End Set
	End Property

	Public Property sf404001() As String
		Get
			Return mp_sf404001
		End Get
		Set(ByVal Value As String)
			mp_sf404001 = Value
		End Set
	End Property

	Public Property sf404002() As String
		Get
			Return mp_sf404002
		End Get
		Set(ByVal Value As String)
			mp_sf404002 = Value
		End Set
	End Property

	Public Property sf404003() As String
		Get
			Return mp_sf404003
		End Get
		Set(ByVal Value As String)
			mp_sf404003 = Value
		End Set
	End Property

	Public Property sf404004() As String
		Get
			Return mp_sf404004
		End Get
		Set(ByVal Value As String)
			mp_sf404004 = Value
		End Set
	End Property

	Public Property sf404005() As String
		Get
			Return mp_sf404005
		End Get
		Set(ByVal Value As String)
			mp_sf404005 = Value
		End Set
	End Property

	Public Property sf404006() As String
		Get
			Return mp_sf404006
		End Get
		Set(ByVal Value As String)
			mp_sf404006 = Value
		End Set
	End Property

	Public Property sf404007() As String
		Get
			Return mp_sf404007
		End Get
		Set(ByVal Value As String)
			mp_sf404007 = Value
		End Set
	End Property

	Public Property sf404008() As String
		Get
			Return mp_sf404008
		End Get
		Set(ByVal Value As String)
			mp_sf404008 = Value
		End Set
	End Property

	Public Property sf404009() As String
		Get
			Return mp_sf404009
		End Get
		Set(ByVal Value As String)
			mp_sf404009 = Value
		End Set
	End Property

	Public Property sf40400a() As String
		Get
			Return mp_sf40400a
		End Get
		Set(ByVal Value As String)
			mp_sf40400a = Value
		End Set
	End Property

	Public Property sf40400b() As String
		Get
			Return mp_sf40400b
		End Get
		Set(ByVal Value As String)
			mp_sf40400b = Value
		End Set
	End Property

	Public Property sf40400c() As String
		Get
			Return mp_sf40400c
		End Get
		Set(ByVal Value As String)
			mp_sf40400c = Value
		End Set
	End Property

	Public Property sf40400d() As String
		Get
			Return mp_sf40400d
		End Get
		Set(ByVal Value As String)
			mp_sf40400d = Value
		End Set
	End Property

	Public Property sf40400e() As String
		Get
			Return mp_sf40400e
		End Get
		Set(ByVal Value As String)
			mp_sf40400e = Value
		End Set
	End Property

	Public Property sf40400f() As String
		Get
			Return mp_sf40400f
		End Get
		Set(ByVal Value As String)
			mp_sf40400f = Value
		End Set
	End Property

	Public Property sf404010() As String
		Get
			Return mp_sf404010
		End Get
		Set(ByVal Value As String)
			mp_sf404010 = Value
		End Set
	End Property

	Public Property sf404011() As String
		Get
			Return mp_sf404011
		End Get
		Set(ByVal Value As String)
			mp_sf404011 = Value
		End Set
	End Property

	Public Property sf404012() As String
		Get
			Return mp_sf404012
		End Get
		Set(ByVal Value As String)
			mp_sf404012 = Value
		End Set
	End Property

	Public Property sf404013() As String
		Get
			Return mp_sf404013
		End Get
		Set(ByVal Value As String)
			mp_sf404013 = Value
		End Set
	End Property

	Public Property sf404014() As String
		Get
			Return mp_sf404014
		End Get
		Set(ByVal Value As String)
			mp_sf404014 = Value
		End Set
	End Property

	Public Property sf404015() As String
		Get
			Return mp_sf404015
		End Get
		Set(ByVal Value As String)
			mp_sf404015 = Value
		End Set
	End Property

	Public Property sf404016() As String
		Get
			Return mp_sf404016
		End Get
		Set(ByVal Value As String)
			mp_sf404016 = Value
		End Set
	End Property

	Public Property sf404017() As String
		Get
			Return mp_sf404017
		End Get
		Set(ByVal Value As String)
			mp_sf404017 = Value
		End Set
	End Property

	Public Property sf404018() As String
		Get
			Return mp_sf404018
		End Get
		Set(ByVal Value As String)
			mp_sf404018 = Value
		End Set
	End Property

	Public Property sf404019() As String
		Get
			Return mp_sf404019
		End Get
		Set(ByVal Value As String)
			mp_sf404019 = Value
		End Set
	End Property

	Public Property sf40401a() As String
		Get
			Return mp_sf40401a
		End Get
		Set(ByVal Value As String)
			mp_sf40401a = Value
		End Set
	End Property

	Public Property sf40401b() As String
		Get
			Return mp_sf40401b
		End Get
		Set(ByVal Value As String)
			mp_sf40401b = Value
		End Set
	End Property

	Public Property sf40401c() As String
		Get
			Return mp_sf40401c
		End Get
		Set(ByVal Value As String)
			mp_sf40401c = Value
		End Set
	End Property

	Public Property sf40401d() As String
		Get
			Return mp_sf40401d
		End Get
		Set(ByVal Value As String)
			mp_sf40401d = Value
		End Set
	End Property

	Public Property sf40401e() As String
		Get
			Return mp_sf40401e
		End Get
		Set(ByVal Value As String)
			mp_sf40401e = Value
		End Set
	End Property

	Public Property sf40401f() As String
		Get
			Return mp_sf40401f
		End Get
		Set(ByVal Value As String)
			mp_sf40401f = Value
		End Set
	End Property

	Public Property sf404020() As String
		Get
			Return mp_sf404020
		End Get
		Set(ByVal Value As String)
			mp_sf404020 = Value
		End Set
	End Property

	Public Property sf404021() As String
		Get
			Return mp_sf404021
		End Get
		Set(ByVal Value As String)
			mp_sf404021 = Value
		End Set
	End Property

	Public Property sf404022() As String
		Get
			Return mp_sf404022
		End Get
		Set(ByVal Value As String)
			mp_sf404022 = Value
		End Set
	End Property

	Public Property sf404023() As String
		Get
			Return mp_sf404023
		End Get
		Set(ByVal Value As String)
			mp_sf404023 = Value
		End Set
	End Property

	Public Property sf404024() As String
		Get
			Return mp_sf404024
		End Get
		Set(ByVal Value As String)
			mp_sf404024 = Value
		End Set
	End Property

	Public Property sf404025() As String
		Get
			Return mp_sf404025
		End Get
		Set(ByVal Value As String)
			mp_sf404025 = Value
		End Set
	End Property

	Public Property sf404026() As String
		Get
			Return mp_sf404026
		End Get
		Set(ByVal Value As String)
			mp_sf404026 = Value
		End Set
	End Property

	Public Property sf404027() As String
		Get
			Return mp_sf404027
		End Get
		Set(ByVal Value As String)
			mp_sf404027 = Value
		End Set
	End Property

	Public Property sf404028() As String
		Get
			Return mp_sf404028
		End Get
		Set(ByVal Value As String)
			mp_sf404028 = Value
		End Set
	End Property

	Public Property sf404029() As String
		Get
			Return mp_sf404029
		End Get
		Set(ByVal Value As String)
			mp_sf404029 = Value
		End Set
	End Property

	Public Property sf40402a() As String
		Get
			Return mp_sf40402a
		End Get
		Set(ByVal Value As String)
			mp_sf40402a = Value
		End Set
	End Property

	Public Property sf40402b() As String
		Get
			Return mp_sf40402b
		End Get
		Set(ByVal Value As String)
			mp_sf40402b = Value
		End Set
	End Property

	Public Property sf40402c() As String
		Get
			Return mp_sf40402c
		End Get
		Set(ByVal Value As String)
			mp_sf40402c = Value
		End Set
	End Property

	Public Property sf40402d() As String
		Get
			Return mp_sf40402d
		End Get
		Set(ByVal Value As String)
			mp_sf40402d = Value
		End Set
	End Property

	Public Property sf40402e() As String
		Get
			Return mp_sf40402e
		End Get
		Set(ByVal Value As String)
			mp_sf40402e = Value
		End Set
	End Property

	Public Property sf40402f() As String
		Get
			Return mp_sf40402f
		End Get
		Set(ByVal Value As String)
			mp_sf40402f = Value
		End Set
	End Property

	Public Property sf404030() As String
		Get
			Return mp_sf404030
		End Get
		Set(ByVal Value As String)
			mp_sf404030 = Value
		End Set
	End Property

	Public Property sf404031() As String
		Get
			Return mp_sf404031
		End Get
		Set(ByVal Value As String)
			mp_sf404031 = Value
		End Set
	End Property

	Public Property sf404032() As String
		Get
			Return mp_sf404032
		End Get
		Set(ByVal Value As String)
			mp_sf404032 = Value
		End Set
	End Property

	Public Property sf404033() As String
		Get
			Return mp_sf404033
		End Get
		Set(ByVal Value As String)
			mp_sf404033 = Value
		End Set
	End Property

	Public Property sf404034() As String
		Get
			Return mp_sf404034
		End Get
		Set(ByVal Value As String)
			mp_sf404034 = Value
		End Set
	End Property

	Public Property sf404035() As String
		Get
			Return mp_sf404035
		End Get
		Set(ByVal Value As String)
			mp_sf404035 = Value
		End Set
	End Property

	Public Property sf404036() As String
		Get
			Return mp_sf404036
		End Get
		Set(ByVal Value As String)
			mp_sf404036 = Value
		End Set
	End Property

	Public Property sf404037() As String
		Get
			Return mp_sf404037
		End Get
		Set(ByVal Value As String)
			mp_sf404037 = Value
		End Set
	End Property

	Public Property sf404038() As String
		Get
			Return mp_sf404038
		End Get
		Set(ByVal Value As String)
			mp_sf404038 = Value
		End Set
	End Property

	Public Property sf404039() As String
		Get
			Return mp_sf404039
		End Get
		Set(ByVal Value As String)
			mp_sf404039 = Value
		End Set
	End Property

	Public Property sf40403a() As String
		Get
			Return mp_sf40403a
		End Get
		Set(ByVal Value As String)
			mp_sf40403a = Value
		End Set
	End Property

	Public Property sf40403b() As String
		Get
			Return mp_sf40403b
		End Get
		Set(ByVal Value As String)
			mp_sf40403b = Value
		End Set
	End Property

	Public Property sf40403c() As String
		Get
			Return mp_sf40403c
		End Get
		Set(ByVal Value As String)
			mp_sf40403c = Value
		End Set
	End Property

	Public Property sf40403d() As String
		Get
			Return mp_sf40403d
		End Get
		Set(ByVal Value As String)
			mp_sf40403d = Value
		End Set
	End Property

	Public Property sf40403e() As String
		Get
			Return mp_sf40403e
		End Get
		Set(ByVal Value As String)
			mp_sf40403e = Value
		End Set
	End Property

	Public Property sf40403f() As String
		Get
			Return mp_sf40403f
		End Get
		Set(ByVal Value As String)
			mp_sf40403f = Value
		End Set
	End Property

	Public Property sf404040() As String
		Get
			Return mp_sf404040
		End Get
		Set(ByVal Value As String)
			mp_sf404040 = Value
		End Set
	End Property

	Public Property sf404041() As String
		Get
			Return mp_sf404041
		End Get
		Set(ByVal Value As String)
			mp_sf404041 = Value
		End Set
	End Property

	Public Property sf404042() As String
		Get
			Return mp_sf404042
		End Get
		Set(ByVal Value As String)
			mp_sf404042 = Value
		End Set
	End Property

	Public Property sf404043() As String
		Get
			Return mp_sf404043
		End Get
		Set(ByVal Value As String)
			mp_sf404043 = Value
		End Set
	End Property

	Public Property sf404044() As String
		Get
			Return mp_sf404044
		End Get
		Set(ByVal Value As String)
			mp_sf404044 = Value
		End Set
	End Property

	Public Property sf404045() As String
		Get
			Return mp_sf404045
		End Get
		Set(ByVal Value As String)
			mp_sf404045 = Value
		End Set
	End Property

	Public Property sf404046() As String
		Get
			Return mp_sf404046
		End Get
		Set(ByVal Value As String)
			mp_sf404046 = Value
		End Set
	End Property

	Public Property sf404047() As String
		Get
			Return mp_sf404047
		End Get
		Set(ByVal Value As String)
			mp_sf404047 = Value
		End Set
	End Property

	Public Property sf404048() As String
		Get
			Return mp_sf404048
		End Get
		Set(ByVal Value As String)
			mp_sf404048 = Value
		End Set
	End Property

	Public Property sf404049() As String
		Get
			Return mp_sf404049
		End Get
		Set(ByVal Value As String)
			mp_sf404049 = Value
		End Set
	End Property

	Public Property sf40404a() As String
		Get
			Return mp_sf40404a
		End Get
		Set(ByVal Value As String)
			mp_sf40404a = Value
		End Set
	End Property

	Public Property sf40404b() As String
		Get
			Return mp_sf40404b
		End Get
		Set(ByVal Value As String)
			mp_sf40404b = Value
		End Set
	End Property

	Public Property sf40404c() As String
		Get
			Return mp_sf40404c
		End Get
		Set(ByVal Value As String)
			mp_sf40404c = Value
		End Set
	End Property

	Public Property sf40404d() As String
		Get
			Return mp_sf40404d
		End Get
		Set(ByVal Value As String)
			mp_sf40404d = Value
		End Set
	End Property

	Public Property sf40404e() As String
		Get
			Return mp_sf40404e
		End Get
		Set(ByVal Value As String)
			mp_sf40404e = Value
		End Set
	End Property

	Public Property sf40404f() As String
		Get
			Return mp_sf40404f
		End Get
		Set(ByVal Value As String)
			mp_sf40404f = Value
		End Set
	End Property

	Public Property sf404050() As String
		Get
			Return mp_sf404050
		End Get
		Set(ByVal Value As String)
			mp_sf404050 = Value
		End Set
	End Property

	Public Property sf404051() As String
		Get
			Return mp_sf404051
		End Get
		Set(ByVal Value As String)
			mp_sf404051 = Value
		End Set
	End Property

	Public Property sf404052() As String
		Get
			Return mp_sf404052
		End Get
		Set(ByVal Value As String)
			mp_sf404052 = Value
		End Set
	End Property

	Public Property sf404053() As String
		Get
			Return mp_sf404053
		End Get
		Set(ByVal Value As String)
			mp_sf404053 = Value
		End Set
	End Property

	Public Property sf404054() As String
		Get
			Return mp_sf404054
		End Get
		Set(ByVal Value As String)
			mp_sf404054 = Value
		End Set
	End Property

	Public Property sf404055() As String
		Get
			Return mp_sf404055
		End Get
		Set(ByVal Value As String)
			mp_sf404055 = Value
		End Set
	End Property

	Public Property sf404056() As String
		Get
			Return mp_sf404056
		End Get
		Set(ByVal Value As String)
			mp_sf404056 = Value
		End Set
	End Property

	Public Property sf404057() As String
		Get
			Return mp_sf404057
		End Get
		Set(ByVal Value As String)
			mp_sf404057 = Value
		End Set
	End Property

	Public Property sf404058() As String
		Get
			Return mp_sf404058
		End Get
		Set(ByVal Value As String)
			mp_sf404058 = Value
		End Set
	End Property

	Public Property sf404059() As String
		Get
			Return mp_sf404059
		End Get
		Set(ByVal Value As String)
			mp_sf404059 = Value
		End Set
	End Property

	Public Property sf40405a() As String
		Get
			Return mp_sf40405a
		End Get
		Set(ByVal Value As String)
			mp_sf40405a = Value
		End Set
	End Property

	Public Property sf40405b() As String
		Get
			Return mp_sf40405b
		End Get
		Set(ByVal Value As String)
			mp_sf40405b = Value
		End Set
	End Property

	Public Property sf40405c() As String
		Get
			Return mp_sf40405c
		End Get
		Set(ByVal Value As String)
			mp_sf40405c = Value
		End Set
	End Property

	Public Property sf40405d() As String
		Get
			Return mp_sf40405d
		End Get
		Set(ByVal Value As String)
			mp_sf40405d = Value
		End Set
	End Property

	Public Property sf40405e() As String
		Get
			Return mp_sf40405e
		End Get
		Set(ByVal Value As String)
			mp_sf40405e = Value
		End Set
	End Property

	Public Property sf40405f() As String
		Get
			Return mp_sf40405f
		End Get
		Set(ByVal Value As String)
			mp_sf40405f = Value
		End Set
	End Property

	Public Property sf404060() As String
		Get
			Return mp_sf404060
		End Get
		Set(ByVal Value As String)
			mp_sf404060 = Value
		End Set
	End Property

	Public Property sf404061() As String
		Get
			Return mp_sf404061
		End Get
		Set(ByVal Value As String)
			mp_sf404061 = Value
		End Set
	End Property

	Public Property sf404062() As String
		Get
			Return mp_sf404062
		End Get
		Set(ByVal Value As String)
			mp_sf404062 = Value
		End Set
	End Property

	Public Property sf404063() As String
		Get
			Return mp_sf404063
		End Get
		Set(ByVal Value As String)
			mp_sf404063 = Value
		End Set
	End Property

	Public Property sf404064() As String
		Get
			Return mp_sf404064
		End Get
		Set(ByVal Value As String)
			mp_sf404064 = Value
		End Set
	End Property

	Public Property sf404065() As String
		Get
			Return mp_sf404065
		End Get
		Set(ByVal Value As String)
			mp_sf404065 = Value
		End Set
	End Property

	Public Property sf404066() As String
		Get
			Return mp_sf404066
		End Get
		Set(ByVal Value As String)
			mp_sf404066 = Value
		End Set
	End Property

	Public Property sf404067() As String
		Get
			Return mp_sf404067
		End Get
		Set(ByVal Value As String)
			mp_sf404067 = Value
		End Set
	End Property

	Public Property sf404068() As String
		Get
			Return mp_sf404068
		End Get
		Set(ByVal Value As String)
			mp_sf404068 = Value
		End Set
	End Property

	Public Property sf404069() As String
		Get
			Return mp_sf404069
		End Get
		Set(ByVal Value As String)
			mp_sf404069 = Value
		End Set
	End Property

	Public Property sf40406a() As String
		Get
			Return mp_sf40406a
		End Get
		Set(ByVal Value As String)
			mp_sf40406a = Value
		End Set
	End Property

	Public Property sf40406b() As String
		Get
			Return mp_sf40406b
		End Get
		Set(ByVal Value As String)
			mp_sf40406b = Value
		End Set
	End Property

	Public Property sf40406c() As String
		Get
			Return mp_sf40406c
		End Get
		Set(ByVal Value As String)
			mp_sf40406c = Value
		End Set
	End Property

	Public Property sf40406d() As String
		Get
			Return mp_sf40406d
		End Get
		Set(ByVal Value As String)
			mp_sf40406d = Value
		End Set
	End Property

	Public Property sf40406e() As String
		Get
			Return mp_sf40406e
		End Get
		Set(ByVal Value As String)
			mp_sf40406e = Value
		End Set
	End Property

	Public Property sf40406f() As String
		Get
			Return mp_sf40406f
		End Get
		Set(ByVal Value As String)
			mp_sf40406f = Value
		End Set
	End Property

	Public Property sf404070() As String
		Get
			Return mp_sf404070
		End Get
		Set(ByVal Value As String)
			mp_sf404070 = Value
		End Set
	End Property

	Public Property sf404071() As String
		Get
			Return mp_sf404071
		End Get
		Set(ByVal Value As String)
			mp_sf404071 = Value
		End Set
	End Property

	Public Property sf404072() As String
		Get
			Return mp_sf404072
		End Get
		Set(ByVal Value As String)
			mp_sf404072 = Value
		End Set
	End Property

	Public Property sf404073() As String
		Get
			Return mp_sf404073
		End Get
		Set(ByVal Value As String)
			mp_sf404073 = Value
		End Set
	End Property

	Public Property sf404074() As String
		Get
			Return mp_sf404074
		End Get
		Set(ByVal Value As String)
			mp_sf404074 = Value
		End Set
	End Property

	Public Property sf404075() As String
		Get
			Return mp_sf404075
		End Get
		Set(ByVal Value As String)
			mp_sf404075 = Value
		End Set
	End Property

	Public Property sf404076() As String
		Get
			Return mp_sf404076
		End Get
		Set(ByVal Value As String)
			mp_sf404076 = Value
		End Set
	End Property

	Public Property sf404077() As String
		Get
			Return mp_sf404077
		End Get
		Set(ByVal Value As String)
			mp_sf404077 = Value
		End Set
	End Property

	Public Property sf404078() As String
		Get
			Return mp_sf404078
		End Get
		Set(ByVal Value As String)
			mp_sf404078 = Value
		End Set
	End Property

	Public Property sf404079() As String
		Get
			Return mp_sf404079
		End Get
		Set(ByVal Value As String)
			mp_sf404079 = Value
		End Set
	End Property

	Public Property sf40407a() As String
		Get
			Return mp_sf40407a
		End Get
		Set(ByVal Value As String)
			mp_sf40407a = Value
		End Set
	End Property

	Public Property sf40407b() As String
		Get
			Return mp_sf40407b
		End Get
		Set(ByVal Value As String)
			mp_sf40407b = Value
		End Set
	End Property

	Public Property sf40407c() As String
		Get
			Return mp_sf40407c
		End Get
		Set(ByVal Value As String)
			mp_sf40407c = Value
		End Set
	End Property

	Public Property sf40407d() As String
		Get
			Return mp_sf40407d
		End Get
		Set(ByVal Value As String)
			mp_sf40407d = Value
		End Set
	End Property

	Public Property sf40407e() As String
		Get
			Return mp_sf40407e
		End Get
		Set(ByVal Value As String)
			mp_sf40407e = Value
		End Set
	End Property

	Public Property sf40407f() As String
		Get
			Return mp_sf40407f
		End Get
		Set(ByVal Value As String)
			mp_sf40407f = Value
		End Set
	End Property

	Public Property sf404080() As String
		Get
			Return mp_sf404080
		End Get
		Set(ByVal Value As String)
			mp_sf404080 = Value
		End Set
	End Property

	Public Property sf404081() As String
		Get
			Return mp_sf404081
		End Get
		Set(ByVal Value As String)
			mp_sf404081 = Value
		End Set
	End Property

	Public Property sf404082() As String
		Get
			Return mp_sf404082
		End Get
		Set(ByVal Value As String)
			mp_sf404082 = Value
		End Set
	End Property

	Public Property sf404083() As String
		Get
			Return mp_sf404083
		End Get
		Set(ByVal Value As String)
			mp_sf404083 = Value
		End Set
	End Property

	Public Property sf404084() As String
		Get
			Return mp_sf404084
		End Get
		Set(ByVal Value As String)
			mp_sf404084 = Value
		End Set
	End Property

	Public Property sf404085() As String
		Get
			Return mp_sf404085
		End Get
		Set(ByVal Value As String)
			mp_sf404085 = Value
		End Set
	End Property

	Public Property sf404086() As String
		Get
			Return mp_sf404086
		End Get
		Set(ByVal Value As String)
			mp_sf404086 = Value
		End Set
	End Property

	Public Property sf404087() As String
		Get
			Return mp_sf404087
		End Get
		Set(ByVal Value As String)
			mp_sf404087 = Value
		End Set
	End Property

	Public Property sf404088() As String
		Get
			Return mp_sf404088
		End Get
		Set(ByVal Value As String)
			mp_sf404088 = Value
		End Set
	End Property

	Public Property sf404089() As String
		Get
			Return mp_sf404089
		End Get
		Set(ByVal Value As String)
			mp_sf404089 = Value
		End Set
	End Property

	Public Property sf40408a() As String
		Get
			Return mp_sf40408a
		End Get
		Set(ByVal Value As String)
			mp_sf40408a = Value
		End Set
	End Property

	Public Property sf40408b() As String
		Get
			Return mp_sf40408b
		End Get
		Set(ByVal Value As String)
			mp_sf40408b = Value
		End Set
	End Property

	Public Property sf40408c() As String
		Get
			Return mp_sf40408c
		End Get
		Set(ByVal Value As String)
			mp_sf40408c = Value
		End Set
	End Property

	Public Property sf40408d() As String
		Get
			Return mp_sf40408d
		End Get
		Set(ByVal Value As String)
			mp_sf40408d = Value
		End Set
	End Property

	Public Property sf40408e() As String
		Get
			Return mp_sf40408e
		End Get
		Set(ByVal Value As String)
			mp_sf40408e = Value
		End Set
	End Property

	Public Property sf40408f() As String
		Get
			Return mp_sf40408f
		End Get
		Set(ByVal Value As String)
			mp_sf40408f = Value
		End Set
	End Property

	Public Property sf404090() As String
		Get
			Return mp_sf404090
		End Get
		Set(ByVal Value As String)
			mp_sf404090 = Value
		End Set
	End Property

	Public Property sf404091() As String
		Get
			Return mp_sf404091
		End Get
		Set(ByVal Value As String)
			mp_sf404091 = Value
		End Set
	End Property

	Public Property sf404092() As String
		Get
			Return mp_sf404092
		End Get
		Set(ByVal Value As String)
			mp_sf404092 = Value
		End Set
	End Property

	Public Property sf404093() As String
		Get
			Return mp_sf404093
		End Get
		Set(ByVal Value As String)
			mp_sf404093 = Value
		End Set
	End Property

	Public Property sf404094() As String
		Get
			Return mp_sf404094
		End Get
		Set(ByVal Value As String)
			mp_sf404094 = Value
		End Set
	End Property

	Public Property sf404095() As String
		Get
			Return mp_sf404095
		End Get
		Set(ByVal Value As String)
			mp_sf404095 = Value
		End Set
	End Property

	Public Property sf404096() As String
		Get
			Return mp_sf404096
		End Get
		Set(ByVal Value As String)
			mp_sf404096 = Value
		End Set
	End Property

	Public Property sf404097() As String
		Get
			Return mp_sf404097
		End Get
		Set(ByVal Value As String)
			mp_sf404097 = Value
		End Set
	End Property

	Public Property sf404098() As String
		Get
			Return mp_sf404098
		End Get
		Set(ByVal Value As String)
			mp_sf404098 = Value
		End Set
	End Property

	Public Property sf404099() As String
		Get
			Return mp_sf404099
		End Get
		Set(ByVal Value As String)
			mp_sf404099 = Value
		End Set
	End Property

	Public Property sf40409a() As String
		Get
			Return mp_sf40409a
		End Get
		Set(ByVal Value As String)
			mp_sf40409a = Value
		End Set
	End Property

	Public Property sf40409b() As String
		Get
			Return mp_sf40409b
		End Get
		Set(ByVal Value As String)
			mp_sf40409b = Value
		End Set
	End Property

	Public Property sf40409c() As String
		Get
			Return mp_sf40409c
		End Get
		Set(ByVal Value As String)
			mp_sf40409c = Value
		End Set
	End Property

	Public Property sf40409d() As String
		Get
			Return mp_sf40409d
		End Get
		Set(ByVal Value As String)
			mp_sf40409d = Value
		End Set
	End Property

	Public Property sf40409e() As String
		Get
			Return mp_sf40409e
		End Get
		Set(ByVal Value As String)
			mp_sf40409e = Value
		End Set
	End Property

	Public Property sf40409f() As String
		Get
			Return mp_sf40409f
		End Get
		Set(ByVal Value As String)
			mp_sf40409f = Value
		End Set
	End Property

	Public Property sf4040a0() As String
		Get
			Return mp_sf4040a0
		End Get
		Set(ByVal Value As String)
			mp_sf4040a0 = Value
		End Set
	End Property

	Public Property sf4040a1() As String
		Get
			Return mp_sf4040a1
		End Get
		Set(ByVal Value As String)
			mp_sf4040a1 = Value
		End Set
	End Property

	Public Property sf4040a2() As String
		Get
			Return mp_sf4040a2
		End Get
		Set(ByVal Value As String)
			mp_sf4040a2 = Value
		End Set
	End Property

	Public Property sf4040a3() As String
		Get
			Return mp_sf4040a3
		End Get
		Set(ByVal Value As String)
			mp_sf4040a3 = Value
		End Set
	End Property

	Public Property sf4040a4() As String
		Get
			Return mp_sf4040a4
		End Get
		Set(ByVal Value As String)
			mp_sf4040a4 = Value
		End Set
	End Property

	Public Property sf4040a5() As String
		Get
			Return mp_sf4040a5
		End Get
		Set(ByVal Value As String)
			mp_sf4040a5 = Value
		End Set
	End Property

	Public Property sf4040a6() As String
		Get
			Return mp_sf4040a6
		End Get
		Set(ByVal Value As String)
			mp_sf4040a6 = Value
		End Set
	End Property

	Public Property sf4040a7() As String
		Get
			Return mp_sf4040a7
		End Get
		Set(ByVal Value As String)
			mp_sf4040a7 = Value
		End Set
	End Property

	Public Property sf4040a8() As String
		Get
			Return mp_sf4040a8
		End Get
		Set(ByVal Value As String)
			mp_sf4040a8 = Value
		End Set
	End Property

	Public Property sf4040a9() As String
		Get
			Return mp_sf4040a9
		End Get
		Set(ByVal Value As String)
			mp_sf4040a9 = Value
		End Set
	End Property

	Public Property sf4040aa() As String
		Get
			Return mp_sf4040aa
		End Get
		Set(ByVal Value As String)
			mp_sf4040aa = Value
		End Set
	End Property

	Public Property sf4040ab() As String
		Get
			Return mp_sf4040ab
		End Get
		Set(ByVal Value As String)
			mp_sf4040ab = Value
		End Set
	End Property

	Public Property sf4040ac() As String
		Get
			Return mp_sf4040ac
		End Get
		Set(ByVal Value As String)
			mp_sf4040ac = Value
		End Set
	End Property

	Public Property sf4040ad() As String
		Get
			Return mp_sf4040ad
		End Get
		Set(ByVal Value As String)
			mp_sf4040ad = Value
		End Set
	End Property

	Public Property sf4040ae() As String
		Get
			Return mp_sf4040ae
		End Get
		Set(ByVal Value As String)
			mp_sf4040ae = Value
		End Set
	End Property

	Public Property sf4040af() As String
		Get
			Return mp_sf4040af
		End Get
		Set(ByVal Value As String)
			mp_sf4040af = Value
		End Set
	End Property

	Public Property sf4040b0() As String
		Get
			Return mp_sf4040b0
		End Get
		Set(ByVal Value As String)
			mp_sf4040b0 = Value
		End Set
	End Property

	Public Property sf4040b1() As String
		Get
			Return mp_sf4040b1
		End Get
		Set(ByVal Value As String)
			mp_sf4040b1 = Value
		End Set
	End Property

	Public Property sf4040b2() As String
		Get
			Return mp_sf4040b2
		End Get
		Set(ByVal Value As String)
			mp_sf4040b2 = Value
		End Set
	End Property

	Public Property sf4040b3() As String
		Get
			Return mp_sf4040b3
		End Get
		Set(ByVal Value As String)
			mp_sf4040b3 = Value
		End Set
	End Property

	Public Property sf4040b4() As String
		Get
			Return mp_sf4040b4
		End Get
		Set(ByVal Value As String)
			mp_sf4040b4 = Value
		End Set
	End Property

	Public Property sf4040b5() As String
		Get
			Return mp_sf4040b5
		End Get
		Set(ByVal Value As String)
			mp_sf4040b5 = Value
		End Set
	End Property

	Public Property sf4040b6() As String
		Get
			Return mp_sf4040b6
		End Get
		Set(ByVal Value As String)
			mp_sf4040b6 = Value
		End Set
	End Property

	Public Property sf4040b7() As String
		Get
			Return mp_sf4040b7
		End Get
		Set(ByVal Value As String)
			mp_sf4040b7 = Value
		End Set
	End Property

	Public Property sf4040b8() As String
		Get
			Return mp_sf4040b8
		End Get
		Set(ByVal Value As String)
			mp_sf4040b8 = Value
		End Set
	End Property

	Public Property sf4040b9() As String
		Get
			Return mp_sf4040b9
		End Get
		Set(ByVal Value As String)
			mp_sf4040b9 = Value
		End Set
	End Property

	Public Property sf4040ba() As String
		Get
			Return mp_sf4040ba
		End Get
		Set(ByVal Value As String)
			mp_sf4040ba = Value
		End Set
	End Property

	Public Property sf4040bb() As String
		Get
			Return mp_sf4040bb
		End Get
		Set(ByVal Value As String)
			mp_sf4040bb = Value
		End Set
	End Property

	Public Property sf4040bc() As String
		Get
			Return mp_sf4040bc
		End Get
		Set(ByVal Value As String)
			mp_sf4040bc = Value
		End Set
	End Property

	Public Property sf4040bd() As String
		Get
			Return mp_sf4040bd
		End Get
		Set(ByVal Value As String)
			mp_sf4040bd = Value
		End Set
	End Property

	Public Property sf4040be() As String
		Get
			Return mp_sf4040be
		End Get
		Set(ByVal Value As String)
			mp_sf4040be = Value
		End Set
	End Property

	Public Property sf4040bf() As String
		Get
			Return mp_sf4040bf
		End Get
		Set(ByVal Value As String)
			mp_sf4040bf = Value
		End Set
	End Property

	Public Property sf4040c0() As String
		Get
			Return mp_sf4040c0
		End Get
		Set(ByVal Value As String)
			mp_sf4040c0 = Value
		End Set
	End Property

	Public Property sf4040c1() As String
		Get
			Return mp_sf4040c1
		End Get
		Set(ByVal Value As String)
			mp_sf4040c1 = Value
		End Set
	End Property

	Public Property sf4040c2() As String
		Get
			Return mp_sf4040c2
		End Get
		Set(ByVal Value As String)
			mp_sf4040c2 = Value
		End Set
	End Property

	Public Property sf4040c3() As String
		Get
			Return mp_sf4040c3
		End Get
		Set(ByVal Value As String)
			mp_sf4040c3 = Value
		End Set
	End Property

	Public Property sf4040c4() As String
		Get
			Return mp_sf4040c4
		End Get
		Set(ByVal Value As String)
			mp_sf4040c4 = Value
		End Set
	End Property

	Public Property sf4040c5() As String
		Get
			Return mp_sf4040c5
		End Get
		Set(ByVal Value As String)
			mp_sf4040c5 = Value
		End Set
	End Property

	Public Property sf4040c6() As String
		Get
			Return mp_sf4040c6
		End Get
		Set(ByVal Value As String)
			mp_sf4040c6 = Value
		End Set
	End Property

	Public Property sf4040c7() As String
		Get
			Return mp_sf4040c7
		End Get
		Set(ByVal Value As String)
			mp_sf4040c7 = Value
		End Set
	End Property

	Public Property sf4040c8() As String
		Get
			Return mp_sf4040c8
		End Get
		Set(ByVal Value As String)
			mp_sf4040c8 = Value
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
		If mp_fPeakUnits <> 0 Then
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
		If mp_bSummary <> False Then
			bReturn = False
		End If
		If mp_fSV <> 0 Then
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
		If mp_sAssnOwner <> "" Then
			bReturn = False
		End If
		If mp_sAssnOwnerGuid <> "" Then
			bReturn = False
		End If
		If mp_cBudgetCost <> 0 Then
			bReturn = False
		End If
		If mp_oBudgetWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_oBaseline_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_sf404000 <> "" Then
			bReturn = False
		End If
		If mp_sf404001 <> "" Then
			bReturn = False
		End If
		If mp_sf404002 <> "" Then
			bReturn = False
		End If
		If mp_sf404003 <> "" Then
			bReturn = False
		End If
		If mp_sf404004 <> "" Then
			bReturn = False
		End If
		If mp_sf404005 <> "" Then
			bReturn = False
		End If
		If mp_sf404006 <> "" Then
			bReturn = False
		End If
		If mp_sf404007 <> "" Then
			bReturn = False
		End If
		If mp_sf404008 <> "" Then
			bReturn = False
		End If
		If mp_sf404009 <> "" Then
			bReturn = False
		End If
		If mp_sf40400a <> "" Then
			bReturn = False
		End If
		If mp_sf40400b <> "" Then
			bReturn = False
		End If
		If mp_sf40400c <> "" Then
			bReturn = False
		End If
		If mp_sf40400d <> "" Then
			bReturn = False
		End If
		If mp_sf40400e <> "" Then
			bReturn = False
		End If
		If mp_sf40400f <> "" Then
			bReturn = False
		End If
		If mp_sf404010 <> "" Then
			bReturn = False
		End If
		If mp_sf404011 <> "" Then
			bReturn = False
		End If
		If mp_sf404012 <> "" Then
			bReturn = False
		End If
		If mp_sf404013 <> "" Then
			bReturn = False
		End If
		If mp_sf404014 <> "" Then
			bReturn = False
		End If
		If mp_sf404015 <> "" Then
			bReturn = False
		End If
		If mp_sf404016 <> "" Then
			bReturn = False
		End If
		If mp_sf404017 <> "" Then
			bReturn = False
		End If
		If mp_sf404018 <> "" Then
			bReturn = False
		End If
		If mp_sf404019 <> "" Then
			bReturn = False
		End If
		If mp_sf40401a <> "" Then
			bReturn = False
		End If
		If mp_sf40401b <> "" Then
			bReturn = False
		End If
		If mp_sf40401c <> "" Then
			bReturn = False
		End If
		If mp_sf40401d <> "" Then
			bReturn = False
		End If
		If mp_sf40401e <> "" Then
			bReturn = False
		End If
		If mp_sf40401f <> "" Then
			bReturn = False
		End If
		If mp_sf404020 <> "" Then
			bReturn = False
		End If
		If mp_sf404021 <> "" Then
			bReturn = False
		End If
		If mp_sf404022 <> "" Then
			bReturn = False
		End If
		If mp_sf404023 <> "" Then
			bReturn = False
		End If
		If mp_sf404024 <> "" Then
			bReturn = False
		End If
		If mp_sf404025 <> "" Then
			bReturn = False
		End If
		If mp_sf404026 <> "" Then
			bReturn = False
		End If
		If mp_sf404027 <> "" Then
			bReturn = False
		End If
		If mp_sf404028 <> "" Then
			bReturn = False
		End If
		If mp_sf404029 <> "" Then
			bReturn = False
		End If
		If mp_sf40402a <> "" Then
			bReturn = False
		End If
		If mp_sf40402b <> "" Then
			bReturn = False
		End If
		If mp_sf40402c <> "" Then
			bReturn = False
		End If
		If mp_sf40402d <> "" Then
			bReturn = False
		End If
		If mp_sf40402e <> "" Then
			bReturn = False
		End If
		If mp_sf40402f <> "" Then
			bReturn = False
		End If
		If mp_sf404030 <> "" Then
			bReturn = False
		End If
		If mp_sf404031 <> "" Then
			bReturn = False
		End If
		If mp_sf404032 <> "" Then
			bReturn = False
		End If
		If mp_sf404033 <> "" Then
			bReturn = False
		End If
		If mp_sf404034 <> "" Then
			bReturn = False
		End If
		If mp_sf404035 <> "" Then
			bReturn = False
		End If
		If mp_sf404036 <> "" Then
			bReturn = False
		End If
		If mp_sf404037 <> "" Then
			bReturn = False
		End If
		If mp_sf404038 <> "" Then
			bReturn = False
		End If
		If mp_sf404039 <> "" Then
			bReturn = False
		End If
		If mp_sf40403a <> "" Then
			bReturn = False
		End If
		If mp_sf40403b <> "" Then
			bReturn = False
		End If
		If mp_sf40403c <> "" Then
			bReturn = False
		End If
		If mp_sf40403d <> "" Then
			bReturn = False
		End If
		If mp_sf40403e <> "" Then
			bReturn = False
		End If
		If mp_sf40403f <> "" Then
			bReturn = False
		End If
		If mp_sf404040 <> "" Then
			bReturn = False
		End If
		If mp_sf404041 <> "" Then
			bReturn = False
		End If
		If mp_sf404042 <> "" Then
			bReturn = False
		End If
		If mp_sf404043 <> "" Then
			bReturn = False
		End If
		If mp_sf404044 <> "" Then
			bReturn = False
		End If
		If mp_sf404045 <> "" Then
			bReturn = False
		End If
		If mp_sf404046 <> "" Then
			bReturn = False
		End If
		If mp_sf404047 <> "" Then
			bReturn = False
		End If
		If mp_sf404048 <> "" Then
			bReturn = False
		End If
		If mp_sf404049 <> "" Then
			bReturn = False
		End If
		If mp_sf40404a <> "" Then
			bReturn = False
		End If
		If mp_sf40404b <> "" Then
			bReturn = False
		End If
		If mp_sf40404c <> "" Then
			bReturn = False
		End If
		If mp_sf40404d <> "" Then
			bReturn = False
		End If
		If mp_sf40404e <> "" Then
			bReturn = False
		End If
		If mp_sf40404f <> "" Then
			bReturn = False
		End If
		If mp_sf404050 <> "" Then
			bReturn = False
		End If
		If mp_sf404051 <> "" Then
			bReturn = False
		End If
		If mp_sf404052 <> "" Then
			bReturn = False
		End If
		If mp_sf404053 <> "" Then
			bReturn = False
		End If
		If mp_sf404054 <> "" Then
			bReturn = False
		End If
		If mp_sf404055 <> "" Then
			bReturn = False
		End If
		If mp_sf404056 <> "" Then
			bReturn = False
		End If
		If mp_sf404057 <> "" Then
			bReturn = False
		End If
		If mp_sf404058 <> "" Then
			bReturn = False
		End If
		If mp_sf404059 <> "" Then
			bReturn = False
		End If
		If mp_sf40405a <> "" Then
			bReturn = False
		End If
		If mp_sf40405b <> "" Then
			bReturn = False
		End If
		If mp_sf40405c <> "" Then
			bReturn = False
		End If
		If mp_sf40405d <> "" Then
			bReturn = False
		End If
		If mp_sf40405e <> "" Then
			bReturn = False
		End If
		If mp_sf40405f <> "" Then
			bReturn = False
		End If
		If mp_sf404060 <> "" Then
			bReturn = False
		End If
		If mp_sf404061 <> "" Then
			bReturn = False
		End If
		If mp_sf404062 <> "" Then
			bReturn = False
		End If
		If mp_sf404063 <> "" Then
			bReturn = False
		End If
		If mp_sf404064 <> "" Then
			bReturn = False
		End If
		If mp_sf404065 <> "" Then
			bReturn = False
		End If
		If mp_sf404066 <> "" Then
			bReturn = False
		End If
		If mp_sf404067 <> "" Then
			bReturn = False
		End If
		If mp_sf404068 <> "" Then
			bReturn = False
		End If
		If mp_sf404069 <> "" Then
			bReturn = False
		End If
		If mp_sf40406a <> "" Then
			bReturn = False
		End If
		If mp_sf40406b <> "" Then
			bReturn = False
		End If
		If mp_sf40406c <> "" Then
			bReturn = False
		End If
		If mp_sf40406d <> "" Then
			bReturn = False
		End If
		If mp_sf40406e <> "" Then
			bReturn = False
		End If
		If mp_sf40406f <> "" Then
			bReturn = False
		End If
		If mp_sf404070 <> "" Then
			bReturn = False
		End If
		If mp_sf404071 <> "" Then
			bReturn = False
		End If
		If mp_sf404072 <> "" Then
			bReturn = False
		End If
		If mp_sf404073 <> "" Then
			bReturn = False
		End If
		If mp_sf404074 <> "" Then
			bReturn = False
		End If
		If mp_sf404075 <> "" Then
			bReturn = False
		End If
		If mp_sf404076 <> "" Then
			bReturn = False
		End If
		If mp_sf404077 <> "" Then
			bReturn = False
		End If
		If mp_sf404078 <> "" Then
			bReturn = False
		End If
		If mp_sf404079 <> "" Then
			bReturn = False
		End If
		If mp_sf40407a <> "" Then
			bReturn = False
		End If
		If mp_sf40407b <> "" Then
			bReturn = False
		End If
		If mp_sf40407c <> "" Then
			bReturn = False
		End If
		If mp_sf40407d <> "" Then
			bReturn = False
		End If
		If mp_sf40407e <> "" Then
			bReturn = False
		End If
		If mp_sf40407f <> "" Then
			bReturn = False
		End If
		If mp_sf404080 <> "" Then
			bReturn = False
		End If
		If mp_sf404081 <> "" Then
			bReturn = False
		End If
		If mp_sf404082 <> "" Then
			bReturn = False
		End If
		If mp_sf404083 <> "" Then
			bReturn = False
		End If
		If mp_sf404084 <> "" Then
			bReturn = False
		End If
		If mp_sf404085 <> "" Then
			bReturn = False
		End If
		If mp_sf404086 <> "" Then
			bReturn = False
		End If
		If mp_sf404087 <> "" Then
			bReturn = False
		End If
		If mp_sf404088 <> "" Then
			bReturn = False
		End If
		If mp_sf404089 <> "" Then
			bReturn = False
		End If
		If mp_sf40408a <> "" Then
			bReturn = False
		End If
		If mp_sf40408b <> "" Then
			bReturn = False
		End If
		If mp_sf40408c <> "" Then
			bReturn = False
		End If
		If mp_sf40408d <> "" Then
			bReturn = False
		End If
		If mp_sf40408e <> "" Then
			bReturn = False
		End If
		If mp_sf40408f <> "" Then
			bReturn = False
		End If
		If mp_sf404090 <> "" Then
			bReturn = False
		End If
		If mp_sf404091 <> "" Then
			bReturn = False
		End If
		If mp_sf404092 <> "" Then
			bReturn = False
		End If
		If mp_sf404093 <> "" Then
			bReturn = False
		End If
		If mp_sf404094 <> "" Then
			bReturn = False
		End If
		If mp_sf404095 <> "" Then
			bReturn = False
		End If
		If mp_sf404096 <> "" Then
			bReturn = False
		End If
		If mp_sf404097 <> "" Then
			bReturn = False
		End If
		If mp_sf404098 <> "" Then
			bReturn = False
		End If
		If mp_sf404099 <> "" Then
			bReturn = False
		End If
		If mp_sf40409a <> "" Then
			bReturn = False
		End If
		If mp_sf40409b <> "" Then
			bReturn = False
		End If
		If mp_sf40409c <> "" Then
			bReturn = False
		End If
		If mp_sf40409d <> "" Then
			bReturn = False
		End If
		If mp_sf40409e <> "" Then
			bReturn = False
		End If
		If mp_sf40409f <> "" Then
			bReturn = False
		End If
		If mp_sf4040a0 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a1 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a2 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a3 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a4 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a5 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a6 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a7 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a8 <> "" Then
			bReturn = False
		End If
		If mp_sf4040a9 <> "" Then
			bReturn = False
		End If
		If mp_sf4040aa <> "" Then
			bReturn = False
		End If
		If mp_sf4040ab <> "" Then
			bReturn = False
		End If
		If mp_sf4040ac <> "" Then
			bReturn = False
		End If
		If mp_sf4040ad <> "" Then
			bReturn = False
		End If
		If mp_sf4040ae <> "" Then
			bReturn = False
		End If
		If mp_sf4040af <> "" Then
			bReturn = False
		End If
		If mp_sf4040b0 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b1 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b2 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b3 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b4 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b5 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b6 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b7 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b8 <> "" Then
			bReturn = False
		End If
		If mp_sf4040b9 <> "" Then
			bReturn = False
		End If
		If mp_sf4040ba <> "" Then
			bReturn = False
		End If
		If mp_sf4040bb <> "" Then
			bReturn = False
		End If
		If mp_sf4040bc <> "" Then
			bReturn = False
		End If
		If mp_sf4040bd <> "" Then
			bReturn = False
		End If
		If mp_sf4040be <> "" Then
			bReturn = False
		End If
		If mp_sf4040bf <> "" Then
			bReturn = False
		End If
		If mp_sf4040c0 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c1 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c2 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c3 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c4 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c5 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c6 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c7 <> "" Then
			bReturn = False
		End If
		If mp_sf4040c8 <> "" Then
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
		oXML.WriteProperty("PeakUnits", mp_fPeakUnits)
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
		oXML.WriteProperty("Summary", mp_bSummary)
		oXML.WriteProperty("SV", mp_fSV)
		oXML.WriteProperty("Units", mp_fUnits)
		oXML.WriteProperty("UpdateNeeded", mp_bUpdateNeeded)
		oXML.WriteProperty("VAC", mp_fVAC)
		oXML.WriteProperty("Work", mp_oWork)
		oXML.WriteProperty("WorkContour", mp_yWorkContour)
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		oXML.WriteProperty("BookingType", mp_yBookingType)
		If mp_oActualWorkProtected.IsNull() = False Then
			oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected)
		End If
		If mp_oActualOvertimeWorkProtected.IsNull() = False Then
			oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		End If
		If mp_dtCreationDate.Ticks <> 0 Then
			oXML.WriteProperty("CreationDate", mp_dtCreationDate)
		End If
		If mp_sAssnOwner <> "" Then
			oXML.WriteProperty("AssnOwner", mp_sAssnOwner)
		End If
		If mp_sAssnOwnerGuid <> "" Then
			oXML.WriteProperty("AssnOwnerGuid", mp_sAssnOwnerGuid)
		End If
		oXML.WriteProperty("BudgetCost", mp_cBudgetCost)
		If mp_oBudgetWork.IsNull() = False Then
			oXML.WriteProperty("BudgetWork", mp_oBudgetWork)
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			mp_oExtendedAttribute_C.WriteObjectProtected(oXML)
		End If
		If mp_oBaseline_C.IsNull() = False Then
			mp_oBaseline_C.WriteObjectProtected(oXML)
		End If
		If mp_sf404000 <> "" Then
			oXML.WriteProperty("f404000", mp_sf404000)
		End If
		If mp_sf404001 <> "" Then
			oXML.WriteProperty("f404001", mp_sf404001)
		End If
		If mp_sf404002 <> "" Then
			oXML.WriteProperty("f404002", mp_sf404002)
		End If
		If mp_sf404003 <> "" Then
			oXML.WriteProperty("f404003", mp_sf404003)
		End If
		If mp_sf404004 <> "" Then
			oXML.WriteProperty("f404004", mp_sf404004)
		End If
		If mp_sf404005 <> "" Then
			oXML.WriteProperty("f404005", mp_sf404005)
		End If
		If mp_sf404006 <> "" Then
			oXML.WriteProperty("f404006", mp_sf404006)
		End If
		If mp_sf404007 <> "" Then
			oXML.WriteProperty("f404007", mp_sf404007)
		End If
		If mp_sf404008 <> "" Then
			oXML.WriteProperty("f404008", mp_sf404008)
		End If
		If mp_sf404009 <> "" Then
			oXML.WriteProperty("f404009", mp_sf404009)
		End If
		If mp_sf40400a <> "" Then
			oXML.WriteProperty("f40400a", mp_sf40400a)
		End If
		If mp_sf40400b <> "" Then
			oXML.WriteProperty("f40400b", mp_sf40400b)
		End If
		If mp_sf40400c <> "" Then
			oXML.WriteProperty("f40400c", mp_sf40400c)
		End If
		If mp_sf40400d <> "" Then
			oXML.WriteProperty("f40400d", mp_sf40400d)
		End If
		If mp_sf40400e <> "" Then
			oXML.WriteProperty("f40400e", mp_sf40400e)
		End If
		If mp_sf40400f <> "" Then
			oXML.WriteProperty("f40400f", mp_sf40400f)
		End If
		If mp_sf404010 <> "" Then
			oXML.WriteProperty("f404010", mp_sf404010)
		End If
		If mp_sf404011 <> "" Then
			oXML.WriteProperty("f404011", mp_sf404011)
		End If
		If mp_sf404012 <> "" Then
			oXML.WriteProperty("f404012", mp_sf404012)
		End If
		If mp_sf404013 <> "" Then
			oXML.WriteProperty("f404013", mp_sf404013)
		End If
		If mp_sf404014 <> "" Then
			oXML.WriteProperty("f404014", mp_sf404014)
		End If
		If mp_sf404015 <> "" Then
			oXML.WriteProperty("f404015", mp_sf404015)
		End If
		If mp_sf404016 <> "" Then
			oXML.WriteProperty("f404016", mp_sf404016)
		End If
		If mp_sf404017 <> "" Then
			oXML.WriteProperty("f404017", mp_sf404017)
		End If
		If mp_sf404018 <> "" Then
			oXML.WriteProperty("f404018", mp_sf404018)
		End If
		If mp_sf404019 <> "" Then
			oXML.WriteProperty("f404019", mp_sf404019)
		End If
		If mp_sf40401a <> "" Then
			oXML.WriteProperty("f40401a", mp_sf40401a)
		End If
		If mp_sf40401b <> "" Then
			oXML.WriteProperty("f40401b", mp_sf40401b)
		End If
		If mp_sf40401c <> "" Then
			oXML.WriteProperty("f40401c", mp_sf40401c)
		End If
		If mp_sf40401d <> "" Then
			oXML.WriteProperty("f40401d", mp_sf40401d)
		End If
		If mp_sf40401e <> "" Then
			oXML.WriteProperty("f40401e", mp_sf40401e)
		End If
		If mp_sf40401f <> "" Then
			oXML.WriteProperty("f40401f", mp_sf40401f)
		End If
		If mp_sf404020 <> "" Then
			oXML.WriteProperty("f404020", mp_sf404020)
		End If
		If mp_sf404021 <> "" Then
			oXML.WriteProperty("f404021", mp_sf404021)
		End If
		If mp_sf404022 <> "" Then
			oXML.WriteProperty("f404022", mp_sf404022)
		End If
		If mp_sf404023 <> "" Then
			oXML.WriteProperty("f404023", mp_sf404023)
		End If
		If mp_sf404024 <> "" Then
			oXML.WriteProperty("f404024", mp_sf404024)
		End If
		If mp_sf404025 <> "" Then
			oXML.WriteProperty("f404025", mp_sf404025)
		End If
		If mp_sf404026 <> "" Then
			oXML.WriteProperty("f404026", mp_sf404026)
		End If
		If mp_sf404027 <> "" Then
			oXML.WriteProperty("f404027", mp_sf404027)
		End If
		If mp_sf404028 <> "" Then
			oXML.WriteProperty("f404028", mp_sf404028)
		End If
		If mp_sf404029 <> "" Then
			oXML.WriteProperty("f404029", mp_sf404029)
		End If
		If mp_sf40402a <> "" Then
			oXML.WriteProperty("f40402a", mp_sf40402a)
		End If
		If mp_sf40402b <> "" Then
			oXML.WriteProperty("f40402b", mp_sf40402b)
		End If
		If mp_sf40402c <> "" Then
			oXML.WriteProperty("f40402c", mp_sf40402c)
		End If
		If mp_sf40402d <> "" Then
			oXML.WriteProperty("f40402d", mp_sf40402d)
		End If
		If mp_sf40402e <> "" Then
			oXML.WriteProperty("f40402e", mp_sf40402e)
		End If
		If mp_sf40402f <> "" Then
			oXML.WriteProperty("f40402f", mp_sf40402f)
		End If
		If mp_sf404030 <> "" Then
			oXML.WriteProperty("f404030", mp_sf404030)
		End If
		If mp_sf404031 <> "" Then
			oXML.WriteProperty("f404031", mp_sf404031)
		End If
		If mp_sf404032 <> "" Then
			oXML.WriteProperty("f404032", mp_sf404032)
		End If
		If mp_sf404033 <> "" Then
			oXML.WriteProperty("f404033", mp_sf404033)
		End If
		If mp_sf404034 <> "" Then
			oXML.WriteProperty("f404034", mp_sf404034)
		End If
		If mp_sf404035 <> "" Then
			oXML.WriteProperty("f404035", mp_sf404035)
		End If
		If mp_sf404036 <> "" Then
			oXML.WriteProperty("f404036", mp_sf404036)
		End If
		If mp_sf404037 <> "" Then
			oXML.WriteProperty("f404037", mp_sf404037)
		End If
		If mp_sf404038 <> "" Then
			oXML.WriteProperty("f404038", mp_sf404038)
		End If
		If mp_sf404039 <> "" Then
			oXML.WriteProperty("f404039", mp_sf404039)
		End If
		If mp_sf40403a <> "" Then
			oXML.WriteProperty("f40403a", mp_sf40403a)
		End If
		If mp_sf40403b <> "" Then
			oXML.WriteProperty("f40403b", mp_sf40403b)
		End If
		If mp_sf40403c <> "" Then
			oXML.WriteProperty("f40403c", mp_sf40403c)
		End If
		If mp_sf40403d <> "" Then
			oXML.WriteProperty("f40403d", mp_sf40403d)
		End If
		If mp_sf40403e <> "" Then
			oXML.WriteProperty("f40403e", mp_sf40403e)
		End If
		If mp_sf40403f <> "" Then
			oXML.WriteProperty("f40403f", mp_sf40403f)
		End If
		If mp_sf404040 <> "" Then
			oXML.WriteProperty("f404040", mp_sf404040)
		End If
		If mp_sf404041 <> "" Then
			oXML.WriteProperty("f404041", mp_sf404041)
		End If
		If mp_sf404042 <> "" Then
			oXML.WriteProperty("f404042", mp_sf404042)
		End If
		If mp_sf404043 <> "" Then
			oXML.WriteProperty("f404043", mp_sf404043)
		End If
		If mp_sf404044 <> "" Then
			oXML.WriteProperty("f404044", mp_sf404044)
		End If
		If mp_sf404045 <> "" Then
			oXML.WriteProperty("f404045", mp_sf404045)
		End If
		If mp_sf404046 <> "" Then
			oXML.WriteProperty("f404046", mp_sf404046)
		End If
		If mp_sf404047 <> "" Then
			oXML.WriteProperty("f404047", mp_sf404047)
		End If
		If mp_sf404048 <> "" Then
			oXML.WriteProperty("f404048", mp_sf404048)
		End If
		If mp_sf404049 <> "" Then
			oXML.WriteProperty("f404049", mp_sf404049)
		End If
		If mp_sf40404a <> "" Then
			oXML.WriteProperty("f40404a", mp_sf40404a)
		End If
		If mp_sf40404b <> "" Then
			oXML.WriteProperty("f40404b", mp_sf40404b)
		End If
		If mp_sf40404c <> "" Then
			oXML.WriteProperty("f40404c", mp_sf40404c)
		End If
		If mp_sf40404d <> "" Then
			oXML.WriteProperty("f40404d", mp_sf40404d)
		End If
		If mp_sf40404e <> "" Then
			oXML.WriteProperty("f40404e", mp_sf40404e)
		End If
		If mp_sf40404f <> "" Then
			oXML.WriteProperty("f40404f", mp_sf40404f)
		End If
		If mp_sf404050 <> "" Then
			oXML.WriteProperty("f404050", mp_sf404050)
		End If
		If mp_sf404051 <> "" Then
			oXML.WriteProperty("f404051", mp_sf404051)
		End If
		If mp_sf404052 <> "" Then
			oXML.WriteProperty("f404052", mp_sf404052)
		End If
		If mp_sf404053 <> "" Then
			oXML.WriteProperty("f404053", mp_sf404053)
		End If
		If mp_sf404054 <> "" Then
			oXML.WriteProperty("f404054", mp_sf404054)
		End If
		If mp_sf404055 <> "" Then
			oXML.WriteProperty("f404055", mp_sf404055)
		End If
		If mp_sf404056 <> "" Then
			oXML.WriteProperty("f404056", mp_sf404056)
		End If
		If mp_sf404057 <> "" Then
			oXML.WriteProperty("f404057", mp_sf404057)
		End If
		If mp_sf404058 <> "" Then
			oXML.WriteProperty("f404058", mp_sf404058)
		End If
		If mp_sf404059 <> "" Then
			oXML.WriteProperty("f404059", mp_sf404059)
		End If
		If mp_sf40405a <> "" Then
			oXML.WriteProperty("f40405a", mp_sf40405a)
		End If
		If mp_sf40405b <> "" Then
			oXML.WriteProperty("f40405b", mp_sf40405b)
		End If
		If mp_sf40405c <> "" Then
			oXML.WriteProperty("f40405c", mp_sf40405c)
		End If
		If mp_sf40405d <> "" Then
			oXML.WriteProperty("f40405d", mp_sf40405d)
		End If
		If mp_sf40405e <> "" Then
			oXML.WriteProperty("f40405e", mp_sf40405e)
		End If
		If mp_sf40405f <> "" Then
			oXML.WriteProperty("f40405f", mp_sf40405f)
		End If
		If mp_sf404060 <> "" Then
			oXML.WriteProperty("f404060", mp_sf404060)
		End If
		If mp_sf404061 <> "" Then
			oXML.WriteProperty("f404061", mp_sf404061)
		End If
		If mp_sf404062 <> "" Then
			oXML.WriteProperty("f404062", mp_sf404062)
		End If
		If mp_sf404063 <> "" Then
			oXML.WriteProperty("f404063", mp_sf404063)
		End If
		If mp_sf404064 <> "" Then
			oXML.WriteProperty("f404064", mp_sf404064)
		End If
		If mp_sf404065 <> "" Then
			oXML.WriteProperty("f404065", mp_sf404065)
		End If
		If mp_sf404066 <> "" Then
			oXML.WriteProperty("f404066", mp_sf404066)
		End If
		If mp_sf404067 <> "" Then
			oXML.WriteProperty("f404067", mp_sf404067)
		End If
		If mp_sf404068 <> "" Then
			oXML.WriteProperty("f404068", mp_sf404068)
		End If
		If mp_sf404069 <> "" Then
			oXML.WriteProperty("f404069", mp_sf404069)
		End If
		If mp_sf40406a <> "" Then
			oXML.WriteProperty("f40406a", mp_sf40406a)
		End If
		If mp_sf40406b <> "" Then
			oXML.WriteProperty("f40406b", mp_sf40406b)
		End If
		If mp_sf40406c <> "" Then
			oXML.WriteProperty("f40406c", mp_sf40406c)
		End If
		If mp_sf40406d <> "" Then
			oXML.WriteProperty("f40406d", mp_sf40406d)
		End If
		If mp_sf40406e <> "" Then
			oXML.WriteProperty("f40406e", mp_sf40406e)
		End If
		If mp_sf40406f <> "" Then
			oXML.WriteProperty("f40406f", mp_sf40406f)
		End If
		If mp_sf404070 <> "" Then
			oXML.WriteProperty("f404070", mp_sf404070)
		End If
		If mp_sf404071 <> "" Then
			oXML.WriteProperty("f404071", mp_sf404071)
		End If
		If mp_sf404072 <> "" Then
			oXML.WriteProperty("f404072", mp_sf404072)
		End If
		If mp_sf404073 <> "" Then
			oXML.WriteProperty("f404073", mp_sf404073)
		End If
		If mp_sf404074 <> "" Then
			oXML.WriteProperty("f404074", mp_sf404074)
		End If
		If mp_sf404075 <> "" Then
			oXML.WriteProperty("f404075", mp_sf404075)
		End If
		If mp_sf404076 <> "" Then
			oXML.WriteProperty("f404076", mp_sf404076)
		End If
		If mp_sf404077 <> "" Then
			oXML.WriteProperty("f404077", mp_sf404077)
		End If
		If mp_sf404078 <> "" Then
			oXML.WriteProperty("f404078", mp_sf404078)
		End If
		If mp_sf404079 <> "" Then
			oXML.WriteProperty("f404079", mp_sf404079)
		End If
		If mp_sf40407a <> "" Then
			oXML.WriteProperty("f40407a", mp_sf40407a)
		End If
		If mp_sf40407b <> "" Then
			oXML.WriteProperty("f40407b", mp_sf40407b)
		End If
		If mp_sf40407c <> "" Then
			oXML.WriteProperty("f40407c", mp_sf40407c)
		End If
		If mp_sf40407d <> "" Then
			oXML.WriteProperty("f40407d", mp_sf40407d)
		End If
		If mp_sf40407e <> "" Then
			oXML.WriteProperty("f40407e", mp_sf40407e)
		End If
		If mp_sf40407f <> "" Then
			oXML.WriteProperty("f40407f", mp_sf40407f)
		End If
		If mp_sf404080 <> "" Then
			oXML.WriteProperty("f404080", mp_sf404080)
		End If
		If mp_sf404081 <> "" Then
			oXML.WriteProperty("f404081", mp_sf404081)
		End If
		If mp_sf404082 <> "" Then
			oXML.WriteProperty("f404082", mp_sf404082)
		End If
		If mp_sf404083 <> "" Then
			oXML.WriteProperty("f404083", mp_sf404083)
		End If
		If mp_sf404084 <> "" Then
			oXML.WriteProperty("f404084", mp_sf404084)
		End If
		If mp_sf404085 <> "" Then
			oXML.WriteProperty("f404085", mp_sf404085)
		End If
		If mp_sf404086 <> "" Then
			oXML.WriteProperty("f404086", mp_sf404086)
		End If
		If mp_sf404087 <> "" Then
			oXML.WriteProperty("f404087", mp_sf404087)
		End If
		If mp_sf404088 <> "" Then
			oXML.WriteProperty("f404088", mp_sf404088)
		End If
		If mp_sf404089 <> "" Then
			oXML.WriteProperty("f404089", mp_sf404089)
		End If
		If mp_sf40408a <> "" Then
			oXML.WriteProperty("f40408a", mp_sf40408a)
		End If
		If mp_sf40408b <> "" Then
			oXML.WriteProperty("f40408b", mp_sf40408b)
		End If
		If mp_sf40408c <> "" Then
			oXML.WriteProperty("f40408c", mp_sf40408c)
		End If
		If mp_sf40408d <> "" Then
			oXML.WriteProperty("f40408d", mp_sf40408d)
		End If
		If mp_sf40408e <> "" Then
			oXML.WriteProperty("f40408e", mp_sf40408e)
		End If
		If mp_sf40408f <> "" Then
			oXML.WriteProperty("f40408f", mp_sf40408f)
		End If
		If mp_sf404090 <> "" Then
			oXML.WriteProperty("f404090", mp_sf404090)
		End If
		If mp_sf404091 <> "" Then
			oXML.WriteProperty("f404091", mp_sf404091)
		End If
		If mp_sf404092 <> "" Then
			oXML.WriteProperty("f404092", mp_sf404092)
		End If
		If mp_sf404093 <> "" Then
			oXML.WriteProperty("f404093", mp_sf404093)
		End If
		If mp_sf404094 <> "" Then
			oXML.WriteProperty("f404094", mp_sf404094)
		End If
		If mp_sf404095 <> "" Then
			oXML.WriteProperty("f404095", mp_sf404095)
		End If
		If mp_sf404096 <> "" Then
			oXML.WriteProperty("f404096", mp_sf404096)
		End If
		If mp_sf404097 <> "" Then
			oXML.WriteProperty("f404097", mp_sf404097)
		End If
		If mp_sf404098 <> "" Then
			oXML.WriteProperty("f404098", mp_sf404098)
		End If
		If mp_sf404099 <> "" Then
			oXML.WriteProperty("f404099", mp_sf404099)
		End If
		If mp_sf40409a <> "" Then
			oXML.WriteProperty("f40409a", mp_sf40409a)
		End If
		If mp_sf40409b <> "" Then
			oXML.WriteProperty("f40409b", mp_sf40409b)
		End If
		If mp_sf40409c <> "" Then
			oXML.WriteProperty("f40409c", mp_sf40409c)
		End If
		If mp_sf40409d <> "" Then
			oXML.WriteProperty("f40409d", mp_sf40409d)
		End If
		If mp_sf40409e <> "" Then
			oXML.WriteProperty("f40409e", mp_sf40409e)
		End If
		If mp_sf40409f <> "" Then
			oXML.WriteProperty("f40409f", mp_sf40409f)
		End If
		If mp_sf4040a0 <> "" Then
			oXML.WriteProperty("f4040a0", mp_sf4040a0)
		End If
		If mp_sf4040a1 <> "" Then
			oXML.WriteProperty("f4040a1", mp_sf4040a1)
		End If
		If mp_sf4040a2 <> "" Then
			oXML.WriteProperty("f4040a2", mp_sf4040a2)
		End If
		If mp_sf4040a3 <> "" Then
			oXML.WriteProperty("f4040a3", mp_sf4040a3)
		End If
		If mp_sf4040a4 <> "" Then
			oXML.WriteProperty("f4040a4", mp_sf4040a4)
		End If
		If mp_sf4040a5 <> "" Then
			oXML.WriteProperty("f4040a5", mp_sf4040a5)
		End If
		If mp_sf4040a6 <> "" Then
			oXML.WriteProperty("f4040a6", mp_sf4040a6)
		End If
		If mp_sf4040a7 <> "" Then
			oXML.WriteProperty("f4040a7", mp_sf4040a7)
		End If
		If mp_sf4040a8 <> "" Then
			oXML.WriteProperty("f4040a8", mp_sf4040a8)
		End If
		If mp_sf4040a9 <> "" Then
			oXML.WriteProperty("f4040a9", mp_sf4040a9)
		End If
		If mp_sf4040aa <> "" Then
			oXML.WriteProperty("f4040aa", mp_sf4040aa)
		End If
		If mp_sf4040ab <> "" Then
			oXML.WriteProperty("f4040ab", mp_sf4040ab)
		End If
		If mp_sf4040ac <> "" Then
			oXML.WriteProperty("f4040ac", mp_sf4040ac)
		End If
		If mp_sf4040ad <> "" Then
			oXML.WriteProperty("f4040ad", mp_sf4040ad)
		End If
		If mp_sf4040ae <> "" Then
			oXML.WriteProperty("f4040ae", mp_sf4040ae)
		End If
		If mp_sf4040af <> "" Then
			oXML.WriteProperty("f4040af", mp_sf4040af)
		End If
		If mp_sf4040b0 <> "" Then
			oXML.WriteProperty("f4040b0", mp_sf4040b0)
		End If
		If mp_sf4040b1 <> "" Then
			oXML.WriteProperty("f4040b1", mp_sf4040b1)
		End If
		If mp_sf4040b2 <> "" Then
			oXML.WriteProperty("f4040b2", mp_sf4040b2)
		End If
		If mp_sf4040b3 <> "" Then
			oXML.WriteProperty("f4040b3", mp_sf4040b3)
		End If
		If mp_sf4040b4 <> "" Then
			oXML.WriteProperty("f4040b4", mp_sf4040b4)
		End If
		If mp_sf4040b5 <> "" Then
			oXML.WriteProperty("f4040b5", mp_sf4040b5)
		End If
		If mp_sf4040b6 <> "" Then
			oXML.WriteProperty("f4040b6", mp_sf4040b6)
		End If
		If mp_sf4040b7 <> "" Then
			oXML.WriteProperty("f4040b7", mp_sf4040b7)
		End If
		If mp_sf4040b8 <> "" Then
			oXML.WriteProperty("f4040b8", mp_sf4040b8)
		End If
		If mp_sf4040b9 <> "" Then
			oXML.WriteProperty("f4040b9", mp_sf4040b9)
		End If
		If mp_sf4040ba <> "" Then
			oXML.WriteProperty("f4040ba", mp_sf4040ba)
		End If
		If mp_sf4040bb <> "" Then
			oXML.WriteProperty("f4040bb", mp_sf4040bb)
		End If
		If mp_sf4040bc <> "" Then
			oXML.WriteProperty("f4040bc", mp_sf4040bc)
		End If
		If mp_sf4040bd <> "" Then
			oXML.WriteProperty("f4040bd", mp_sf4040bd)
		End If
		If mp_sf4040be <> "" Then
			oXML.WriteProperty("f4040be", mp_sf4040be)
		End If
		If mp_sf4040bf <> "" Then
			oXML.WriteProperty("f4040bf", mp_sf4040bf)
		End If
		If mp_sf4040c0 <> "" Then
			oXML.WriteProperty("f4040c0", mp_sf4040c0)
		End If
		If mp_sf4040c1 <> "" Then
			oXML.WriteProperty("f4040c1", mp_sf4040c1)
		End If
		If mp_sf4040c2 <> "" Then
			oXML.WriteProperty("f4040c2", mp_sf4040c2)
		End If
		If mp_sf4040c3 <> "" Then
			oXML.WriteProperty("f4040c3", mp_sf4040c3)
		End If
		If mp_sf4040c4 <> "" Then
			oXML.WriteProperty("f4040c4", mp_sf4040c4)
		End If
		If mp_sf4040c5 <> "" Then
			oXML.WriteProperty("f4040c5", mp_sf4040c5)
		End If
		If mp_sf4040c6 <> "" Then
			oXML.WriteProperty("f4040c6", mp_sf4040c6)
		End If
		If mp_sf4040c7 <> "" Then
			oXML.WriteProperty("f4040c7", mp_sf4040c7)
		End If
		If mp_sf4040c8 <> "" Then
			oXML.WriteProperty("f4040c8", mp_sf4040c8)
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
		oXML.ReadProperty("PeakUnits", mp_fPeakUnits)
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
		oXML.ReadProperty("Summary", mp_bSummary)
		oXML.ReadProperty("SV", mp_fSV)
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
		oXML.ReadProperty("AssnOwner", mp_sAssnOwner)
		oXML.ReadProperty("AssnOwnerGuid", mp_sAssnOwnerGuid)
		oXML.ReadProperty("BudgetCost", mp_cBudgetCost)
		oXML.ReadProperty("BudgetWork", mp_oBudgetWork)
		mp_oExtendedAttribute_C.ReadObjectProtected(oXML)
		mp_oBaseline_C.ReadObjectProtected(oXML)
		oXML.ReadProperty("f404000", mp_sf404000)
		oXML.ReadProperty("f404001", mp_sf404001)
		oXML.ReadProperty("f404002", mp_sf404002)
		oXML.ReadProperty("f404003", mp_sf404003)
		oXML.ReadProperty("f404004", mp_sf404004)
		oXML.ReadProperty("f404005", mp_sf404005)
		oXML.ReadProperty("f404006", mp_sf404006)
		oXML.ReadProperty("f404007", mp_sf404007)
		oXML.ReadProperty("f404008", mp_sf404008)
		oXML.ReadProperty("f404009", mp_sf404009)
		oXML.ReadProperty("f40400a", mp_sf40400a)
		oXML.ReadProperty("f40400b", mp_sf40400b)
		oXML.ReadProperty("f40400c", mp_sf40400c)
		oXML.ReadProperty("f40400d", mp_sf40400d)
		oXML.ReadProperty("f40400e", mp_sf40400e)
		oXML.ReadProperty("f40400f", mp_sf40400f)
		oXML.ReadProperty("f404010", mp_sf404010)
		oXML.ReadProperty("f404011", mp_sf404011)
		oXML.ReadProperty("f404012", mp_sf404012)
		oXML.ReadProperty("f404013", mp_sf404013)
		oXML.ReadProperty("f404014", mp_sf404014)
		oXML.ReadProperty("f404015", mp_sf404015)
		oXML.ReadProperty("f404016", mp_sf404016)
		oXML.ReadProperty("f404017", mp_sf404017)
		oXML.ReadProperty("f404018", mp_sf404018)
		oXML.ReadProperty("f404019", mp_sf404019)
		oXML.ReadProperty("f40401a", mp_sf40401a)
		oXML.ReadProperty("f40401b", mp_sf40401b)
		oXML.ReadProperty("f40401c", mp_sf40401c)
		oXML.ReadProperty("f40401d", mp_sf40401d)
		oXML.ReadProperty("f40401e", mp_sf40401e)
		oXML.ReadProperty("f40401f", mp_sf40401f)
		oXML.ReadProperty("f404020", mp_sf404020)
		oXML.ReadProperty("f404021", mp_sf404021)
		oXML.ReadProperty("f404022", mp_sf404022)
		oXML.ReadProperty("f404023", mp_sf404023)
		oXML.ReadProperty("f404024", mp_sf404024)
		oXML.ReadProperty("f404025", mp_sf404025)
		oXML.ReadProperty("f404026", mp_sf404026)
		oXML.ReadProperty("f404027", mp_sf404027)
		oXML.ReadProperty("f404028", mp_sf404028)
		oXML.ReadProperty("f404029", mp_sf404029)
		oXML.ReadProperty("f40402a", mp_sf40402a)
		oXML.ReadProperty("f40402b", mp_sf40402b)
		oXML.ReadProperty("f40402c", mp_sf40402c)
		oXML.ReadProperty("f40402d", mp_sf40402d)
		oXML.ReadProperty("f40402e", mp_sf40402e)
		oXML.ReadProperty("f40402f", mp_sf40402f)
		oXML.ReadProperty("f404030", mp_sf404030)
		oXML.ReadProperty("f404031", mp_sf404031)
		oXML.ReadProperty("f404032", mp_sf404032)
		oXML.ReadProperty("f404033", mp_sf404033)
		oXML.ReadProperty("f404034", mp_sf404034)
		oXML.ReadProperty("f404035", mp_sf404035)
		oXML.ReadProperty("f404036", mp_sf404036)
		oXML.ReadProperty("f404037", mp_sf404037)
		oXML.ReadProperty("f404038", mp_sf404038)
		oXML.ReadProperty("f404039", mp_sf404039)
		oXML.ReadProperty("f40403a", mp_sf40403a)
		oXML.ReadProperty("f40403b", mp_sf40403b)
		oXML.ReadProperty("f40403c", mp_sf40403c)
		oXML.ReadProperty("f40403d", mp_sf40403d)
		oXML.ReadProperty("f40403e", mp_sf40403e)
		oXML.ReadProperty("f40403f", mp_sf40403f)
		oXML.ReadProperty("f404040", mp_sf404040)
		oXML.ReadProperty("f404041", mp_sf404041)
		oXML.ReadProperty("f404042", mp_sf404042)
		oXML.ReadProperty("f404043", mp_sf404043)
		oXML.ReadProperty("f404044", mp_sf404044)
		oXML.ReadProperty("f404045", mp_sf404045)
		oXML.ReadProperty("f404046", mp_sf404046)
		oXML.ReadProperty("f404047", mp_sf404047)
		oXML.ReadProperty("f404048", mp_sf404048)
		oXML.ReadProperty("f404049", mp_sf404049)
		oXML.ReadProperty("f40404a", mp_sf40404a)
		oXML.ReadProperty("f40404b", mp_sf40404b)
		oXML.ReadProperty("f40404c", mp_sf40404c)
		oXML.ReadProperty("f40404d", mp_sf40404d)
		oXML.ReadProperty("f40404e", mp_sf40404e)
		oXML.ReadProperty("f40404f", mp_sf40404f)
		oXML.ReadProperty("f404050", mp_sf404050)
		oXML.ReadProperty("f404051", mp_sf404051)
		oXML.ReadProperty("f404052", mp_sf404052)
		oXML.ReadProperty("f404053", mp_sf404053)
		oXML.ReadProperty("f404054", mp_sf404054)
		oXML.ReadProperty("f404055", mp_sf404055)
		oXML.ReadProperty("f404056", mp_sf404056)
		oXML.ReadProperty("f404057", mp_sf404057)
		oXML.ReadProperty("f404058", mp_sf404058)
		oXML.ReadProperty("f404059", mp_sf404059)
		oXML.ReadProperty("f40405a", mp_sf40405a)
		oXML.ReadProperty("f40405b", mp_sf40405b)
		oXML.ReadProperty("f40405c", mp_sf40405c)
		oXML.ReadProperty("f40405d", mp_sf40405d)
		oXML.ReadProperty("f40405e", mp_sf40405e)
		oXML.ReadProperty("f40405f", mp_sf40405f)
		oXML.ReadProperty("f404060", mp_sf404060)
		oXML.ReadProperty("f404061", mp_sf404061)
		oXML.ReadProperty("f404062", mp_sf404062)
		oXML.ReadProperty("f404063", mp_sf404063)
		oXML.ReadProperty("f404064", mp_sf404064)
		oXML.ReadProperty("f404065", mp_sf404065)
		oXML.ReadProperty("f404066", mp_sf404066)
		oXML.ReadProperty("f404067", mp_sf404067)
		oXML.ReadProperty("f404068", mp_sf404068)
		oXML.ReadProperty("f404069", mp_sf404069)
		oXML.ReadProperty("f40406a", mp_sf40406a)
		oXML.ReadProperty("f40406b", mp_sf40406b)
		oXML.ReadProperty("f40406c", mp_sf40406c)
		oXML.ReadProperty("f40406d", mp_sf40406d)
		oXML.ReadProperty("f40406e", mp_sf40406e)
		oXML.ReadProperty("f40406f", mp_sf40406f)
		oXML.ReadProperty("f404070", mp_sf404070)
		oXML.ReadProperty("f404071", mp_sf404071)
		oXML.ReadProperty("f404072", mp_sf404072)
		oXML.ReadProperty("f404073", mp_sf404073)
		oXML.ReadProperty("f404074", mp_sf404074)
		oXML.ReadProperty("f404075", mp_sf404075)
		oXML.ReadProperty("f404076", mp_sf404076)
		oXML.ReadProperty("f404077", mp_sf404077)
		oXML.ReadProperty("f404078", mp_sf404078)
		oXML.ReadProperty("f404079", mp_sf404079)
		oXML.ReadProperty("f40407a", mp_sf40407a)
		oXML.ReadProperty("f40407b", mp_sf40407b)
		oXML.ReadProperty("f40407c", mp_sf40407c)
		oXML.ReadProperty("f40407d", mp_sf40407d)
		oXML.ReadProperty("f40407e", mp_sf40407e)
		oXML.ReadProperty("f40407f", mp_sf40407f)
		oXML.ReadProperty("f404080", mp_sf404080)
		oXML.ReadProperty("f404081", mp_sf404081)
		oXML.ReadProperty("f404082", mp_sf404082)
		oXML.ReadProperty("f404083", mp_sf404083)
		oXML.ReadProperty("f404084", mp_sf404084)
		oXML.ReadProperty("f404085", mp_sf404085)
		oXML.ReadProperty("f404086", mp_sf404086)
		oXML.ReadProperty("f404087", mp_sf404087)
		oXML.ReadProperty("f404088", mp_sf404088)
		oXML.ReadProperty("f404089", mp_sf404089)
		oXML.ReadProperty("f40408a", mp_sf40408a)
		oXML.ReadProperty("f40408b", mp_sf40408b)
		oXML.ReadProperty("f40408c", mp_sf40408c)
		oXML.ReadProperty("f40408d", mp_sf40408d)
		oXML.ReadProperty("f40408e", mp_sf40408e)
		oXML.ReadProperty("f40408f", mp_sf40408f)
		oXML.ReadProperty("f404090", mp_sf404090)
		oXML.ReadProperty("f404091", mp_sf404091)
		oXML.ReadProperty("f404092", mp_sf404092)
		oXML.ReadProperty("f404093", mp_sf404093)
		oXML.ReadProperty("f404094", mp_sf404094)
		oXML.ReadProperty("f404095", mp_sf404095)
		oXML.ReadProperty("f404096", mp_sf404096)
		oXML.ReadProperty("f404097", mp_sf404097)
		oXML.ReadProperty("f404098", mp_sf404098)
		oXML.ReadProperty("f404099", mp_sf404099)
		oXML.ReadProperty("f40409a", mp_sf40409a)
		oXML.ReadProperty("f40409b", mp_sf40409b)
		oXML.ReadProperty("f40409c", mp_sf40409c)
		oXML.ReadProperty("f40409d", mp_sf40409d)
		oXML.ReadProperty("f40409e", mp_sf40409e)
		oXML.ReadProperty("f40409f", mp_sf40409f)
		oXML.ReadProperty("f4040a0", mp_sf4040a0)
		oXML.ReadProperty("f4040a1", mp_sf4040a1)
		oXML.ReadProperty("f4040a2", mp_sf4040a2)
		oXML.ReadProperty("f4040a3", mp_sf4040a3)
		oXML.ReadProperty("f4040a4", mp_sf4040a4)
		oXML.ReadProperty("f4040a5", mp_sf4040a5)
		oXML.ReadProperty("f4040a6", mp_sf4040a6)
		oXML.ReadProperty("f4040a7", mp_sf4040a7)
		oXML.ReadProperty("f4040a8", mp_sf4040a8)
		oXML.ReadProperty("f4040a9", mp_sf4040a9)
		oXML.ReadProperty("f4040aa", mp_sf4040aa)
		oXML.ReadProperty("f4040ab", mp_sf4040ab)
		oXML.ReadProperty("f4040ac", mp_sf4040ac)
		oXML.ReadProperty("f4040ad", mp_sf4040ad)
		oXML.ReadProperty("f4040ae", mp_sf4040ae)
		oXML.ReadProperty("f4040af", mp_sf4040af)
		oXML.ReadProperty("f4040b0", mp_sf4040b0)
		oXML.ReadProperty("f4040b1", mp_sf4040b1)
		oXML.ReadProperty("f4040b2", mp_sf4040b2)
		oXML.ReadProperty("f4040b3", mp_sf4040b3)
		oXML.ReadProperty("f4040b4", mp_sf4040b4)
		oXML.ReadProperty("f4040b5", mp_sf4040b5)
		oXML.ReadProperty("f4040b6", mp_sf4040b6)
		oXML.ReadProperty("f4040b7", mp_sf4040b7)
		oXML.ReadProperty("f4040b8", mp_sf4040b8)
		oXML.ReadProperty("f4040b9", mp_sf4040b9)
		oXML.ReadProperty("f4040ba", mp_sf4040ba)
		oXML.ReadProperty("f4040bb", mp_sf4040bb)
		oXML.ReadProperty("f4040bc", mp_sf4040bc)
		oXML.ReadProperty("f4040bd", mp_sf4040bd)
		oXML.ReadProperty("f4040be", mp_sf4040be)
		oXML.ReadProperty("f4040bf", mp_sf4040bf)
		oXML.ReadProperty("f4040c0", mp_sf4040c0)
		oXML.ReadProperty("f4040c1", mp_sf4040c1)
		oXML.ReadProperty("f4040c2", mp_sf4040c2)
		oXML.ReadProperty("f4040c3", mp_sf4040c3)
		oXML.ReadProperty("f4040c4", mp_sf4040c4)
		oXML.ReadProperty("f4040c5", mp_sf4040c5)
		oXML.ReadProperty("f4040c6", mp_sf4040c6)
		oXML.ReadProperty("f4040c7", mp_sf4040c7)
		oXML.ReadProperty("f4040c8", mp_sf4040c8)
		mp_oTimephasedData_C.ReadObjectProtected(oXML)
	End Sub

End Class
