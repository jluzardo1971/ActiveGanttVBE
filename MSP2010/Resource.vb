Option Explicit On

Public Class Resource
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_lID As Integer
	Private mp_sName As String
	Private mp_yType As E_TYPE_6
	Private mp_bIsNull As Boolean
	Private mp_sInitials As String
	Private mp_sPhonetics As String
	Private mp_sNTAccount As String
	Private mp_sMaterialLabel As String
	Private mp_sCode As String
	Private mp_sGroup As String
	Private mp_yWorkGroup As E_WORKGROUP
	Private mp_sEmailAddress As String
	Private mp_sHyperlink As String
	Private mp_sHyperlinkAddress As String
	Private mp_sHyperlinkSubAddress As String
	Private mp_fMaxUnits As Single
	Private mp_fPeakUnits As Single
	Private mp_bOverAllocated As Boolean
	Private mp_dtAvailableFrom As System.DateTime
	Private mp_dtAvailableTo As System.DateTime
	Private mp_dtStart As System.DateTime
	Private mp_dtFinish As System.DateTime
	Private mp_bCanLevel As Boolean
	Private mp_yAccrueAt As E_ACCRUEAT
	Private mp_oWork As Duration
	Private mp_oRegularWork As Duration
	Private mp_oOvertimeWork As Duration
	Private mp_oActualWork As Duration
	Private mp_oRemainingWork As Duration
	Private mp_oActualOvertimeWork As Duration
	Private mp_oRemainingOvertimeWork As Duration
	Private mp_lPercentWorkComplete As Integer
	Private mp_cStandardRate As Decimal
	Private mp_yStandardRateFormat As E_STANDARDRATEFORMAT
	Private mp_cCost As Decimal
	Private mp_cOvertimeRate As Decimal
	Private mp_yOvertimeRateFormat As E_OVERTIMERATEFORMAT
	Private mp_cOvertimeCost As Decimal
	Private mp_cCostPerUse As Decimal
	Private mp_cActualCost As Decimal
	Private mp_cActualOvertimeCost As Decimal
	Private mp_cRemainingCost As Decimal
	Private mp_cRemainingOvertimeCost As Decimal
	Private mp_fWorkVariance As Single
	Private mp_fCostVariance As Single
	Private mp_fSV As Single
	Private mp_fCV As Single
	Private mp_fACWP As Single
	Private mp_lCalendarUID As Integer
	Private mp_sNotes As String
	Private mp_fBCWS As Single
	Private mp_fBCWP As Single
	Private mp_bIsGeneric As Boolean
	Private mp_bIsInactive As Boolean
	Private mp_bIsEnterprise As Boolean
	Private mp_yBookingType As E_BOOKINGTYPE
	Private mp_oActualWorkProtected As Duration
	Private mp_oActualOvertimeWorkProtected As Duration
	Private mp_sActiveDirectoryGUID As String
	Private mp_dtCreationDate As System.DateTime
	Private mp_oExtendedAttribute_C As ResourceExtendedAttribute_C
	Private mp_oBaseline_C As ResourceBaseline_C
	Private mp_oOutlineCode_C As ResourceOutlineCode_C
	Private mp_bIsCostResource As Boolean
	Private mp_sAssnOwner As String
	Private mp_sAssnOwnerGuid As String
	Private mp_bIsBudget As Boolean
	Private mp_oAvailabilityPeriods As ResourceAvailabilityPeriods
	Private mp_oRates As ResourceRates
	Private mp_oTimephasedData_C As TimephasedData_C

	Public Sub New()
		mp_lUID = 0
		mp_lID = 0
		mp_sName = ""
		mp_yType = E_TYPE_6.T_6_MATERIAL
		mp_bIsNull = False
		mp_sInitials = ""
		mp_sPhonetics = ""
		mp_sNTAccount = ""
		mp_sMaterialLabel = ""
		mp_sCode = ""
		mp_sGroup = ""
		mp_yWorkGroup = E_WORKGROUP.WG_DEFAULT
		mp_sEmailAddress = ""
		mp_sHyperlink = ""
		mp_sHyperlinkAddress = ""
		mp_sHyperlinkSubAddress = ""
		mp_fMaxUnits = System.Convert.ToSingle("1.0")
		mp_fPeakUnits = 0
		mp_bOverAllocated = False
		mp_dtAvailableFrom = New System.DateTime(0)
		mp_dtAvailableTo = New System.DateTime(0)
		mp_dtStart = New System.DateTime(0)
		mp_dtFinish = New System.DateTime(0)
		mp_bCanLevel = False
		mp_yAccrueAt = E_ACCRUEAT.AA_START
		mp_oWork = New Duration()
		mp_oRegularWork = New Duration()
		mp_oOvertimeWork = New Duration()
		mp_oActualWork = New Duration()
		mp_oRemainingWork = New Duration()
		mp_oActualOvertimeWork = New Duration()
		mp_oRemainingOvertimeWork = New Duration()
		mp_lPercentWorkComplete = 0
		mp_cStandardRate = 0
		mp_yStandardRateFormat = E_STANDARDRATEFORMAT.SRF_M
		mp_cCost = 0
		mp_cOvertimeRate = 0
		mp_yOvertimeRateFormat = E_OVERTIMERATEFORMAT.ORF_M
		mp_cOvertimeCost = 0
		mp_cCostPerUse = 0
		mp_cActualCost = 0
		mp_cActualOvertimeCost = 0
		mp_cRemainingCost = 0
		mp_cRemainingOvertimeCost = 0
		mp_fWorkVariance = 0
		mp_fCostVariance = 0
		mp_fSV = 0
		mp_fCV = 0
		mp_fACWP = 0
		mp_lCalendarUID = 0
		mp_sNotes = ""
		mp_fBCWS = 0
		mp_fBCWP = 0
		mp_bIsGeneric = False
		mp_bIsInactive = False
		mp_bIsEnterprise = False
		mp_yBookingType = E_BOOKINGTYPE.BT_COMMITED
		mp_oActualWorkProtected = New Duration()
		mp_oActualOvertimeWorkProtected = New Duration()
		mp_sActiveDirectoryGUID = ""
		mp_dtCreationDate = New System.DateTime(0)
		mp_oExtendedAttribute_C = New ResourceExtendedAttribute_C()
		mp_oBaseline_C = New ResourceBaseline_C()
		mp_oOutlineCode_C = New ResourceOutlineCode_C()
		mp_bIsCostResource = False
		mp_sAssnOwner = ""
		mp_sAssnOwnerGuid = ""
		mp_bIsBudget = False
		mp_oAvailabilityPeriods = New ResourceAvailabilityPeriods()
		mp_oRates = New ResourceRates()
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

	Public Property yType() As E_TYPE_6
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE_6)
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

	Public Property sInitials() As String
		Get
			Return mp_sInitials
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sInitials = Value
		End Set
	End Property

	Public Property sPhonetics() As String
		Get
			Return mp_sPhonetics
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sPhonetics = Value
		End Set
	End Property

	Public Property sNTAccount() As String
		Get
			Return mp_sNTAccount
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sNTAccount = Value
		End Set
	End Property

	Public Property sMaterialLabel() As String
		Get
			Return mp_sMaterialLabel
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sMaterialLabel = Value
		End Set
	End Property

	Public Property sCode() As String
		Get
			Return mp_sCode
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sCode = Value
		End Set
	End Property

	Public Property sGroup() As String
		Get
			Return mp_sGroup
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sGroup = Value
		End Set
	End Property

	Public Property yWorkGroup() As E_WORKGROUP
		Get
			Return mp_yWorkGroup
		End Get
		Set(ByVal Value As E_WORKGROUP)
			mp_yWorkGroup = Value
		End Set
	End Property

	Public Property sEmailAddress() As String
		Get
			Return mp_sEmailAddress
		End Get
		Set(ByVal Value As String)
			If Value.Length > 512 Then
				Value = Value.Substring(0, 512)
			End If
			mp_sEmailAddress = Value
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

	Public Property fMaxUnits() As Single
		Get
			Return mp_fMaxUnits
		End Get
		Set(ByVal Value As Single)
			mp_fMaxUnits = Value
		End Set
	End Property

	Public Property fPeakUnits() As Single
		Get
			Return mp_fPeakUnits
		End Get
		Set(ByVal Value As Single)
			mp_fPeakUnits = Value
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

	Public Property dtAvailableFrom() As System.DateTime
		Get
			Return mp_dtAvailableFrom
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtAvailableFrom = Value
		End Set
	End Property

	Public Property dtAvailableTo() As System.DateTime
		Get
			Return mp_dtAvailableTo
		End Get
		Set(ByVal Value As System.DateTime)
			mp_dtAvailableTo = Value
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

	Public Property bCanLevel() As Boolean
		Get
			Return mp_bCanLevel
		End Get
		Set(ByVal Value As Boolean)
			mp_bCanLevel = Value
		End Set
	End Property

	Public Property yAccrueAt() As E_ACCRUEAT
		Get
			Return mp_yAccrueAt
		End Get
		Set(ByVal Value As E_ACCRUEAT)
			mp_yAccrueAt = Value
		End Set
	End Property

	Public ReadOnly Property oWork() As Duration
		Get
			Return mp_oWork
		End Get
	End Property

	Public ReadOnly Property oRegularWork() As Duration
		Get
			Return mp_oRegularWork
		End Get
	End Property

	Public ReadOnly Property oOvertimeWork() As Duration
		Get
			Return mp_oOvertimeWork
		End Get
	End Property

	Public ReadOnly Property oActualWork() As Duration
		Get
			Return mp_oActualWork
		End Get
	End Property

	Public ReadOnly Property oRemainingWork() As Duration
		Get
			Return mp_oRemainingWork
		End Get
	End Property

	Public ReadOnly Property oActualOvertimeWork() As Duration
		Get
			Return mp_oActualOvertimeWork
		End Get
	End Property

	Public ReadOnly Property oRemainingOvertimeWork() As Duration
		Get
			Return mp_oRemainingOvertimeWork
		End Get
	End Property

	Public Property lPercentWorkComplete() As Integer
		Get
			Return mp_lPercentWorkComplete
		End Get
		Set(ByVal Value As Integer)
			mp_lPercentWorkComplete = Value
		End Set
	End Property

	Public Property cStandardRate() As Decimal
		Get
			Return mp_cStandardRate
		End Get
		Set(ByVal Value As Decimal)
			mp_cStandardRate = Value
		End Set
	End Property

	Public Property yStandardRateFormat() As E_STANDARDRATEFORMAT
		Get
			Return mp_yStandardRateFormat
		End Get
		Set(ByVal Value As E_STANDARDRATEFORMAT)
			mp_yStandardRateFormat = Value
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

	Public Property cOvertimeRate() As Decimal
		Get
			Return mp_cOvertimeRate
		End Get
		Set(ByVal Value As Decimal)
			mp_cOvertimeRate = Value
		End Set
	End Property

	Public Property yOvertimeRateFormat() As E_OVERTIMERATEFORMAT
		Get
			Return mp_yOvertimeRateFormat
		End Get
		Set(ByVal Value As E_OVERTIMERATEFORMAT)
			mp_yOvertimeRateFormat = Value
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

	Public Property cCostPerUse() As Decimal
		Get
			Return mp_cCostPerUse
		End Get
		Set(ByVal Value As Decimal)
			mp_cCostPerUse = Value
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

	Public Property cActualOvertimeCost() As Decimal
		Get
			Return mp_cActualOvertimeCost
		End Get
		Set(ByVal Value As Decimal)
			mp_cActualOvertimeCost = Value
		End Set
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

	Public Property fWorkVariance() As Single
		Get
			Return mp_fWorkVariance
		End Get
		Set(ByVal Value As Single)
			mp_fWorkVariance = Value
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

	Public Property fSV() As Single
		Get
			Return mp_fSV
		End Get
		Set(ByVal Value As Single)
			mp_fSV = Value
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

	Public Property fACWP() As Single
		Get
			Return mp_fACWP
		End Get
		Set(ByVal Value As Single)
			mp_fACWP = Value
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

	Public Property sNotes() As String
		Get
			Return mp_sNotes
		End Get
		Set(ByVal Value As String)
			mp_sNotes = Value
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

	Public Property bIsGeneric() As Boolean
		Get
			Return mp_bIsGeneric
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsGeneric = Value
		End Set
	End Property

	Public Property bIsInactive() As Boolean
		Get
			Return mp_bIsInactive
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsInactive = Value
		End Set
	End Property

	Public Property bIsEnterprise() As Boolean
		Get
			Return mp_bIsEnterprise
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsEnterprise = Value
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

	Public Property sActiveDirectoryGUID() As String
		Get
			Return mp_sActiveDirectoryGUID
		End Get
		Set(ByVal Value As String)
			If Value.Length > 16 Then
				Value = Value.Substring(0, 16)
			End If
			mp_sActiveDirectoryGUID = Value
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

	Public ReadOnly Property oExtendedAttribute_C() As ResourceExtendedAttribute_C
		Get
			Return mp_oExtendedAttribute_C
		End Get
	End Property

	Public ReadOnly Property oBaseline_C() As ResourceBaseline_C
		Get
			Return mp_oBaseline_C
		End Get
	End Property

	Public ReadOnly Property oOutlineCode_C() As ResourceOutlineCode_C
		Get
			Return mp_oOutlineCode_C
		End Get
	End Property

	Public Property bIsCostResource() As Boolean
		Get
			Return mp_bIsCostResource
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsCostResource = Value
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

	Public Property bIsBudget() As Boolean
		Get
			Return mp_bIsBudget
		End Get
		Set(ByVal Value As Boolean)
			mp_bIsBudget = Value
		End Set
	End Property

	Public ReadOnly Property oAvailabilityPeriods() As ResourceAvailabilityPeriods
		Get
			Return mp_oAvailabilityPeriods
		End Get
	End Property

	Public ReadOnly Property oRates() As ResourceRates
		Get
			Return mp_oRates
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
		If mp_lID <> 0 Then
			bReturn = False
		End If
		If mp_sName <> "" Then
			bReturn = False
		End If
		If mp_yType <> E_TYPE_6.T_6_MATERIAL Then
			bReturn = False
		End If
		If mp_bIsNull <> False Then
			bReturn = False
		End If
		If mp_sInitials <> "" Then
			bReturn = False
		End If
		If mp_sPhonetics <> "" Then
			bReturn = False
		End If
		If mp_sNTAccount <> "" Then
			bReturn = False
		End If
		If mp_sMaterialLabel <> "" Then
			bReturn = False
		End If
		If mp_sCode <> "" Then
			bReturn = False
		End If
		If mp_sGroup <> "" Then
			bReturn = False
		End If
		If mp_yWorkGroup <> E_WORKGROUP.WG_DEFAULT Then
			bReturn = False
		End If
		If mp_sEmailAddress <> "" Then
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
		If mp_fMaxUnits <> 1.0 Then
			bReturn = False
		End If
		If mp_fPeakUnits <> 0 Then
			bReturn = False
		End If
		If mp_bOverAllocated <> False Then
			bReturn = False
		End If
		If mp_dtAvailableFrom.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtAvailableTo.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtStart.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_dtFinish.Ticks <> 0 Then
			bReturn = False
		End If
		If mp_bCanLevel <> False Then
			bReturn = False
		End If
		If mp_yAccrueAt <> E_ACCRUEAT.AA_START Then
			bReturn = False
		End If
		If mp_oWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRegularWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRemainingWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oActualOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRemainingOvertimeWork.IsNull() = False Then
			bReturn = False
		End If
		If mp_lPercentWorkComplete <> 0 Then
			bReturn = False
		End If
		If mp_cStandardRate <> 0 Then
			bReturn = False
		End If
		If mp_yStandardRateFormat <> E_STANDARDRATEFORMAT.SRF_M Then
			bReturn = False
		End If
		If mp_cCost <> 0 Then
			bReturn = False
		End If
		If mp_cOvertimeRate <> 0 Then
			bReturn = False
		End If
		If mp_yOvertimeRateFormat <> E_OVERTIMERATEFORMAT.ORF_M Then
			bReturn = False
		End If
		If mp_cOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_cCostPerUse <> 0 Then
			bReturn = False
		End If
		If mp_cActualCost <> 0 Then
			bReturn = False
		End If
		If mp_cActualOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_cRemainingCost <> 0 Then
			bReturn = False
		End If
		If mp_cRemainingOvertimeCost <> 0 Then
			bReturn = False
		End If
		If mp_fWorkVariance <> 0 Then
			bReturn = False
		End If
		If mp_fCostVariance <> 0 Then
			bReturn = False
		End If
		If mp_fSV <> 0 Then
			bReturn = False
		End If
		If mp_fCV <> 0 Then
			bReturn = False
		End If
		If mp_fACWP <> 0 Then
			bReturn = False
		End If
		If mp_lCalendarUID <> 0 Then
			bReturn = False
		End If
		If mp_sNotes <> "" Then
			bReturn = False
		End If
		If mp_fBCWS <> 0 Then
			bReturn = False
		End If
		If mp_fBCWP <> 0 Then
			bReturn = False
		End If
		If mp_bIsGeneric <> False Then
			bReturn = False
		End If
		If mp_bIsInactive <> False Then
			bReturn = False
		End If
		If mp_bIsEnterprise <> False Then
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
		If mp_sActiveDirectoryGUID <> "" Then
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
		If mp_oOutlineCode_C.IsNull() = False Then
			bReturn = False
		End If
		If mp_bIsCostResource <> False Then
			bReturn = False
		End If
		If mp_sAssnOwner <> "" Then
			bReturn = False
		End If
		If mp_sAssnOwnerGuid <> "" Then
			bReturn = False
		End If
		If mp_bIsBudget <> False Then
			bReturn = False
		End If
		If mp_oAvailabilityPeriods.IsNull() = False Then
			bReturn = False
		End If
		If mp_oRates.IsNull() = False Then
			bReturn = False
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Resource/>"
		End if
		Dim oXML As New clsXML("Resource")
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
		If mp_sInitials <> "" Then
			oXML.WriteProperty("Initials", mp_sInitials)
		End If
		If mp_sPhonetics <> "" Then
			oXML.WriteProperty("Phonetics", mp_sPhonetics)
		End If
		If mp_sNTAccount <> "" Then
			oXML.WriteProperty("NTAccount", mp_sNTAccount)
		End If
		If mp_sMaterialLabel <> "" Then
			oXML.WriteProperty("MaterialLabel", mp_sMaterialLabel)
		End If
		If mp_sCode <> "" Then
			oXML.WriteProperty("Code", mp_sCode)
		End If
		If mp_sGroup <> "" Then
			oXML.WriteProperty("Group", mp_sGroup)
		End If
		oXML.WriteProperty("WorkGroup", mp_yWorkGroup)
		If mp_sEmailAddress <> "" Then
			oXML.WriteProperty("EmailAddress", mp_sEmailAddress)
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
		oXML.WriteProperty("MaxUnits", mp_fMaxUnits)
		oXML.WriteProperty("PeakUnits", mp_fPeakUnits)
		oXML.WriteProperty("OverAllocated", mp_bOverAllocated)
		If mp_dtAvailableFrom.Ticks <> 0 Then
			oXML.WriteProperty("AvailableFrom", mp_dtAvailableFrom)
		End If
		If mp_dtAvailableTo.Ticks <> 0 Then
			oXML.WriteProperty("AvailableTo", mp_dtAvailableTo)
		End If
		If mp_dtStart.Ticks <> 0 Then
			oXML.WriteProperty("Start", mp_dtStart)
		End If
		If mp_dtFinish.Ticks <> 0 Then
			oXML.WriteProperty("Finish", mp_dtFinish)
		End If
		oXML.WriteProperty("CanLevel", mp_bCanLevel)
		oXML.WriteProperty("AccrueAt", mp_yAccrueAt)
		oXML.WriteProperty("Work", mp_oWork)
		oXML.WriteProperty("RegularWork", mp_oRegularWork)
		oXML.WriteProperty("OvertimeWork", mp_oOvertimeWork)
		oXML.WriteProperty("ActualWork", mp_oActualWork)
		oXML.WriteProperty("RemainingWork", mp_oRemainingWork)
		oXML.WriteProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		oXML.WriteProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.WriteProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.WriteProperty("StandardRate", mp_cStandardRate)
		oXML.WriteProperty("StandardRateFormat", mp_yStandardRateFormat)
		oXML.WriteProperty("Cost", mp_cCost)
		oXML.WriteProperty("OvertimeRate", mp_cOvertimeRate)
		oXML.WriteProperty("OvertimeRateFormat", mp_yOvertimeRateFormat)
		oXML.WriteProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.WriteProperty("CostPerUse", mp_cCostPerUse)
		oXML.WriteProperty("ActualCost", mp_cActualCost)
		oXML.WriteProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.WriteProperty("RemainingCost", mp_cRemainingCost)
		oXML.WriteProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.WriteProperty("WorkVariance", mp_fWorkVariance)
		oXML.WriteProperty("CostVariance", mp_fCostVariance)
		oXML.WriteProperty("SV", mp_fSV)
		oXML.WriteProperty("CV", mp_fCV)
		oXML.WriteProperty("ACWP", mp_fACWP)
		oXML.WriteProperty("CalendarUID", mp_lCalendarUID)
		If mp_sNotes <> "" Then
			oXML.WriteProperty("Notes", mp_sNotes)
		End If
		oXML.WriteProperty("BCWS", mp_fBCWS)
		oXML.WriteProperty("BCWP", mp_fBCWP)
		oXML.WriteProperty("IsGeneric", mp_bIsGeneric)
		oXML.WriteProperty("IsInactive", mp_bIsInactive)
		oXML.WriteProperty("IsEnterprise", mp_bIsEnterprise)
		oXML.WriteProperty("BookingType", mp_yBookingType)
		oXML.WriteProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.WriteProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		If mp_sActiveDirectoryGUID <> "" Then
			oXML.WriteProperty("ActiveDirectoryGUID", mp_sActiveDirectoryGUID)
		End If
		If mp_dtCreationDate.Ticks <> 0 Then
			oXML.WriteProperty("CreationDate", mp_dtCreationDate)
		End If
		If mp_oExtendedAttribute_C.IsNull() = False Then
			mp_oExtendedAttribute_C.WriteObjectProtected(oXML)
		End If
		If mp_oBaseline_C.IsNull() = False Then
			mp_oBaseline_C.WriteObjectProtected(oXML)
		End If
		If mp_oOutlineCode_C.IsNull() = False Then
			mp_oOutlineCode_C.WriteObjectProtected(oXML)
		End If
		oXML.WriteProperty("IsCostResource", mp_bIsCostResource)
		If mp_sAssnOwner <> "" Then
			oXML.WriteProperty("AssnOwner", mp_sAssnOwner)
		End If
		If mp_sAssnOwnerGuid <> "" Then
			oXML.WriteProperty("AssnOwnerGuid", mp_sAssnOwnerGuid)
		End If
		oXML.WriteProperty("IsBudget", mp_bIsBudget)
		If mp_oAvailabilityPeriods.IsNull() = False Then
			oXML.WriteObject(mp_oAvailabilityPeriods.GetXML())
		End If
		If mp_oRates.IsNull() = False Then
			oXML.WriteObject(mp_oRates.GetXML())
		End If
		If mp_oTimephasedData_C.IsNull() = False Then
			mp_oTimephasedData_C.WriteObjectProtected(oXML)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Resource")
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
		oXML.ReadProperty("Initials", mp_sInitials)
		If mp_sInitials.Length > 512 Then
			mp_sInitials = mp_sInitials.Substring(0, 512)
		End If
		oXML.ReadProperty("Phonetics", mp_sPhonetics)
		If mp_sPhonetics.Length > 512 Then
			mp_sPhonetics = mp_sPhonetics.Substring(0, 512)
		End If
		oXML.ReadProperty("NTAccount", mp_sNTAccount)
		If mp_sNTAccount.Length > 512 Then
			mp_sNTAccount = mp_sNTAccount.Substring(0, 512)
		End If
		oXML.ReadProperty("MaterialLabel", mp_sMaterialLabel)
		If mp_sMaterialLabel.Length > 512 Then
			mp_sMaterialLabel = mp_sMaterialLabel.Substring(0, 512)
		End If
		oXML.ReadProperty("Code", mp_sCode)
		If mp_sCode.Length > 512 Then
			mp_sCode = mp_sCode.Substring(0, 512)
		End If
		oXML.ReadProperty("Group", mp_sGroup)
		If mp_sGroup.Length > 512 Then
			mp_sGroup = mp_sGroup.Substring(0, 512)
		End If
		oXML.ReadProperty("WorkGroup", mp_yWorkGroup)
		oXML.ReadProperty("EmailAddress", mp_sEmailAddress)
		If mp_sEmailAddress.Length > 512 Then
			mp_sEmailAddress = mp_sEmailAddress.Substring(0, 512)
		End If
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
		oXML.ReadProperty("MaxUnits", mp_fMaxUnits)
		oXML.ReadProperty("PeakUnits", mp_fPeakUnits)
		oXML.ReadProperty("OverAllocated", mp_bOverAllocated)
		oXML.ReadProperty("AvailableFrom", mp_dtAvailableFrom)
		oXML.ReadProperty("AvailableTo", mp_dtAvailableTo)
		oXML.ReadProperty("Start", mp_dtStart)
		oXML.ReadProperty("Finish", mp_dtFinish)
		oXML.ReadProperty("CanLevel", mp_bCanLevel)
		oXML.ReadProperty("AccrueAt", mp_yAccrueAt)
		oXML.ReadProperty("Work", mp_oWork)
		oXML.ReadProperty("RegularWork", mp_oRegularWork)
		oXML.ReadProperty("OvertimeWork", mp_oOvertimeWork)
		oXML.ReadProperty("ActualWork", mp_oActualWork)
		oXML.ReadProperty("RemainingWork", mp_oRemainingWork)
		oXML.ReadProperty("ActualOvertimeWork", mp_oActualOvertimeWork)
		oXML.ReadProperty("RemainingOvertimeWork", mp_oRemainingOvertimeWork)
		oXML.ReadProperty("PercentWorkComplete", mp_lPercentWorkComplete)
		oXML.ReadProperty("StandardRate", mp_cStandardRate)
		oXML.ReadProperty("StandardRateFormat", mp_yStandardRateFormat)
		oXML.ReadProperty("Cost", mp_cCost)
		oXML.ReadProperty("OvertimeRate", mp_cOvertimeRate)
		oXML.ReadProperty("OvertimeRateFormat", mp_yOvertimeRateFormat)
		oXML.ReadProperty("OvertimeCost", mp_cOvertimeCost)
		oXML.ReadProperty("CostPerUse", mp_cCostPerUse)
		oXML.ReadProperty("ActualCost", mp_cActualCost)
		oXML.ReadProperty("ActualOvertimeCost", mp_cActualOvertimeCost)
		oXML.ReadProperty("RemainingCost", mp_cRemainingCost)
		oXML.ReadProperty("RemainingOvertimeCost", mp_cRemainingOvertimeCost)
		oXML.ReadProperty("WorkVariance", mp_fWorkVariance)
		oXML.ReadProperty("CostVariance", mp_fCostVariance)
		oXML.ReadProperty("SV", mp_fSV)
		oXML.ReadProperty("CV", mp_fCV)
		oXML.ReadProperty("ACWP", mp_fACWP)
		oXML.ReadProperty("CalendarUID", mp_lCalendarUID)
		oXML.ReadProperty("Notes", mp_sNotes)
		oXML.ReadProperty("BCWS", mp_fBCWS)
		oXML.ReadProperty("BCWP", mp_fBCWP)
		oXML.ReadProperty("IsGeneric", mp_bIsGeneric)
		oXML.ReadProperty("IsInactive", mp_bIsInactive)
		oXML.ReadProperty("IsEnterprise", mp_bIsEnterprise)
		oXML.ReadProperty("BookingType", mp_yBookingType)
		oXML.ReadProperty("ActualWorkProtected", mp_oActualWorkProtected)
		oXML.ReadProperty("ActualOvertimeWorkProtected", mp_oActualOvertimeWorkProtected)
		oXML.ReadProperty("ActiveDirectoryGUID", mp_sActiveDirectoryGUID)
		If mp_sActiveDirectoryGUID.Length > 16 Then
			mp_sActiveDirectoryGUID = mp_sActiveDirectoryGUID.Substring(0, 16)
		End If
		oXML.ReadProperty("CreationDate", mp_dtCreationDate)
		mp_oExtendedAttribute_C.ReadObjectProtected(oXML)
		mp_oBaseline_C.ReadObjectProtected(oXML)
		mp_oOutlineCode_C.ReadObjectProtected(oXML)
		oXML.ReadProperty("IsCostResource", mp_bIsCostResource)
		oXML.ReadProperty("AssnOwner", mp_sAssnOwner)
		oXML.ReadProperty("AssnOwnerGuid", mp_sAssnOwnerGuid)
		oXML.ReadProperty("IsBudget", mp_bIsBudget)
		mp_oAvailabilityPeriods.SetXML(oXML.ReadObject("AvailabilityPeriods"))
		mp_oRates.SetXML(oXML.ReadObject("Rates"))
		mp_oTimephasedData_C.ReadObjectProtected(oXML)
	End Sub

End Class
