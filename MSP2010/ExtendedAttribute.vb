Option Explicit On

Public Class ExtendedAttribute
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_sFieldID As String
	Private mp_sFieldName As String
	Private mp_yCFType As E_CFTYPE
	Private mp_sGuid As String
	Private mp_yElemType As E_ELEMTYPE
	Private mp_lMaxMultiValues As Integer
	Private mp_bUserDef As Boolean
	Private mp_sAlias As String
	Private mp_sSecondaryPID As String
	Private mp_bAutoRollDown As Boolean
	Private mp_sDefaultGuid As String
	Private mp_sLtuid As String
	Private mp_sSecondaryGuid As String
	Private mp_sPhoneticAlias As String
	Private mp_yRollupType As E_ROLLUPTYPE
	Private mp_yCalculationType As E_CALCULATIONTYPE
	Private mp_sFormula As String
	Private mp_bRestrictValues As Boolean
	Private mp_yValuelistSortOrder As E_VALUELISTSORTORDER
	Private mp_bAppendNewValues As Boolean
	Private mp_sDefault As String
	Private mp_oValueList As ExtendedAttributeValueList

	Public Sub New()
		mp_sFieldID = ""
		mp_sFieldName = ""
		mp_yCFType = E_CFTYPE.CFT_COST
		mp_sGuid = ""
		mp_yElemType = E_ELEMTYPE.ET_TASK
		mp_lMaxMultiValues = 0
		mp_bUserDef = False
		mp_sAlias = ""
		mp_sSecondaryPID = ""
		mp_bAutoRollDown = False
		mp_sDefaultGuid = ""
		mp_sLtuid = ""
		mp_sSecondaryGuid = ""
		mp_sPhoneticAlias = ""
		mp_yRollupType = E_ROLLUPTYPE.RT_MAXIMUM_OR_FOR_FLAG_FIELDS
		mp_yCalculationType = E_CALCULATIONTYPE.CT_NONE
		mp_sFormula = ""
		mp_bRestrictValues = False
		mp_yValuelistSortOrder = E_VALUELISTSORTORDER.VSO_DESCENDING
		mp_bAppendNewValues = False
		mp_sDefault = ""
		mp_oValueList = New ExtendedAttributeValueList()
	End Sub

	Public Property sFieldID() As String
		Get
			Return mp_sFieldID
		End Get
		Set(ByVal Value As String)
			mp_sFieldID = Value
		End Set
	End Property

	Public Property sFieldName() As String
		Get
			Return mp_sFieldName
		End Get
		Set(ByVal Value As String)
			mp_sFieldName = Value
		End Set
	End Property

	Public Property yCFType() As E_CFTYPE
		Get
			Return mp_yCFType
		End Get
		Set(ByVal Value As E_CFTYPE)
			mp_yCFType = Value
		End Set
	End Property

	Public Property sGuid() As String
		Get
			Return mp_sGuid
		End Get
		Set(ByVal Value As String)
			mp_sGuid = Value
		End Set
	End Property

	Public Property yElemType() As E_ELEMTYPE
		Get
			Return mp_yElemType
		End Get
		Set(ByVal Value As E_ELEMTYPE)
			mp_yElemType = Value
		End Set
	End Property

	Public Property lMaxMultiValues() As Integer
		Get
			Return mp_lMaxMultiValues
		End Get
		Set(ByVal Value As Integer)
			mp_lMaxMultiValues = Value
		End Set
	End Property

	Public Property bUserDef() As Boolean
		Get
			Return mp_bUserDef
		End Get
		Set(ByVal Value As Boolean)
			mp_bUserDef = Value
		End Set
	End Property

	Public Property sAlias() As String
		Get
			Return mp_sAlias
		End Get
		Set(ByVal Value As String)
			If Value.Length > 50 Then
				Value = Value.Substring(0, 50)
			End If
			mp_sAlias = Value
		End Set
	End Property

	Public Property sSecondaryPID() As String
		Get
			Return mp_sSecondaryPID
		End Get
		Set(ByVal Value As String)
			mp_sSecondaryPID = Value
		End Set
	End Property

	Public Property bAutoRollDown() As Boolean
		Get
			Return mp_bAutoRollDown
		End Get
		Set(ByVal Value As Boolean)
			mp_bAutoRollDown = Value
		End Set
	End Property

	Public Property sDefaultGuid() As String
		Get
			Return mp_sDefaultGuid
		End Get
		Set(ByVal Value As String)
			mp_sDefaultGuid = Value
		End Set
	End Property

	Public Property sLtuid() As String
		Get
			Return mp_sLtuid
		End Get
		Set(ByVal Value As String)
			mp_sLtuid = Value
		End Set
	End Property

	Public Property sSecondaryGuid() As String
		Get
			Return mp_sSecondaryGuid
		End Get
		Set(ByVal Value As String)
			mp_sSecondaryGuid = Value
		End Set
	End Property

	Public Property sPhoneticAlias() As String
		Get
			Return mp_sPhoneticAlias
		End Get
		Set(ByVal Value As String)
			If Value.Length > 50 Then
				Value = Value.Substring(0, 50)
			End If
			mp_sPhoneticAlias = Value
		End Set
	End Property

	Public Property yRollupType() As E_ROLLUPTYPE
		Get
			Return mp_yRollupType
		End Get
		Set(ByVal Value As E_ROLLUPTYPE)
			mp_yRollupType = Value
		End Set
	End Property

	Public Property yCalculationType() As E_CALCULATIONTYPE
		Get
			Return mp_yCalculationType
		End Get
		Set(ByVal Value As E_CALCULATIONTYPE)
			mp_yCalculationType = Value
		End Set
	End Property

	Public Property sFormula() As String
		Get
			Return mp_sFormula
		End Get
		Set(ByVal Value As String)
			mp_sFormula = Value
		End Set
	End Property

	Public Property bRestrictValues() As Boolean
		Get
			Return mp_bRestrictValues
		End Get
		Set(ByVal Value As Boolean)
			mp_bRestrictValues = Value
		End Set
	End Property

	Public Property yValuelistSortOrder() As E_VALUELISTSORTORDER
		Get
			Return mp_yValuelistSortOrder
		End Get
		Set(ByVal Value As E_VALUELISTSORTORDER)
			mp_yValuelistSortOrder = Value
		End Set
	End Property

	Public Property bAppendNewValues() As Boolean
		Get
			Return mp_bAppendNewValues
		End Get
		Set(ByVal Value As Boolean)
			mp_bAppendNewValues = Value
		End Set
	End Property

	Public Property sDefault() As String
		Get
			Return mp_sDefault
		End Get
		Set(ByVal Value As String)
			mp_sDefault = Value
		End Set
	End Property

	Public ReadOnly Property oValueList() As ExtendedAttributeValueList
		Get
			Return mp_oValueList
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
		If mp_sFieldID <> "" Then
			bReturn = False
		End If
		If mp_sFieldName <> "" Then
			bReturn = False
		End If
		If mp_yCFType <> E_CFTYPE.CFT_COST Then
			bReturn = False
		End If
		If mp_sGuid <> "" Then
			bReturn = False
		End If
		If mp_yElemType <> E_ELEMTYPE.ET_TASK Then
			bReturn = False
		End If
		If mp_lMaxMultiValues <> 0 Then
			bReturn = False
		End If
		If mp_bUserDef <> False Then
			bReturn = False
		End If
		If mp_sAlias <> "" Then
			bReturn = False
		End If
		If mp_sSecondaryPID <> "" Then
			bReturn = False
		End If
		If mp_bAutoRollDown <> False Then
			bReturn = False
		End If
		If mp_sDefaultGuid <> "" Then
			bReturn = False
		End If
		If mp_sLtuid <> "" Then
			bReturn = False
		End If
		If mp_sSecondaryGuid <> "" Then
			bReturn = False
		End If
		If mp_sPhoneticAlias <> "" Then
			bReturn = False
		End If
		If mp_yRollupType <> E_ROLLUPTYPE.RT_MAXIMUM_OR_FOR_FLAG_FIELDS Then
			bReturn = False
		End If
		If mp_yCalculationType <> E_CALCULATIONTYPE.CT_NONE Then
			bReturn = False
		End If
		If mp_sFormula <> "" Then
			bReturn = False
		End If
		If mp_bRestrictValues <> False Then
			bReturn = False
		End If
		If mp_yValuelistSortOrder <> E_VALUELISTSORTORDER.VSO_DESCENDING Then
			bReturn = False
		End If
		If mp_bAppendNewValues <> False Then
			bReturn = False
		End If
		If mp_sDefault <> "" Then
			bReturn = False
		End If
		If mp_oValueList.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<ExtendedAttribute/>"
		End if
		Dim oXML As New clsXML("ExtendedAttribute")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_sFieldID <> "" Then
			oXML.WriteProperty("FieldID", mp_sFieldID)
		End If
		If mp_sFieldName <> "" Then
			oXML.WriteProperty("FieldName", mp_sFieldName)
		End If
		oXML.WriteProperty("CFType", mp_yCFType)
		If mp_sGuid <> "" Then
			oXML.WriteProperty("Guid", mp_sGuid)
		End If
		oXML.WriteProperty("ElemType", mp_yElemType)
		oXML.WriteProperty("MaxMultiValues", mp_lMaxMultiValues)
		oXML.WriteProperty("UserDef", mp_bUserDef)
		If mp_sAlias <> "" Then
			oXML.WriteProperty("Alias", mp_sAlias)
		End If
		If mp_sSecondaryPID <> "" Then
			oXML.WriteProperty("SecondaryPID", mp_sSecondaryPID)
		End If
		oXML.WriteProperty("AutoRollDown", mp_bAutoRollDown)
		If mp_sDefaultGuid <> "" Then
			oXML.WriteProperty("DefaultGuid", mp_sDefaultGuid)
		End If
		If mp_sLtuid <> "" Then
			oXML.WriteProperty("Ltuid", mp_sLtuid)
		End If
		If mp_sSecondaryGuid <> "" Then
			oXML.WriteProperty("SecondaryGuid", mp_sSecondaryGuid)
		End If
		If mp_sPhoneticAlias <> "" Then
			oXML.WriteProperty("PhoneticAlias", mp_sPhoneticAlias)
		End If
		oXML.WriteProperty("RollupType", mp_yRollupType)
		oXML.WriteProperty("CalculationType", mp_yCalculationType)
		If mp_sFormula <> "" Then
			oXML.WriteProperty("Formula", mp_sFormula)
		End If
		oXML.WriteProperty("RestrictValues", mp_bRestrictValues)
		oXML.WriteProperty("ValuelistSortOrder", mp_yValuelistSortOrder)
		oXML.WriteProperty("AppendNewValues", mp_bAppendNewValues)
		If mp_sDefault <> "" Then
			oXML.WriteProperty("Default", mp_sDefault)
		End If
		If mp_oValueList.IsNull() = False Then
			oXML.WriteObject(mp_oValueList.GetXML())
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("ExtendedAttribute")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FieldID", mp_sFieldID)
		oXML.ReadProperty("FieldName", mp_sFieldName)
		oXML.ReadProperty("CFType", mp_yCFType)
		oXML.ReadProperty("Guid", mp_sGuid)
		oXML.ReadProperty("ElemType", mp_yElemType)
		oXML.ReadProperty("MaxMultiValues", mp_lMaxMultiValues)
		oXML.ReadProperty("UserDef", mp_bUserDef)
		oXML.ReadProperty("Alias", mp_sAlias)
		If mp_sAlias.Length > 50 Then
			mp_sAlias = mp_sAlias.Substring(0, 50)
		End If
		oXML.ReadProperty("SecondaryPID", mp_sSecondaryPID)
		oXML.ReadProperty("AutoRollDown", mp_bAutoRollDown)
		oXML.ReadProperty("DefaultGuid", mp_sDefaultGuid)
		oXML.ReadProperty("Ltuid", mp_sLtuid)
		oXML.ReadProperty("SecondaryGuid", mp_sSecondaryGuid)
		oXML.ReadProperty("PhoneticAlias", mp_sPhoneticAlias)
		If mp_sPhoneticAlias.Length > 50 Then
			mp_sPhoneticAlias = mp_sPhoneticAlias.Substring(0, 50)
		End If
		oXML.ReadProperty("RollupType", mp_yRollupType)
		oXML.ReadProperty("CalculationType", mp_yCalculationType)
		oXML.ReadProperty("Formula", mp_sFormula)
		oXML.ReadProperty("RestrictValues", mp_bRestrictValues)
		oXML.ReadProperty("ValuelistSortOrder", mp_yValuelistSortOrder)
		oXML.ReadProperty("AppendNewValues", mp_bAppendNewValues)
		oXML.ReadProperty("Default", mp_sDefault)
		mp_oValueList.SetXML(oXML.ReadObject("ValueList"))
	End Sub

End Class
