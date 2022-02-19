Option Explicit On

Public Class OutlineCode
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_sFieldID As String
	Private mp_sFieldName As String
	Private mp_sAlias As String
	Private mp_sPhoneticAlias As String
	Private mp_oValues As OutlineCodeValues
	Private mp_bEnterprise As Boolean
	Private mp_lEnterpriseOutlineCodeAlias As Integer
	Private mp_bResourceSubstitutionEnabled As Boolean
	Private mp_bLeafOnly As Boolean
	Private mp_bAllLevelsRequired As Boolean
	Private mp_bOnlyTableValuesAllowed As Boolean
	Private mp_oMasks As OutlineCodeMasks

	Public Sub New()
		mp_sFieldID = ""
		mp_sFieldName = ""
		mp_sAlias = ""
		mp_sPhoneticAlias = ""
		mp_oValues = New OutlineCodeValues()
		mp_bEnterprise = False
		mp_lEnterpriseOutlineCodeAlias = 0
		mp_bResourceSubstitutionEnabled = False
		mp_bLeafOnly = False
		mp_bAllLevelsRequired = False
		mp_bOnlyTableValuesAllowed = False
		mp_oMasks = New OutlineCodeMasks()
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

	Public Property sAlias() As String
		Get
			Return mp_sAlias
		End Get
		Set(ByVal Value As String)
			mp_sAlias = Value
		End Set
	End Property

	Public Property sPhoneticAlias() As String
		Get
			Return mp_sPhoneticAlias
		End Get
		Set(ByVal Value As String)
			mp_sPhoneticAlias = Value
		End Set
	End Property

	Public ReadOnly Property oValues() As OutlineCodeValues
		Get
			Return mp_oValues
		End Get
	End Property

	Public Property bEnterprise() As Boolean
		Get
			Return mp_bEnterprise
		End Get
		Set(ByVal Value As Boolean)
			mp_bEnterprise = Value
		End Set
	End Property

	Public Property lEnterpriseOutlineCodeAlias() As Integer
		Get
			Return mp_lEnterpriseOutlineCodeAlias
		End Get
		Set(ByVal Value As Integer)
			mp_lEnterpriseOutlineCodeAlias = Value
		End Set
	End Property

	Public Property bResourceSubstitutionEnabled() As Boolean
		Get
			Return mp_bResourceSubstitutionEnabled
		End Get
		Set(ByVal Value As Boolean)
			mp_bResourceSubstitutionEnabled = Value
		End Set
	End Property

	Public Property bLeafOnly() As Boolean
		Get
			Return mp_bLeafOnly
		End Get
		Set(ByVal Value As Boolean)
			mp_bLeafOnly = Value
		End Set
	End Property

	Public Property bAllLevelsRequired() As Boolean
		Get
			Return mp_bAllLevelsRequired
		End Get
		Set(ByVal Value As Boolean)
			mp_bAllLevelsRequired = Value
		End Set
	End Property

	Public Property bOnlyTableValuesAllowed() As Boolean
		Get
			Return mp_bOnlyTableValuesAllowed
		End Get
		Set(ByVal Value As Boolean)
			mp_bOnlyTableValuesAllowed = Value
		End Set
	End Property

	Public ReadOnly Property oMasks() As OutlineCodeMasks
		Get
			Return mp_oMasks
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
		If mp_sAlias <> "" Then
			bReturn = False
		End If
		If mp_sPhoneticAlias <> "" Then
			bReturn = False
		End If
		If mp_oValues.IsNull() = False Then
			bReturn = False
		End If
		If mp_bEnterprise <> False Then
			bReturn = False
		End If
		If mp_lEnterpriseOutlineCodeAlias <> 0 Then
			bReturn = False
		End If
		If mp_bResourceSubstitutionEnabled <> False Then
			bReturn = False
		End If
		If mp_bLeafOnly <> False Then
			bReturn = False
		End If
		If mp_bAllLevelsRequired <> False Then
			bReturn = False
		End If
		If mp_bOnlyTableValuesAllowed <> False Then
			bReturn = False
		End If
		If mp_oMasks.IsNull() = False Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<OutlineCode/>"
		End if
		Dim oXML As New clsXML("OutlineCode")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		If mp_sFieldID <> "" Then
			oXML.WriteProperty("FieldID", mp_sFieldID)
		End If
		If mp_sFieldName <> "" Then
			oXML.WriteProperty("FieldName", mp_sFieldName)
		End If
		If mp_sAlias <> "" Then
			oXML.WriteProperty("Alias", mp_sAlias)
		End If
		If mp_sPhoneticAlias <> "" Then
			oXML.WriteProperty("PhoneticAlias", mp_sPhoneticAlias)
		End If
		If mp_oValues.IsNull() = False Then
			oXML.WriteObject(mp_oValues.GetXML())
		End If
		oXML.WriteProperty("Enterprise", mp_bEnterprise)
		oXML.WriteProperty("EnterpriseOutlineCodeAlias", mp_lEnterpriseOutlineCodeAlias)
		oXML.WriteProperty("ResourceSubstitutionEnabled", mp_bResourceSubstitutionEnabled)
		oXML.WriteProperty("LeafOnly", mp_bLeafOnly)
		oXML.WriteProperty("AllLevelsRequired", mp_bAllLevelsRequired)
		oXML.WriteProperty("OnlyTableValuesAllowed", mp_bOnlyTableValuesAllowed)
		If mp_oMasks.IsNull() = False Then
			oXML.WriteObject(mp_oMasks.GetXML())
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("OutlineCode")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FieldID", mp_sFieldID)
		oXML.ReadProperty("FieldName", mp_sFieldName)
		oXML.ReadProperty("Alias", mp_sAlias)
		oXML.ReadProperty("PhoneticAlias", mp_sPhoneticAlias)
		mp_oValues.SetXML(oXML.ReadObject("Values"))
		oXML.ReadProperty("Enterprise", mp_bEnterprise)
		oXML.ReadProperty("EnterpriseOutlineCodeAlias", mp_lEnterpriseOutlineCodeAlias)
		oXML.ReadProperty("ResourceSubstitutionEnabled", mp_bResourceSubstitutionEnabled)
		oXML.ReadProperty("LeafOnly", mp_bLeafOnly)
		oXML.ReadProperty("AllLevelsRequired", mp_bAllLevelsRequired)
		oXML.ReadProperty("OnlyTableValuesAllowed", mp_bOnlyTableValuesAllowed)
		mp_oMasks.SetXML(oXML.ReadObject("Masks"))
	End Sub

End Class
