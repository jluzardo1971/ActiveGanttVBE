Option Explicit On

Public Class OutlineCodeMask
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lLevel As Integer
	Private mp_yType As E_TYPE
	Private mp_lLength As Integer
	Private mp_sSeparator As String

	Public Sub New()
		mp_lLevel = 0
		mp_yType = E_TYPE.T_NUMBERS
		mp_lLength = 0
		mp_sSeparator = ""
	End Sub

	Public Property lLevel() As Integer
		Get
			Return mp_lLevel
		End Get
		Set(ByVal Value As Integer)
			mp_lLevel = Value
		End Set
	End Property

	Public Property yType() As E_TYPE
		Get
			Return mp_yType
		End Get
		Set(ByVal Value As E_TYPE)
			mp_yType = Value
		End Set
	End Property

	Public Property lLength() As Integer
		Get
			Return mp_lLength
		End Get
		Set(ByVal Value As Integer)
			mp_lLength = Value
		End Set
	End Property

	Public Property sSeparator() As String
		Get
			Return mp_sSeparator
		End Get
		Set(ByVal Value As String)
			mp_sSeparator = Value
		End Set
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
		If mp_lLevel <> 0 Then
			bReturn = False
		End If
		If mp_yType <> E_TYPE.T_NUMBERS Then
			bReturn = False
		End If
		If mp_lLength <> 0 Then
			bReturn = False
		End If
		If mp_sSeparator <> "" Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Mask/>"
		End if
		Dim oXML As New clsXML("Mask")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("Level", mp_lLevel)
		oXML.WriteProperty("Type", mp_yType)
		oXML.WriteProperty("Length", mp_lLength)
		If mp_sSeparator <> "" Then
			oXML.WriteProperty("Separator", mp_sSeparator)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Mask")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("Level", mp_lLevel)
		oXML.ReadProperty("Type", mp_yType)
		oXML.ReadProperty("Length", mp_lLength)
		oXML.ReadProperty("Separator", mp_sSeparator)
	End Sub

End Class
