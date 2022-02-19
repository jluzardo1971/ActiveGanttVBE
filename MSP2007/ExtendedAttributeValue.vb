Option Explicit On

Public Class ExtendedAttributeValue
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lID As Integer
	Private mp_sValue As String
	Private mp_sDescription As String
	Private mp_sPhonetic As String

	Public Sub New()
		mp_lID = 0
		mp_sValue = ""
		mp_sDescription = ""
		mp_sPhonetic = ""
	End Sub

	Public Property lID() As Integer
		Get
			Return mp_lID
		End Get
		Set(ByVal Value As Integer)
			mp_lID = Value
		End Set
	End Property

	Public Property sValue() As String
		Get
			Return mp_sValue
		End Get
		Set(ByVal Value As String)
			mp_sValue = Value
		End Set
	End Property

	Public Property sDescription() As String
		Get
			Return mp_sDescription
		End Get
		Set(ByVal Value As String)
			mp_sDescription = Value
		End Set
	End Property

	Public Property sPhonetic() As String
		Get
			Return mp_sPhonetic
		End Get
		Set(ByVal Value As String)
			mp_sPhonetic = Value
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
		If mp_lID <> 0 Then
			bReturn = False
		End If
		If mp_sValue <> "" Then
			bReturn = False
		End If
		If mp_sDescription <> "" Then
			bReturn = False
		End If
		If mp_sPhonetic <> "" Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<Value/>"
		End if
		Dim oXML As New clsXML("Value")
		oXML.InitializeWriter()
		oXML.SupportOptional = True
		oXML.BoolsAreNumeric = True
		oXML.WriteProperty("ID", mp_lID)
		If mp_sValue <> "" Then
			oXML.WriteProperty("Value", mp_sValue)
		End If
		If mp_sDescription <> "" Then
			oXML.WriteProperty("Description", mp_sDescription)
		End If
		If mp_sPhonetic <> "" Then
			oXML.WriteProperty("Phonetic", mp_sPhonetic)
		End If
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("Value")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("ID", mp_lID)
		oXML.ReadProperty("Value", mp_sValue)
		oXML.ReadProperty("Description", mp_sDescription)
		oXML.ReadProperty("Phonetic", mp_sPhonetic)
	End Sub

End Class
