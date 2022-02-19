Option Explicit On

Public Class ResourceExtendedAttribute
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_lUID As Integer
	Private mp_sFieldID As String
	Private mp_sValue As String
	Private mp_lValueID As Integer
	Private mp_yDurationFormat As E_DURATIONFORMAT

	Public Sub New()
		mp_lUID = 0
		mp_sFieldID = ""
		mp_sValue = ""
		mp_lValueID = 0
		mp_yDurationFormat = E_DURATIONFORMAT.DF_M
	End Sub

	Public Property lUID() As Integer
		Get
			Return mp_lUID
		End Get
		Set(ByVal Value As Integer)
			mp_lUID = Value
		End Set
	End Property

	Public Property sFieldID() As String
		Get
			Return mp_sFieldID
		End Get
		Set(ByVal Value As String)
			mp_sFieldID = Value
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

	Public Property lValueID() As Integer
		Get
			Return mp_lValueID
		End Get
		Set(ByVal Value As Integer)
			mp_lValueID = Value
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
		If mp_sFieldID <> "" Then
			bReturn = False
		End If
		If mp_sValue <> "" Then
			bReturn = False
		End If
		If mp_lValueID <> 0 Then
			bReturn = False
		End If
		If mp_yDurationFormat <> E_DURATIONFORMAT.DF_M Then
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
		oXML.WriteProperty("UID", mp_lUID)
		If mp_sFieldID <> "" Then
			oXML.WriteProperty("FieldID", mp_sFieldID)
		End If
		If mp_sValue <> "" Then
			oXML.WriteProperty("Value", mp_sValue)
		End If
		oXML.WriteProperty("ValueID", mp_lValueID)
		oXML.WriteProperty("DurationFormat", mp_yDurationFormat)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("ExtendedAttribute")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("UID", mp_lUID)
		oXML.ReadProperty("FieldID", mp_sFieldID)
		oXML.ReadProperty("Value", mp_sValue)
		oXML.ReadProperty("ValueID", mp_lValueID)
		oXML.ReadProperty("DurationFormat", mp_yDurationFormat)
	End Sub

End Class
