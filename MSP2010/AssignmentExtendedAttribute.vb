Option Explicit On

Public Class AssignmentExtendedAttribute
	Inherits clsItemBase


	Friend mp_oCollection As clsCollectionBase
	Private mp_sFieldID As String
	Private mp_sValue As String
	Private mp_lValueGUID As Integer
	Private mp_yDurationFormat As E_DURATIONFORMAT

	Public Sub New()
		mp_sFieldID = ""
		mp_sValue = ""
		mp_lValueGUID = 0
		mp_yDurationFormat = E_DURATIONFORMAT.DF_M
	End Sub

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

	Public Property lValueGUID() As Integer
		Get
			Return mp_lValueGUID
		End Get
		Set(ByVal Value As Integer)
			mp_lValueGUID = Value
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
		If mp_sFieldID <> "" Then
			bReturn = False
		End If
		If mp_sValue <> "" Then
			bReturn = False
		End If
		If mp_lValueGUID <> 0 Then
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
		If mp_sFieldID <> "" Then
			oXML.WriteProperty("FieldID", mp_sFieldID)
		End If
		If mp_sValue <> "" Then
			oXML.WriteProperty("Value", mp_sValue)
		End If
		oXML.WriteProperty("ValueGUID", mp_lValueGUID)
		oXML.WriteProperty("DurationFormat", mp_yDurationFormat)
		Return oXML.GetXML()
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim oXML As New clsXML("ExtendedAttribute")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		oXML.ReadProperty("FieldID", mp_sFieldID)
		oXML.ReadProperty("Value", mp_sValue)
		oXML.ReadProperty("ValueGUID", mp_lValueGUID)
		oXML.ReadProperty("DurationFormat", mp_yDurationFormat)
	End Sub

End Class
